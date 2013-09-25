import java.io.IOException;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.w3c.dom.Element;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Scanner;
import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.StringReader;
import java.io.UnsupportedEncodingException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.sax.SAXSource;
import javax.xml.transform.sax.SAXTransformerFactory;
import javax.xml.transform.stream.StreamResult;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.auth.AuthScope;
import org.apache.http.auth.NTCredentials;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.io.FileUtils;
import org.xml.sax.InputSource;

public class MainApp 
{
  public static void main(String[] argv) 
  {
    try
    {
      // NOTE:  for most of these tests, you'll need to edit the associated XML file , and change parms to match your account.
	  // e.g. for the "qs" example, edit the "qs.xml" file and change it as you see fit.
      // You'll also need to edit the exchange credentials in the "test" method.
    
      // A search
      test("qs");
      
      // Get folder info
      test("get_folder");
      
      // Get out of office settings
      test("ooo");
      
      /*
      // Find a list of all folders
      test("find_folders");
      
      // Get info about particular items
      test("get_items");
      
      // Get message body for particular items
      getMessageBody();
      
      // Get info about a folder
      getFolderIdByName("Sent Items");
      */
    }
    catch(Exception e)
    {

    }
  }
  public static String test(String testName)
  {
    try
    {
      int serverPort = 443;
      String serverProtocol = "https";
      String workstation = "";
      String domain = "";

      String username = "yourusername@test.com";
      String password = "YourPassword";
      String server = "connect.emailsrvr.com";
      String requestAbsoluteUrl = "https://connect.emailsrvr.com/EWS/Exchange.asmx";
      String requestData;

      requestData = getFile(testName);

      String result = null;
      try
      {
        result = sendRequestUsingBasicAuthentication(testName, username, password, workstation, domain, server, serverPort, serverProtocol, requestAbsoluteUrl, requestData);
      }
      catch (Throwable e)
      {
        result = null;
      }

      if (result == null)
      {
        result = sendRequestUsingNtlmAuthentication(testName, username, password, workstation, domain, server, serverPort, serverProtocol, requestAbsoluteUrl, requestData);
      }
      
      return result;
    }
    catch (Throwable e)
    {
      return null;
    }
  }

  public static long getDateFromStr(String target)
  {
    try 
    {
      String f = "yyyy-MM-dd'T'HH:mm:ss'Z'";

      DateFormat df = new SimpleDateFormat(f);
      Date date = df.parse(target);

      return date.getTime();
    } 
    catch (ParseException e) 
    {
      return -1;
    }  
  }

  public static String getElementEmailSummaryFromElement(Element ndParent)
  {
    String result = "";

    NodeList items = ndParent.getElementsByTagName("t:Mailbox");

    for (int i=0; i < items.getLength(); i++)
    {
      Element ndItem = (Element) items.item(i);
      String from = getElementTextFromElement(ndItem, "t:Name");
      String fromEmail = getElementTextFromElement(ndItem, "t:EmailAddress");

      result += String.format("%s <%s>", from, fromEmail);

      if (i != items.getLength() - 1)
      {
        result += ", ";
      }
    }
    return result;
  }

  public static String getFolderIdByName(String name)
  {
    String result = "";

    try
    {
      String xml = test("find_folders");
      NodeList items = getItemsFromXml(xml, "t:Folder");

      for (int i=0; i < items.getLength(); i++)
      {
        Element ndItem = (Element) items.item(i);
        String displayName = getElementTextFromElement(ndItem, "t:DisplayName");

        if (name.equals(displayName))
        {
          result = getElementAttributeFromElement(ndItem, "t:FolderId", "Id");
          break;
        }
      }

      System.out.println("getFolderIdByName:" + name + ", id:" + result);
    }
    catch(Exception e)
    {

    }

    return result;
  }

  public static String getMessageBody()
  {
    try
    {
      String xml = test("getitem");
      Element ndMessage = getElementFromXml(xml, "t:Message");

      Element ndFrom = getElementFromElement(ndMessage, "t:From");
      String data = getElementEmailSummaryFromElement(ndFrom);

      System.out.println("from:" + data);

      String result = "";

      return result;
    }
    catch(Exception e)
    {
      return null;
    }
  }

  public static String getElementAttributeFromElement(Element e, String xmlElementName, String attributeName)
  {
    Element child = getElementFromElement(e, xmlElementName);

    if (child == null)
      return "";

    return child.getAttribute(attributeName);
  }

  public static String getElementTextFromElement(Element e, String xmlElementName)
  {
    Element child = getElementFromElement(e, xmlElementName);

    if (child == null)
      return "";

    return child.getTextContent();
  }

  public static Element getElementFromElement(Element e, String xmlElementName)
  {
    try
    {
      Element result = (Element) e.getElementsByTagName(xmlElementName).item(0);
      return result;
    }
    catch (Exception ex)
    {
      return null;
    }
  }

  public static Element getElementFromXml(String xml, String xmlElementName)
  {
    NodeList items = getItemsFromXml(xml, xmlElementName);

    if (items == null || items.getLength() == 0)
      return null;

    Element result = (Element) items.item(0);

    return result;
  }

  public static NodeList getItemsFromXml(String xml, String xmlElementName)
  {
    try 
    {
      DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
      DocumentBuilder dBuilder;
      dBuilder = dbFactory.newDocumentBuilder();
      InputSource is = new InputSource(new StringReader(xml));
      Document doc = dBuilder.parse(is);

      //optional, but recommended
      //read this - http://stackoverflow.com/questions/13786607/normalization-in-dom-parsing-with-java-how-does-it-work
      doc.getDocumentElement().normalize();

      NodeList nList = doc.getElementsByTagName(xmlElementName);

      return nList;
    } 
    catch (Exception e) 
    {
      return null;
    }
  }


  public static String simpleTest() throws Exception
  {
    SimpleEwsRequestProcessor rp = new SimpleEwsRequestProcessor();

    String request;

    //request = "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"      xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"> <soap:Header> <t:RequestServerVersion Version=\"Exchange2010\" /> </soap:Header> <soap:Body> <m:FindItem Traversal=\"Shallow\"> <m:ItemShape> <t:BaseShape>IdOnly</t:BaseShape> <t:AdditionalProperties> <t:FieldURI FieldURI=\"item:Subject\" /> <t:FieldURI FieldURI=\"item:DateTimeReceived\" /> </t:AdditionalProperties> </m:ItemShape> <m:IndexedPageItemView MaxEntriesReturned=\"50\" Offset=\"0\" BasePoint=\"Beginning\" /> <m:Restriction> <t:And> <t:Contains ContainmentMode=\"Substring\" ContainmentComparison=\"IgnoreCase\"> <t:FieldURI FieldURI=\"item:Subject\" /> <t:Constant Value=\"Test\" /> </t:Contains> <t:Contains ContainmentMode=\"Substring\" ContainmentComparison=\"IgnoreCase\"> <t:FieldURI FieldURI=\"item:Body\" /> <t:Constant Value=\"Brad\" /> </t:Contains> </t:And> </m:Restriction> <m:SortOrder> <t:FieldOrder Order=\"Descending\"> <t:FieldURI FieldURI=\"item:DateTimeReceived\" /> </t:FieldOrder> </m:SortOrder> <m:ParentFolderIds> <t:DistinguishedFolderId Id=\"inbox\" /> </m:ParentFolderIds> </m:FindItem> </soap:Body> </soap:Envelope>";
    request = getFile("qs");

    String result;

    {
      String subject = "Test";
      String from = "brad";
      String body = "";
      String date = "06/01/2013";
      String folder = "";

      result = rp.processRequest(request, subject, from, body, date, folder);
    }

    System.out.println(result);

    DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
    DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
    InputSource is = new InputSource(new StringReader(result));
    Document doc = dBuilder.parse(is);

    //optional, but recommended
    //read this - http://stackoverflow.com/questions/13786607/normalization-in-dom-parsing-with-java-how-does-it-work
    doc.getDocumentElement().normalize();

    System.out.println("Root element :" + doc.getDocumentElement().getNodeName());

    NodeList nList = doc.getElementsByTagName("t:Message");
    for (int s=0; s < nList.getLength() ; s++)
    {
      Element ndMessage = (Element) nList.item(s);

      Element ndItemId = (Element) ndMessage.getElementsByTagName("t:ItemId").item(0);
      Element ndSubject = (Element) ndMessage.getElementsByTagName("t:Subject").item(0);
      Element ndDateTimeReceived = (Element) ndMessage.getElementsByTagName("t:DateTimeReceived").item(0);
      Element ndInternetMessageId = (Element) ndMessage.getElementsByTagName("t:InternetMessageId").item(0);

      String itemId = ndItemId.getAttribute("Id");
      String subject = ndSubject.getTextContent();
      String dateTimeReceived = ndDateTimeReceived.getTextContent();
      String internetMessageId = ndInternetMessageId.getTextContent();

      String msg = String.format("%s, %s, %s, %s", subject, dateTimeReceived, internetMessageId, itemId);

      System.out.println(msg);
    }
    int count = nList.getLength();
    System.out.println("count:" + count);

    return result;
  }

  private static String appFile(String filename, Object... args)
  {

	You'll need to change this path to the path where the app is ;-)
	  
    String result = "/Users/Brad/ws5/TestNTLM/" + filename + ".xml";
    return result;
  }

  private static String getFile(String filename, Object... args)
  {
    String absoluteFilename = appFile(filename, args);
    String result = readFile(absoluteFilename);

    return result;
  }

  private static void setFile(String filename, String data, Object... args)
  {
    String absoluteFilename = appFile(filename, args);
    try {
      FileUtils.writeStringToFile(new File(absoluteFilename), data);
    } catch (IOException e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
  }

  private static String readFile(String pathname) 
  {
    try
    {
      File file = new File(pathname);
      StringBuilder fileContents = new StringBuilder((int)file.length());
      Scanner scanner = new Scanner(file);
      String lineSeparator = System.getProperty("line.separator");

      try {
        while(scanner.hasNextLine()) {        
          fileContents.append(scanner.nextLine() + lineSeparator);
        }
        return fileContents.toString();
      } finally {
        scanner.close();
      }
    }
    catch (Exception e)
    {
      return null;
    }
  }

  private static String convertStreamToString(java.io.InputStream is) 
  {
    java.util.Scanner s = new java.util.Scanner(is).useDelimiter("\\A");
    return s.hasNext() ? s.next() : "";
  }

  public static String formatXml(String xml)
  {
    try{
      Transformer serializer= SAXTransformerFactory.newInstance().newTransformer();
      serializer.setOutputProperty(OutputKeys.INDENT, "yes");
      serializer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
      Source xmlSource=new SAXSource(new InputSource(new ByteArrayInputStream(xml.getBytes())));
      StreamResult res =  new StreamResult(new ByteArrayOutputStream());            
      serializer.transform(xmlSource, res);
      return new String(((ByteArrayOutputStream)res.getOutputStream()).toByteArray());
    }catch(Exception e){
      return xml;
    }
  }

  public static String sendRequestUsingNtlmAuthentication(String cmd, String username, String password, String workstation, String domain, String server, int serverPort, String serverProtocol, String requestAbsoluteUrl, String requestData) throws Throwable
  {
    DefaultHttpClient httpClient = new DefaultHttpClient();

    NTCredentials creds = new NTCredentials(username, password, workstation, domain);
    httpClient.getCredentialsProvider().setCredentials(AuthScope.ANY, creds);

    System.out.println("");
    System.out.println(cmd);
    String result = sendRequest(httpClient, requestData, requestAbsoluteUrl, null);

    if (result != null)
    {
      result = formatXml(result);

      System.out.println("Result formatted");
      System.out.println(result);
      setFile(cmd + ".out", result);
    }

    return result;
  }

  public static String runShellCmd(String cmd) 
  {
    try
    {
      String result = "";

      Runtime rt = Runtime.getRuntime();

      Process proc = rt.exec(cmd);
      String s;

      // read the output from the command
      BufferedReader stdInput = new BufferedReader(new InputStreamReader(proc.getInputStream()));
      while ((s = stdInput.readLine()) != null) {
        result += s;
      }

      // read any errors from the attempted command
      BufferedReader stdError = new BufferedReader(new InputStreamReader(proc.getErrorStream()));
      while ((s = stdError.readLine()) != null) {
        result += s;
      }

      return result;
    }
    catch(Exception e)
    {
      return null;
    }
  }


  public static String sendRequestUsingBasicAuthentication(String cmd, String username, String password, String workstation, String domain, String server, int serverPort, String serverProtocol, String requestAbsoluteUrl, String requestData) throws Throwable
  {
    DefaultHttpClient httpClient = new DefaultHttpClient();
    String authHeaderValue = "Basic " + Base64.encodeBase64String((username + ":" + password).getBytes());

    // Uncomment and Use the following line instead, if targetting android
    // String authHeaderValue = "Basic " + org.kobjects.base64.Base64.encode((username + ":" + password).getBytes());

    System.out.println("");
    System.out.println(cmd);
    
    return sendRequest(httpClient, requestData, requestAbsoluteUrl, authHeaderValue);
  }

  private static String sendRequest(DefaultHttpClient httpClient, String requestContent, String serviceUrl, String authHeaderValue) throws Throwable 
  {
    // initialize HTTP post
    HttpPost httpPost = null;
    try {
      httpPost = new HttpPost(serviceUrl);
      httpPost.addHeader("Content-Type", "text/xml");
      if (authHeaderValue != null)
      {
        httpPost.addHeader("Authorization", authHeaderValue);
      }
    } catch (Throwable e) {
      throw e;
    }

    // load content to be sent
    try {
      HttpEntity postEntity = new StringEntity(requestContent);
      httpPost.setEntity(postEntity);
    } catch (UnsupportedEncodingException e) {
      throw e;
    }

    // send request
    HttpResponse httpResponse = null;
    try {
      httpResponse = httpClient.execute(httpPost);
    } catch (Throwable e) {
      throw e;
    }

    // get SOAP response
    try {
      // get response code
      int responseStatusCode = httpResponse.getStatusLine().getStatusCode();

      // if the response code is not 200 - OK, or 500 - Internal error,
      // then communication error occurred
      if (responseStatusCode != 200 && responseStatusCode != 500) {
        String errorMsg = "Got SOAP response code " + responseStatusCode + " "
          + httpResponse.getStatusLine().getReasonPhrase();

        System.out.println(errorMsg);
      }

      // get the response content
      HttpEntity httpEntity = httpResponse.getEntity();
      InputStream result = httpEntity.getContent();

      // Dump the response so we can see what's coming back.
      String resultAsString = convertStreamToString(result);
      System.out.println(resultAsString);

      return resultAsString;
    } catch (Throwable e) {
      throw e;
    }
  }
}
