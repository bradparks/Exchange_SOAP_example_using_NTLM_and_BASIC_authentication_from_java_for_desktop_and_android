import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
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
import org.xml.sax.InputSource;

public class SimpleEwsRequestProcessor {
	
	public boolean isDefined(String value)
	{
		if (value == null)
			return false;
		if (value.trim().length() == 0)
			return false;
		return true;
	}
	
	private String addToQueryStringIfDefined(String key, String value)
	{
        return isDefined(value) ? key + ":" + value + " ": "";
	}
	
	public String processRequest(String requestData, String subject, String from, String body, String date, String folder)
	{
        //<m:QueryString>Sent:06/01/2013..06/30/2013 subject:Test body:Brad from:super@lashpoint.com</m:QueryString>

        String searchFor = "<m:QueryString></m:QueryString>";
        
        String queryString = "";
        
        queryString += addToQueryStringIfDefined("subject", subject);
        queryString += addToQueryStringIfDefined("from", from);
        queryString += addToQueryStringIfDefined("body", body);
        if (isDefined(date))
        {
	        queryString += addToQueryStringIfDefined("Sent", date + "..Today");
        }
	        
        String replaceWith = String.format("<m:QueryString>%s</m:QueryString>", queryString.trim());
        
        String request = requestData.replaceAll(searchFor, replaceWith);
        System.out.println("*****************************");
        System.out.println(request);
        System.out.println("*****************************");

        return processRequestEx(request);
	}
	
	public String processRequestEx(String requestData)
	{
		try
		{
			int serverPort = 443;
			String serverProtocol = "https";
			String workstation = "";
			String domain = "";

			String username = "android2@sc4j.org";
			String password = "Password1";
			String server = "connect.emailsrvr.com";
			String requestAbsoluteUrl = "https://connect.emailsrvr.com/EWS/Exchange.asmx";
			
			/*
			username = "office@teknision.com";
			password = "foozball2012";
			requestAbsoluteUrl = "https://m.google.com/EWS/Exchange.asmx";
			*/
			
			/*
			String requestData = createOutOfOfficeSoapRequest(username);
			sendRequestUsingNtlmAuthentication("ooo", username, password, workstation, domain, server, serverPort, serverProtocol, requestAbsoluteUrl, requestData);
			sendRequestUsingBasicAuthentication(username, password, workstation, domain, server, serverPort, serverProtocol, requestAbsoluteUrl, requestData);
			 */
			
			String result;
			//result = sendRequestUsingNtlmAuthentication(username, password, workstation, domain, server, serverPort, serverProtocol, requestAbsoluteUrl, requestData);
			result = sendRequestUsingBasicAuthentication(username, password, workstation, domain, server, serverPort, serverProtocol, requestAbsoluteUrl, requestData);
			
			return result;
		}
		catch (Throwable e)
		{
			return null;
		}
	}

	private String convertStreamToString(java.io.InputStream is) 
	{
	    java.util.Scanner s = new java.util.Scanner(is).useDelimiter("\\A");
	    return s.hasNext() ? s.next() : "";
	}
	
    private String formatXml(String xml)
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
    
	public String sendRequestUsingNtlmAuthentication(String username, String password, String workstation, String domain, String server, int serverPort, String serverProtocol, String requestAbsoluteUrl, String requestData) throws Throwable
	{
		DefaultHttpClient httpClient = new DefaultHttpClient();
		
		NTCredentials creds = new NTCredentials(username, password, workstation, domain);
		httpClient.getCredentialsProvider().setCredentials(AuthScope.ANY, creds);
		
		String result = sendRequest(httpClient, requestData, requestAbsoluteUrl, null);
		return result;
	}

	public String sendRequestUsingBasicAuthentication(String username, String password, String workstation, String domain, String server, int serverPort, String serverProtocol, String requestAbsoluteUrl, String requestData) throws Throwable
	{
		DefaultHttpClient httpClient = new DefaultHttpClient();
		String authHeaderValue = "Basic " + Base64.encodeBase64String((username + ":" + password).getBytes());

		// Uncomment and Use the following line instead, if targetting android
		//String authHeaderValue = "Basic " + org.kobjects.base64.Base64.encode((username + ":" + password).getBytes());

		return sendRequest(httpClient, requestData, requestAbsoluteUrl, authHeaderValue);
	}

	private String sendRequest(DefaultHttpClient httpClient, String requestContent, String serviceUrl, String authHeaderValue) throws Throwable 
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
				String errorMsg = "Got SOAP response code " + responseStatusCode + " " + httpResponse.getStatusLine().getReasonPhrase();
				System.out.println(errorMsg);
			}

			// get the response content
			HttpEntity httpEntity = httpResponse.getEntity();
			InputStream rs = httpEntity.getContent();
			
			String result = convertStreamToString(rs);
		
			if (result != null)
			{
				result = formatXml(result);
			}
			
			return result;
			
		} catch (Throwable e) {
			throw e;
		}
	}
}
