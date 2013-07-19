import java.io.InputStream;
import java.io.UnsupportedEncodingException;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.auth.AuthScope;
import org.apache.http.auth.NTCredentials;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.commons.codec.binary.Base64;

public class MainApp 
{
	public static void main(String[] argv) 
	{
		try
		{
			int serverPort = 443;
			String serverProtocol = "https";
			String workstation = "";
			String domain = "";

			String username = "example@emailaddress.com";
			String password = "YourPassword";
			String server = "exchange.server.com";
			String requestAbsoluteUrl = "https://exchange.server.com/EWS/Exchange.asmx";
			String requestData = createOutOfOfficeSoapRequest(username);

			sendRequestUsingNtlmAuthentication(username, password, workstation, domain, server, serverPort, serverProtocol, requestAbsoluteUrl, requestData);
			sendRequestUsingBasicAuthentication(username, password, workstation, domain, server, serverPort, serverProtocol, requestAbsoluteUrl, requestData);
		}
		catch (Throwable e)
		{
			e.printStackTrace();
		}
	}

	private static String createOutOfOfficeSoapRequest(String emailAddress)
	{
		return "<v:Envelope xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:d=\"http://www.w3.org/2001/XMLSchema\" xmlns:c=\"http://schemas.xmlsoap.org/soap/encoding/\" xmlns:v=\"http://schemas.xmlsoap.org/soap/envelope/\"><v:Body><GetUserOofSettingsRequest xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\"><Mailbox xmlns=\"http://schemas.microsoft.com/exchange/services/2006/types\"><Address>"+emailAddress+"</Address></Mailbox></GetUserOofSettingsRequest></v:Body></v:Envelope>";
	}   

	private static String convertStreamToString(java.io.InputStream is) {
	    java.util.Scanner s = new java.util.Scanner(is).useDelimiter("\\A");
	    return s.hasNext() ? s.next() : "";
	}
	
	public static void sendRequestUsingNtlmAuthentication(String username, String password, String workstation, String domain, String server, int serverPort, String serverProtocol, String requestAbsoluteUrl, String requestData) throws Throwable
	{
		DefaultHttpClient httpClient = new DefaultHttpClient();
		NTCredentials creds = new NTCredentials(username, password, workstation, domain);
		httpClient.getCredentialsProvider().setCredentials(AuthScope.ANY, creds);
		
		sendRequest(httpClient, requestData, requestAbsoluteUrl, null);
	}

	public static void sendRequestUsingBasicAuthentication(String username, String password, String workstation, String domain, String server, int serverPort, String serverProtocol, String requestAbsoluteUrl, String requestData) throws Throwable
	{
		DefaultHttpClient httpClient = new DefaultHttpClient();
		String authHeaderValue = "Basic " + Base64.encodeBase64String((username + ":" + password).getBytes());
		
		// Uncomment and Use the following line instead, if targetting android
		// String authHeaderValue = "Basic " + org.kobjects.base64.Base64.encode((username + ":" + password).getBytes());
		
		sendRequest(httpClient, requestData, requestAbsoluteUrl, authHeaderValue);
	}

	private static InputStream sendRequest(DefaultHttpClient httpClient, String requestContent, String serviceUrl, String authHeaderValue) throws Throwable 
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
			
			return result;
		} catch (Throwable e) {
			throw e;
		}
	}
}
