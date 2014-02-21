/**
 * 
 */
package config.sandbox.dashboard;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;

import ariba.portlet.component.PortletContent;
import ariba.portlet.core.PortletConfig;
import ariba.portlet.util.Constants;
import ariba.util.core.Fmt;
import ariba.util.log.Log;

/**
 * @author fhenri
 * 
 */
public class DemoPortletContent extends PortletContent {

	public String _content;
	public boolean _hasErrors;

	/**
	 * @see ariba.portlet.component.PortletContent#initialize(ariba.portlet.core.PortletConfig)
	 */
	public void initialize(PortletConfig portletConfig) {
		Log.customer.debug("initialize portlet %s", portletConfig);
		super.initialize(portletConfig);
		_content = null;
		_hasErrors = false;
		try {
			_content = getContent();
		} catch (IOException ex) {
			_hasErrors = true;
		}
	}

	// webservice : http://www.xmethods.com/ve2/ViewListing.po?key=467542
	public final static String getContent() throws IOException {
		String content = "";

		content += Fmt.S("1 EUR = %s USD<br>", getRate("EUR", "USD"));
		content += Fmt.S("1 EUR = %s CHF<br>", getRate("EUR", "CHF"));
		content += Fmt.S("1 EUR = %s GBP<br>", getRate("EUR", "GBP"));

		return content;
	}

	/**
	 * @see ariba.portlet.component.PortletContent#getContentType()
	 */
	@Override
	public String getContentType() {
		return Constants.TypeHTML;
	}

	// we call the service and return the values
	private static String getRate (String fromCurrency, String toCurrency) {
		try {
	        String DestAddressURL = Fmt.S(
	                "http://www.webservicex.net/CurrencyConvertor.asmx/ConversionRate?FromCurrency=%s&ToCurrency=%s",
	                fromCurrency, toCurrency);
	        
            URL url = new URL(DestAddressURL);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();

            connection.setDoOutput(true);
            connection.setDoInput(true);
            connection.setUseCaches(false);

            connection.setRequestMethod("GET");

            
            // setting the header parameters
            connection.setRequestProperty("Cache-Control", "no-cache");

            BufferedReader in = new BufferedReader(
            		new InputStreamReader( connection.getInputStream()));
            in.readLine();
            String doubleLine = in.readLine();
            in.close();
            return doubleLine.substring(44, 50);

		} catch (Exception e) {
			return Fmt.S("error when getting exchange rate for %s to %s", fromCurrency, toCurrency);
		}
	}
	
	/*http://ws.apache.org/axis/java/user-guide.html
	  public class TestClient {
6     public static void main(String [] args) {
7       try {
8         String endpoint =
9             "http://ws.apache.org:5049/axis/services/echo";
10  
11        Service  service = new Service();
12        Call     call    = (Call) service.createCall();
13  
14        call.setTargetEndpointAddress( new java.net.URL(endpoint) );
15        call.setOperationName(new QName("http://soapinterop.org/", echoString"));
16  
17        String ret = (String) call.invoke( new Object[] { "Hello!" } );
18  
19        System.out.println("Sent 'Hello!', got '" + ret + "'");
20      } catch (Exception e) {
21        System.err.println(e.toString());
22      }
23    }
24  }
	 */
}
