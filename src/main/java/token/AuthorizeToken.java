package token;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.List;

/**
 * Created by ws5103 on 4/11/18.
 */
public class AuthorizeToken {

	/** Directory to store user credentials for this application. */
	private static java.io.File DATA_STORE_DIR;

	/** Global instance of the {@link FileDataStoreFactory}. */
	private static FileDataStoreFactory DATA_STORE_FACTORY;

	private static JsonFactory JSON_FACTORY;

	/** Global instance of the HTTP transport. */
	private static HttpTransport HTTP_TRANSPORT;

	/**List for scope storage */
	private static List<String> SCOPES;

	private AuthorizeToken(AuthorizationBuilder authorizationBuilder){
		this.DATA_STORE_FACTORY = authorizationBuilder.fileDataStoreFactory;
    this.JSON_FACTORY = authorizationBuilder.jsonFactory;
    this.HTTP_TRANSPORT = authorizationBuilder.httpTransport;
    this.SCOPES = authorizationBuilder.scopes;
    this.DATA_STORE_DIR = authorizationBuilder.DATA_STORE_DIR;
	}

	/**
	 * Build and Authorize the provided credentials
	 *
	 *@return an authorized credential instance
	 */

	public Credential authorize() throws IOException {
		// Load client secrets.
		InputStream in =
				AuthorizeToken.class.getResourceAsStream("/client_secret.json");
		GoogleClientSecrets clientSecrets =
				GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

		// Build flow and trigger user authorization request.
		GoogleAuthorizationCodeFlow flow =
				new GoogleAuthorizationCodeFlow.Builder(
						HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
						.setDataStoreFactory(DATA_STORE_FACTORY)
						.setAccessType("offline")
						.build();
		Credential credential = new AuthorizationCodeInstalledApp(
				flow, new LocalServerReceiver()).authorize("user");
		System.out.println(
				"Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
		return credential;
	}

	public static class AuthorizationBuilder {
		private FileDataStoreFactory fileDataStoreFactory;
		private JsonFactory jsonFactory;
		private HttpTransport httpTransport;
		private List<String> scopes;
		private java.io.File DATA_STORE_DIR;

		public AuthorizationBuilder(FileDataStoreFactory fileDataStoreFactory,JsonFactory jsonFactory,HttpTransport httpTransport,List<String> scopes){
			this.fileDataStoreFactory = fileDataStoreFactory;
			this.jsonFactory = jsonFactory;
			this.httpTransport = httpTransport;
			this.scopes = scopes;
		}

		public AuthorizationBuilder setDataStoreDirectory(java.io.File dataStoreDirectory){
			this.DATA_STORE_DIR = dataStoreDirectory;
			return this;
		}

		public AuthorizeToken build(){
			return new AuthorizeToken(this);
		}

	}
}
