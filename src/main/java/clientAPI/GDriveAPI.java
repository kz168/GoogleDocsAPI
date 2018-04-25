package clientAPI;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.drive.Drive;
import token.AuthorizeToken;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

/**
 * Created by ws5103 on 4/18/18.
 */
public class GDriveAPI {
	private static GDriveAPI instance;

	/** Application name. */
	private static final String APPLICATION_NAME =
			"Google Apps Script API Java clientAPI.GDriveAPI";

	/** Directory to store user credentials for this application. */
	private static final java.io.File DEFAULT_DATA_STORE_DIR = new java.io.File(
			"./src/main/credentials/Google-API-Access-Credentials(Google Drive)");

	/** Global instance of the {@link FileDataStoreFactory}. */
	private static FileDataStoreFactory DATA_STORE_FACTORY;

	/** Global instance of the JSON factory. */
	private static final JsonFactory JSON_FACTORY =
			JacksonFactory.getDefaultInstance();

	/** Global instance of the HTTP transport. */
	private static HttpTransport HTTP_TRANSPORT;

	/** Global instance for authorize Token */
	private static AuthorizeToken authorizeToken;

	/** Global Instance for Credential Object*/
	private static Credential credential;


	/** Global instance of the scopes required by this quickstart.
	 *
	 * If modifying these scopes, delete your previously saved credentials
	 * at ~/.credentials/Google-API-Access-Credentials
	 */
	private static final List<String> SCOPES =
			Arrays.asList("https://www.googleapis.com/auth/drive");

	private GDriveAPI(){

	}

	public static synchronized GDriveAPI getInstance(){
		if(instance == null){
			init();
			instance = new GDriveAPI();
		}

		return instance;
	}

	/** Initializing the API with default values and setting the credentials
	 */
	private static void init(){
		try {
			HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			DATA_STORE_FACTORY = new FileDataStoreFactory(DEFAULT_DATA_STORE_DIR);
			authorizeToken = new AuthorizeToken.AuthorizationBuilder(DATA_STORE_FACTORY,JSON_FACTORY,HTTP_TRANSPORT,SCOPES)
					.setDataStoreDirectory(DEFAULT_DATA_STORE_DIR)
					.build();
			credential = authorizeToken.authorize();
		} catch (Throwable t) {
			t.printStackTrace();
			System.exit(1);
		}
	}

	/**
	 * Build and return an authorized Drive client service.
	 * @return an authorized Drive client service
	 * @throws IOException
	 */
	private Drive getDriveService() throws IOException {
		Credential credential = this.credential;
		return new Drive.Builder(
				HTTP_TRANSPORT, JSON_FACTORY, credential)
				.setApplicationName(APPLICATION_NAME)
				.build();
	}

	public void DownloadFile(String fileId,String nameOfFile)throws IOException{
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		Drive driveService = getDriveService();
		driveService.files().export(fileId, "application/pdf")
				.executeMediaAndDownloadTo(outputStream);

		File file = new File("./Downloads/" + nameOfFile);
		if(file.getParentFile().mkdir()){
			if(file.createNewFile()){
				FileOutputStream fin = new FileOutputStream(file);
				fin.write(outputStream.toByteArray());
				System.out.println("File Downloaded At : " + "Download/" + nameOfFile);
			}
		}else{
			if(file.createNewFile()){
				FileOutputStream fin = new FileOutputStream(file);
				fin.write(outputStream.toByteArray());
				System.out.println("File Downloaded At : " + "Downloads/" + nameOfFile);
			}else{
				throw new IOException("File Name already Exists");
			}
		}

	}
}
