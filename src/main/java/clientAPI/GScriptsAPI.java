package clientAPI;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.script.Script;
import com.google.api.services.script.model.Content;
import com.google.api.services.script.model.ExecutionRequest;
import com.google.api.services.script.model.File;
import com.google.api.services.script.model.Operation;
import token.AuthorizeToken;

import java.io.IOException;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

/**
 * Created by ws5103 on 4/27/18.
 */
public class GScriptsAPI {
	private static GScriptsAPI instance;
	/** Application name. */
	private static final String APPLICATION_NAME =
			"Google Apps Script API Java GScriptsAPI";

	/** Directory to store user credentials for this application. */
	private static final java.io.File DEFAULT_DATA_STORE_DIR = new java.io.File(
			"./credentials/Google-API-Access-Credentials(Google Scripts)");

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
			Arrays.asList("https://www.googleapis.com/auth/script.projects","https://www.googleapis.com/auth/documents");

	private GScriptsAPI(){

	}

	public static synchronized GScriptsAPI getInstance(){
		if(instance == null){
			init();
			instance = new GScriptsAPI();
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

	/** Generating a script service instance with the authorised credentials
	 *
	 * @return an authorized script service*/

	private Script getScriptsService() throws IOException {
		Credential credential = this.credential;
		return new Script.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
				.setApplicationName(APPLICATION_NAME)
				.build();
	}

	public static void main(String[] args)throws IOException {
		GScriptsAPI gsAPI = GScriptsAPI.getInstance();
		Script service = gsAPI.getScriptsService();

		Script.Projects projects = service.projects();
		Content projectContents = projects.getContent("14QVSmsBvpGUWNt7OtB6P7jTwxeiq-T2En0XYss3UIlNDYBT5pUeqhSqQ").execute();
    List<File> prevFiles = projectContents.getFiles();
		Iterator<File> iterator = prevFiles.iterator();
    boolean FILE_CREATE_FLAG = true;

    while(iterator.hasNext()){
			if(iterator.next().getName().equals("hello")){
				FILE_CREATE_FLAG = false;
				break;
			}
		}

		if(FILE_CREATE_FLAG){
			File jsFile = new File()
					.setName("hello")
					.setType("SERVER_JS")
					.setSource("function helloWorld() {\n  Logger.log(\"Hello, world!\");\n}");

			prevFiles.add(jsFile);
			Content content = new Content().setFiles(prevFiles);
			Content updatedContent = projects.updateContent("14QVSmsBvpGUWNt7OtB6P7jTwxeiq-T2En0XYss3UIlNDYBT5pUeqhSqQ", content).execute();
		}

		/*iterator = prevFiles.iterator();
		while(iterator.hasNext()){
			File scriptFile = iterator.next();
			if(scriptFile.getName().equals("Code")){
				scriptFile.setSource("var doc = DocumentApp.getActiveDocument();\n" + "  var body = doc.getBody();\n" + "  var firstTable = body.getTables()[0];\n" + "  var rowNum = firstTable.getNumRows();\n" + "  var cell\n" + "  while(rowNum--){\n" + "    cell = firstTable.getCell(rowNum, 0)\n" + "    Logger.log(cell.getText());\n" + "  }");
				Content content = new Content().setFiles(prevFiles);
				projects.updateContent("14QVSmsBvpGUWNt7OtB6P7jTwxeiq-T2En0XYss3UIlNDYBT5pUeqhSqQ", content).execute();
				break;
			}
		}*/

		for(int index = 0; index < prevFiles.size(); index++){
			if(prevFiles.get(index).getName().equals("hello")){
				prevFiles.get(index).setSource("function zawadFunction() {\n" + "  var doc = DocumentApp.getActiveDocument();\n" + "  var body = doc.getBody();\n" + "  var firstTable = body.getTables()[0];\n" + "  var rowNum = firstTable.getNumRows();\n" + "  var cell\n" + "  while(rowNum--){\n" + "   firstTable.getCell(rowNum, 1).replaceText(\"^.*?$\", \"zawad\");\n" + "  }\n" + " \n" + "}");
				Content content = new Content().setFiles(prevFiles);
				projects.updateContent("14QVSmsBvpGUWNt7OtB6P7jTwxeiq-T2En0XYss3UIlNDYBT5pUeqhSqQ", content).execute();
			}
		}

		ExecutionRequest request = new ExecutionRequest().setFunction("zawadFunction").setDevMode(false);

		Operation op =
				service.scripts().run("14QVSmsBvpGUWNt7OtB6P7jTwxeiq-T2En0XYss3UIlNDYBT5pUeqhSqQ", request).execute();



	}
}
