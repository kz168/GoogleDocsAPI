package clientAPI;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import token.AuthorizeToken;
import utils.Range;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class GSheetsAPI {
    private static GSheetsAPI instance;
    /** Application name. */
    private static final String APPLICATION_NAME =
        "Google Apps Script API Java clientAPI.GSheetsAPI";

    /** Directory to store user credentials for this application. */
    private static final java.io.File DEFAULT_DATA_STORE_DIR = new java.io.File(
        "./credentials/Google-API-Access-Credentials(Google Sheets)");

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
        Arrays.asList("https://www.googleapis.com/auth/spreadsheets");

    private GSheetsAPI(){

    }

    public static synchronized GSheetsAPI getInstance(){
        if(instance == null){
            init();
            instance = new GSheetsAPI();
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
     * @return an authorized Sheets service*/

    private Sheets getSheetsService() throws IOException {
        Credential credential = this.credential;
        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
            .setApplicationName(APPLICATION_NAME)
            .build();
    }

    /** Updating the title of the sheet associated with the spreadsheetId
     *
     * @param spreadSheetUrl url of the desired spreadsheet
     * @param title the updated title
     * @return an authorized Sheets service*/

    public void setTitle(String spreadSheetUrl,String title){
        Pattern pattern = Pattern.compile("(?is)d/([^/]+)/(?:.*?gid=([^$]+))?");
        Matcher matchID = pattern.matcher(spreadSheetUrl);
        String spreadSheetId = "";
        List<Request> requests = new ArrayList<>();

        if(matchID.find()){
            spreadSheetId = matchID.group(1).trim();
        }

        requests.add(new Request()
            .setUpdateSpreadsheetProperties(new UpdateSpreadsheetPropertiesRequest()
                .setProperties(new SpreadsheetProperties()
                    .setTitle(title))
                .setFields("title")));
        try {
            BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(
                requests);
            BatchUpdateSpreadsheetResponse response = getSheetsService().spreadsheets().batchUpdate(
                spreadSheetId, body).execute();

        }catch (IOException ioe){
            System.out.println("Google API returned an error when trying to set the title");
            System.out.println(ioe.getMessage());
        }


    }
    /** Fetching the full data of sheet contained in a spreadsheet
     *
     * @param spreadSheetUrl url of the desired spreadsheet
     * @param sheetTitle title of the sheet to access the data if null or empty then the first sheet is accessed
     * @return List<List<Object>> a list of objects containing the result*/

    public List<List<Object>> getfullData(String spreadSheetUrl,String sheetTitle) throws IOException{
        Pattern pattern = Pattern.compile("(?is)d/([^/]+)/(?:.*?gid=([^$]+))?");
        Matcher matchID = pattern.matcher(spreadSheetUrl);
        String spreadSheetId = "";
        Sheets service = getSheetsService();

        if(matchID.find()){
            spreadSheetId = matchID.group(1).trim();
        }


        if(sheetTitle == null || sheetTitle.equals("")){
            Sheets.Spreadsheets.Get request = service.spreadsheets().get(spreadSheetId);
            Spreadsheet response = request.execute();
          List<Sheet> listofSheets = response.getSheets();
          sheetTitle = listofSheets.get(0).getProperties().getTitle();
        }

        ValueRange result = service.spreadsheets().values().get(spreadSheetId,sheetTitle).execute();
        List<List<Object>> returnValue = result.getValues();

        return returnValue;
    }

    /** Fetching ranged data of sheet contained in a spreadsheet
     *
     * @param spreadSheetUrl url of the desired spreadsheet
     * @param range range Object for the access range
     *
     * The range value for data access is a string like sheetTitle!A1:B2
     * sheetTitle!A1:B2 refers to the first two cells in the top two rows of Sheet1.
     * sheetTitle!A:A refers to all the cells in the first column of Sheet1.
     * sheetTitle!1:2 refers to the all the cells in the first two rows of Sheet1.
     * sheetTitle!A5:A refers to all the cells of the first column of Sheet 1, from row 5 onward.
     *
     * For more information use this "https://developers.google.com/sheets/api/guides/concepts"
     * @return List<List<Object>> a list of objects containing the result*/

    public List<List<Object>> getDataForRange(String spreadSheetUrl, Range range)throws IOException{
        Pattern pattern = Pattern.compile("(?is)d/([^/]+)/(?:.*?gid=([^$]+))?");
        Matcher matchID = pattern.matcher(spreadSheetUrl);
        String spreadSheetId = "";
        Sheets service = getSheetsService();

        if(matchID.find()){
            spreadSheetId = matchID.group(1).trim();
        }

        if(range.getSheetTitle() == null || range.getSheetTitle().equals("")){
            Sheets.Spreadsheets.Get request = service.spreadsheets().get(spreadSheetId);
            Spreadsheet response = request.execute();
            List<Sheet> listofSheets = response.getSheets();
            range.setSheetTitle(listofSheets.get(0).getProperties().getTitle());
        }

        if(range.getFirstCellPointer() == null || range.getLastCellPointer() == null){
            range.setFirstCellPointer("");
            range.setLastCellPointer("");
        }
        String rangeValue = range.toString();


        ValueRange result = service.spreadsheets().values().get(spreadSheetId,rangeValue).execute();
        List<List<Object>> returnValue = result.getValues();

        return returnValue;
    }

    /** Updating ranged data of sheet contained in a spreadsheet
     *
     * @param spreadSheetUrl url of the desired spreadsheet
     * @param range range Object for the access range
     * @param updateData a List<List<Object>> which holds the data for update data is updated row wise
     * @param valueInputOption RAW or USER_ENTERED depending on the parsing of data in server side
     *
     * Details of the range notation
     *
     * The range value for data access is a string like sheetTitle!A1:B2
     * sheetTitle!A1:B2 refers to the first two cells in the top two rows of Sheet1.
     * sheetTitle!A:A refers to all the cells in the first column of Sheet1.
     * sheetTitle!1:2 refers to the all the cells in the first two rows of Sheet1.
     * sheetTitle!A5:A refers to all the cells of the first column of Sheet 1, from row 5 onward.
     *
     * For more information use this "https://developers.google.com/sheets/api/guides/concepts"
     * and "https://developers.google.com/sheets/api/guides/values"*/

    public void updateDataToRange(String spreadSheetUrl,Range range,List<List<Object>> updateData,String valueInputOption)throws IOException{
        Pattern pattern = Pattern.compile("(?is)d/([^/]+)/(?:.*?gid=([^$]+))?");
        Matcher matchID = pattern.matcher(spreadSheetUrl);
        String spreadSheetId = "";
        Sheets service = getSheetsService();

        if(matchID.find()){
            spreadSheetId = matchID.group(1).trim();
        }

        if(range.getSheetTitle() == null || range.getSheetTitle().equals("")){
            Sheets.Spreadsheets.Get request = service.spreadsheets().get(spreadSheetId);
            Spreadsheet response = request.execute();
            List<Sheet> listofSheets = response.getSheets();
            range.setSheetTitle(listofSheets.get(0).getProperties().getTitle());
        }

        if(range.getFirstCellPointer() == null || range.getLastCellPointer() == null){
            range.setFirstCellPointer("");
            range.setLastCellPointer("");
        }
        String rangeValue = range.toString();

        ValueRange body = new ValueRange();
        body.setValues(updateData);

        UpdateValuesResponse result =
            service.spreadsheets().values().update(spreadSheetId, rangeValue, body)
                .setValueInputOption(valueInputOption)
                .execute();
        System.out.println(result.getUpdatedCells() + " cells updated");
    }

    /**append data after the table formed by specified range and
     * if no range is specified data is appended after last row
     * For more information use this "https://developers.google.com/sheets/api/guides/values"*/
    public void appendDataAfterRange(String spreadSheetUrl,Range range,List<List<Object>> appendData,String valueInputOption)throws IOException{
        Pattern pattern = Pattern.compile("(?is)d/([^/]+)/(?:.*?gid=([^$]+))?");
        Matcher matchID = pattern.matcher(spreadSheetUrl);
        String spreadSheetId = "";
        Sheets service = getSheetsService();

        if(matchID.find()){
            spreadSheetId = matchID.group(1).trim();
        }

        if(range.getSheetTitle() == null || range.getSheetTitle().equals("")){
            Sheets.Spreadsheets.Get request = service.spreadsheets().get(spreadSheetId);
            Spreadsheet response = request.execute();
            List<Sheet> listofSheets = response.getSheets();
            range.setSheetTitle(listofSheets.get(0).getProperties().getTitle());
        }

        if(range.getFirstCellPointer() == null || range.getLastCellPointer() == null){
            range.setFirstCellPointer("");
            range.setLastCellPointer("");
        }
        String rangeValue = range.toString();

        ValueRange body = new ValueRange()
            .setValues(appendData);
        AppendValuesResponse result =
            service.spreadsheets().values().append(spreadSheetId, rangeValue, body)
                .setValueInputOption(valueInputOption)
                .setInsertDataOption("INSERT_ROWS")
                .execute();
        System.out.println(result.getUpdates().getUpdatedRows() + " rows appended");
    }

    /*Create a new SpreadSheet*/
    public String createSpreadSheet() throws IOException{
        Spreadsheet requestBody = new Spreadsheet();

        Sheets sheetsService = getSheetsService();
        Sheets.Spreadsheets.Create request = sheetsService.spreadsheets().create(requestBody);

        Spreadsheet response = request.execute();
        return response.getSpreadsheetUrl();
    }

    /**Append Data at the end sheet
     *
     * @param spreadSheetUrl url of the desired spreadsheet
     * @param sheetTitle title of the desired sheet inside the spreadsheet
     * @param appendData Data to be appended
     * @param valueInputOption RAW or USER_ENTERED depending on the parsing of data in server side
     * @throws IOException
     */
    public void appendData(String spreadSheetUrl,String sheetTitle,List<List<Object>> appendData,String valueInputOption)throws IOException{
        Pattern pattern = Pattern.compile("(?is)d/([^/]+)/(?:.*?gid=([^$]+))?");
        Matcher matchID = pattern.matcher(spreadSheetUrl);
        String spreadSheetId = "";
        Sheets service = getSheetsService();

        if(matchID.find()){
            spreadSheetId = matchID.group(1).trim();
        }

        if(sheetTitle == null || sheetTitle.equals("")){
            Sheets.Spreadsheets.Get request = service.spreadsheets().get(spreadSheetId);
            Spreadsheet response = request.execute();
            List<Sheet> listofSheets = response.getSheets();
            sheetTitle = listofSheets.get(0).getProperties().getTitle();
        }

        List<List<Object>> fullResult = getfullData(spreadSheetUrl,sheetTitle);
        int endRowCounter = 1;

        for(List<Object> row : fullResult){
            if(row.size() > 1){
                endRowCounter++;
            }else {
               break;
            }
        }
        Range range = new Range(sheetTitle,"A1","F" + endRowCounter);
        appendDataAfterRange(spreadSheetUrl,range,appendData,valueInputOption);
       // System.out.println(range.toString());

    }

    /**Downloads the desired spreadsheet
     *
     * @param spreadSheetUrl url of the desired spreadsheet
     * @param nameOfDownloadFile Name of the file to be downloaded
     * @throws IOException
     */


    public void downloadSpreadSheet(String spreadSheetUrl,String nameOfDownloadFile)throws IOException{
        GDriveAPI gDriveAPI = GDriveAPI.getInstance();

        Pattern pattern = Pattern.compile("(?is)d/([^/]+)/(?:.*?gid=([^$]+))?");
        Matcher matchID = pattern.matcher(spreadSheetUrl);
        String spreadSheetId = "";

        if(matchID.find()){
            spreadSheetId = matchID.group(1).trim();
        }

        gDriveAPI.DownloadFile(spreadSheetId,nameOfDownloadFile);
    }

    /**Clears the row according to set index
     *
     * @param spreadSheetId ID of the spread sheet
     * @param sheetTitle title of the sheet
     * @param startIndex starting index for the rows to be deleted
     * @param endIndex ending index for the rows to be deleted
     * (E.g. startIndex = 0 and endIndex = 3 deletes first three rows)
     * @throws IOException
     */

    public void clearRows(String spreadSheetId,String sheetTitle,Integer startIndex,Integer endIndex)throws IOException{
        Sheets service = getSheetsService();
        Integer sheetID = 0;
        Sheets.Spreadsheets.Get request = service.spreadsheets().get(spreadSheetId);
        Spreadsheet response = request.execute();
        List<Sheet> listofSheets = response.getSheets();

        for(Sheet sheet : listofSheets){
            if(sheet.getProperties().getTitle().equals(sheetTitle)){
                sheetID = sheet.getProperties().getSheetId();
                break;
            }
        }

        BatchUpdateSpreadsheetRequest content = new BatchUpdateSpreadsheetRequest();
        Request requestForDelete = new Request();
        DeleteDimensionRequest deleteDimensionRequest = new DeleteDimensionRequest();
        DimensionRange dimensionRange = new DimensionRange();
        dimensionRange.setDimension("ROWS");
        dimensionRange.setStartIndex(startIndex);
        dimensionRange.setEndIndex(endIndex);

        dimensionRange.setSheetId(sheetID);
        deleteDimensionRequest.setRange(dimensionRange);

        requestForDelete.setDeleteDimension(deleteDimensionRequest);

        List<Request> requests = new ArrayList<Request>();
        requests.add(requestForDelete);
        content.setRequests(requests);

        BatchUpdateSpreadsheetResponse responseForDelete = service.spreadsheets().batchUpdate(spreadSheetId,content).execute();
        if(responseForDelete.getReplies().get(0).isEmpty()){
            System.out.println((endIndex-startIndex) + " rows deleted");
        }
    }

    public void clearAndAppend(String spreadSheetUrl,String sheetTitle,List<List<Object>> appendData)throws IOException{
        Pattern pattern = Pattern.compile("(?is)d/([^/]+)/(?:.*?gid=([^$]+))?");
        Matcher matchID = pattern.matcher(spreadSheetUrl);
        String spreadSheetId = "";
        Integer sheetID = 0;
        Sheets service = getSheetsService();

        if(matchID.find()){
            spreadSheetId = matchID.group(1).trim();
        }

        if(sheetTitle == null || sheetTitle.equals("")){
            Sheets.Spreadsheets.Get request = service.spreadsheets().get(spreadSheetId);
            Spreadsheet response = request.execute();
            List<Sheet> listofSheets = response.getSheets();
            sheetTitle = listofSheets.get(0).getProperties().getTitle();
        }

        List<List<Object>> fullResult = getfullData(spreadSheetUrl,sheetTitle);
        int endRowCounter = 1;

        for(List<Object> row : fullResult){
            if(row.size() > 1){
                endRowCounter++;
            }else {
                break;
            }
        }

        clearRows(spreadSheetId,"Task details",1,endRowCounter-1);
        Range range = new Range(sheetTitle);
        range.setFirstCellPointer("A2");
        range.setLastCellPointer("F2");
        appendDataAfterRange(spreadSheetUrl,range,appendData,"RAW");
        int sumValueCell = (appendData.size()+10);
        String sumData = "=SUM(D2:" + "D" + sumValueCell + ")" ;
        List<List<Object>> updateSumValue = Arrays.asList(
            Arrays.asList(
                new Object[]{sumData}
            )
        );

        range.setFirstCellPointer("D" + (sumValueCell + 1));
        range.setLastCellPointer("");
        updateDataToRange(spreadSheetUrl,range,updateSumValue,"USER_ENTERED");

    }
}