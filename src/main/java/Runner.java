import clientAPI.GSheetsAPI;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;

import java.io.File;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by ws5103 on 4/11/18.
 */
public class Runner {
	public static void main(String[] args) throws IOException{
		String spreadsheetUrl = "https://docs.google.com/spreadsheets/d/136TZW6b2AXbAFLxZPTGk0NHO7Cuh69qPrWJtMKGRKn8/edit#gid=96910449";
		GSheetsAPI api = GSheetsAPI.getInstance();
		File csvFile = new File("/home/ws5103/GAppsScriptProject/Data/testCSV.csv");

		/*CSVParser parser = CSVParser.parse(csvFile, Charset.defaultCharset(),CSVFormat.RFC4180.withHeader(
				"Task No.","Task Title","Date","Tracked Time","Chargeable Time","Company","Project","Person","Comment","Activity Types"
		).withIgnoreHeaderCase());*/

		CSVParser parser = CSVParser.parse(csvFile, Charset.defaultCharset(),
				CSVFormat.RFC4180
				.withFirstRecordAsHeader()
				.withIgnoreHeaderCase());

		List<List<Object>> values = new ArrayList<>();

		for(CSVRecord csvRecord : parser){
			List<Object> row;
			Object[] rowValue = {
					csvRecord.get("Task No."),
					csvRecord.get("Task Title"),
					csvRecord.get("Date"),
					csvRecord.get("Chargeable Time"),
					csvRecord.get("Person"),
					csvRecord.get("Activity Types")};

			row = Arrays.asList(rowValue);
			values.add(row);
		}
		/*Range rangeData = new Range("Task details","A1","F9");
		api.appendDataToRange(spreadsheetUrl,rangeData,values,"RAW");*/



    //List<List<Object>> values = api.getfullData("https://docs.google.com/spreadsheets/d/136TZW6b2AXbAFLxZPTGk0NHO7Cuh69qPrWJtMKGRKn8/edit#gid=96910449","Task details");
		//List<List<Object>> values = api.getDataForRange(spreadsheetUrl,rangeData);

		//System.out.println(values.get(1).get(1));

		/*List<List<Object>> values = Arrays.asList(
				Arrays.asList(
						new Object[]{"First RowStart",
						  "Fdata1"}
						),
				Arrays.asList(
						new Object[]{"Second RowStart",
						 "Sdata1"}
				)
		);*/
		api.clearAndAppend(spreadsheetUrl,"Task details",values);
		//api.appendData(spreadsheetUrl,"Task details",values,"RAW");
		//values = api.getfullData(spreadsheetUrl,"Task details");
		//String valueInputOption = "RAW";
		//api.appendDataToRange(spreadsheetUrl,"","","",values,valueInputOption);
    //System.out.print(values.get(0).get(0));

		//api.setTitle("https://docs.google.com/spreadsheets/d/1Rc1nbcTHAjXZscFMMsSouHBea8QFFYH-xLVSm9O7xgw/edit","Zawad's SpreadSheet");
    //api.downloadSpreadSheet(spreadsheetUrl,"downLoadedfile");
	}


}
