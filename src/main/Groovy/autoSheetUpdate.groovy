import clientAPI.GSheetsAPI
import com.astronlab.ngenhttplib.http.core.HttpInvoker
import com.astronlab.ngenhttplib.http.core.request.InvokerRequestBody
import org.apache.commons.csv.CSVFormat
import org.apache.commons.csv.CSVParser
import org.apache.commons.csv.CSVRecord
import utils.Range

import java.nio.charset.Charset

AutoSheetUpdate sheetUpdate = new AutoSheetUpdate()
sheetUpdate.update()

class AutoSheetUpdate{
	GSheetsAPI gSheetApi
	HttpInvoker invoker

	AutoSheetUpdate(){
		gSheetApi = GSheetsAPI.getInstance()
		invoker = new HttpInvoker()

	}

	def update(){
		def url = "http://pmis.sebpo.net/index.php"
		//Change the spreadSheetUrl to the one you want to access
		def spreadsheetUrl = "https://docs.google.com/spreadsheets/d/136TZW6b2AXbAFLxZPTGk0NHO7Cuh69qPrWJtMKGRKn8/edit#gid=96910449"

		//set the sheet title to the sheet that you want to update
		def sheetTitle = "Task details"
		def data = invoker.getStringData(url)

		def postUrlMatch = data =~ /href="([^"]+).*?class="active">Login/
		def postUrl
		if(postUrlMatch){
			postUrl = url + postUrlMatch[0][1].toString().replaceAll(/&amp;/,"&").trim()
		}
		def request = invoker.init(postUrl)
		def formData = [
				action : "login",
				username : "md.omar",
				passhash : "cc608d65570ba41d666720532dc6f88b",
				remain : "false",
				area: "loginpage"
		]
		def formBody = InvokerRequestBody.createViaSingleFormBody()
		formBody.addParams(formData)
		request = request.post(formBody)
		request.setRequestHeader("x-requested-with","XMLHttpRequest")
		request.execute()

		def startDate = getStartDate()
		def endDate = getEndDate()

		def csvUrl = "http://pmis.sebpo.net/index.php?ext=daytracks&controller=export&export[employee]=&export[employer]=&export[project]=6&export[company]=&export[date_start]=$startDate&export[date_end]=$endDate&action=download"
		def csvData = invoker.getStringData(csvUrl)

		File dataFile = createDataFile("./Data/dataCSV.csv",csvData)

		CSVParser parser = CSVParser.parse(dataFile, Charset.defaultCharset(),
				CSVFormat.DEFAULT
						.withFirstRecordAsHeader()
						.withIgnoreHeaderCase());

		List<List<Object>> values = new ArrayList<>();
		List<List<Object>> hourValue = new ArrayList<>();
		def defaultTimeValue = "00:00"
		for(CSVRecord csvRecord : parser){
			List<Object> row;
			Object[] rowValue = [
				csvRecord.get("Task No."),
				csvRecord.get("Task Title"),
				csvRecord.get("Date"),
        defaultTimeValue,
				csvRecord.get("Person"),
				csvRecord.get("Activity Types")
			];
			row = Arrays.asList(rowValue);
			values.add(row);

			rowValue = [
					csvRecord.get("Chargeable Time")
			]
			row = Arrays.asList(rowValue)
			hourValue.add(row)
		}

		gSheetApi.clearAndAppend(spreadsheetUrl,sheetTitle,values);

		Range range = new Range(sheetTitle,"D2","D" + (hourValue.size() + 1))

		gSheetApi.updateDataToRange(spreadsheetUrl,range,hourValue,"USER_ENTERED")
	}

	def createDataFile(String path,String data){
		File file = new File(path)
		if(file.getParentFile().mkdir()){
			if(file.createNewFile()){
				FileOutputStream fout = new FileOutputStream(file)
				fout.write(data.getBytes())
			}
		}else {
			FileOutputStream fout = new FileOutputStream(file)
			fout.write(data.getBytes())
		}

		return file
	}

	def getStartDate(){
		Calendar c = Calendar.getInstance();
		c.set(Calendar.DAY_OF_WEEK, Calendar.MONDAY);
		def month = c.get(Calendar.MONTH) + 1
		c.add(Calendar.DATE,-7)
		def day = c.get(Calendar.DATE)

		def date = month + "/" + day + "/" + c.get(Calendar.YEAR)
		date = date.replaceAll(/\b(\d{1})\b\//, "0\$1/")

		return date
	}

	def getEndDate(){
		Calendar c = Calendar.getInstance();
		c.set(Calendar.DAY_OF_WEEK, Calendar.FRIDAY);
		def month = c.get(Calendar.MONTH) + 1
		c.add(Calendar.DATE,-7)
		def day = c.get(Calendar.DATE)

		def date = month + "/" + day + "/" + c.get(Calendar.YEAR)
		date = date.replaceAll(/\b(\d{1})\b\//, "0\$1/")

		return date

	}

}