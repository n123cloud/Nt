package getNseMarketData;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.SocketTimeoutException;

//import org.apache.poi.*;
import org.apache.poi.xssf.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;

public class excelParam {

	public static void main(String[] args) throws IOException,NullPointerException {
		// TODO Auto-generated method stub
				
		exceldata();
		}
	
	public static String excelPath() throws IOException {
		///function to get the path of excel which will give the Parameters(columns) and store the scrapped Market Data
		String currentpath;
		currentpath = new java.io.File( "." ).getCanonicalPath();
		String excelpath = currentpath + "\\parameterfile\\MarketData.xlsx";

		return excelpath;

	}
	public static void exceldata() throws IOException {
		//Function to Write MArket Data(current Price) into excel.
		//for constructing URL
		String appendurl= "https://www1.nseindia.com/live_market/dynaContent/live_watch/get_quote/GetQuote.jsp?symbol=";
		FileInputStream fs = new FileInputStream(excelPath());
		//Use Apache POI to read Excel
		XSSFWorkbook wb = new XSSFWorkbook(fs);
		XSSFSheet  sheet1 = wb.getSheetAt(0);
		int maxCol = sheet1.getRow(0).getPhysicalNumberOfCells();
		int maxRow = sheet1.getPhysicalNumberOfRows();
		int expCol = 0;
		String toTest,url;
		//Update the max array Size in case you want to get more number of column
		String[] stringArray = new String[28];
		//First 8 Columns are static.
		String[] colHeader = new String[maxCol-8];
		int colCounter=0;
		//Get all the Parameters which you want to fetch from Website
		for(int z=8;z<=maxCol-1;z++) {
			colHeader[colCounter]=sheet1.getRow(0).getCell(z).getStringCellValue();
			colCounter++;
			
		}
		//Run for each Rows(Stock)
		for(int j=1;j<maxRow;j++) {
			url=appendurl + sheet1.getRow(j).getCell(0).getStringCellValue();
			//Get Market data in String Array
			stringArray=getMarketData(url,colHeader);
			//Save data for each Record in Excel.Replace it with DB Query in case  of Write to DB
			for(int i=8;i<=maxCol-1;i++)
			{
				sheet1.getRow(j).createCell(i);//stringArray(j));
				sheet1.getRow(j).getCell(i).setCellValue(stringArray[i-8]);
				
			}
		}
		FileOutputStream outputStream = new FileOutputStream(excelPath());
		wb.write(outputStream);

		
		
	}
	
	public static String[] getMarketData(String url, String[] listColumn) throws IOException {
		//Parameter,URL=Constructed NSE URL , listColumn= List of parameters(Eg:lastPrice)

		Document docoutput;
		String[] returnvalue = new String[listColumn.length];
		try {
			//Adjust or reduce the timeout if needed
				docoutput = Jsoup.connect(url).userAgent("Mozilla/5.0").timeout(10 * 1000).get();
			
			//Read the Json from the Website Response
			Gson gson =  new Gson();
			//All the Market Data is in this Element
			Element display = docoutput.getElementById("responseDiv");
			
			JsonParser parser = new JsonParser();
			String s = display.text().toString();
			JsonElement obj =   parser.parse(s);
		
			JsonArray jsonsubarray =  (JsonArray) obj.getAsJsonObject().get("data");
			//Trim both the end of the response string 
			String jsonsubstring = jsonsubarray.toString().substring(1, jsonsubarray.toString().length()-1);
			JsonElement jsonsubelement=  parser.parse(jsonsubstring);
			//Construct the String array with all the Market data
			for(int i=0;i<listColumn.length;i++) {
				returnvalue[i]=jsonsubelement.getAsJsonObject().get(listColumn[i]).toString().replace("\"", "");
						
			}
			return returnvalue;
		
		}catch (IllegalStateException e) {
			returnvalue[0]="<<unable to fetch data>>";
			return returnvalue;
		}
		catch (SocketTimeoutException e) {
			// TODO: handle exception
			returnvalue[0]="<<timeout error >>";
			return returnvalue;
		}
	}
	

}
