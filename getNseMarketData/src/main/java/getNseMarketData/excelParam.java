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
		 String currentpath = new java.io.File( "." ).getCanonicalPath();
		 String excelpath = currentpath + "\\parameterfile\\MarketData.xlsx";
		FileInputStream fs = new FileInputStream(excelpath);
		//XSSFWorkbook wb = new XSSFWorkbook(fs);
		//XSSFSheet  sheet1 = wb.getSheetAt(0);
		//String[] passvalue = {"pricebandupper","applicableMargin"};
		//String[] ret=getMarketData("https://www1.nseindia.com/live_market/dynaContent/live_watch/get_quote/GetQuote.jsp?symbol=YESBANK", passvalue);
		exceldata();
		}
	
	public static String excelPath() throws IOException {
		String currentpath;
			currentpath = new java.io.File( "." ).getCanonicalPath();
		String excelpath = currentpath + "\\parameterfile\\MarketData.xlsx";

		return excelpath;

	}
	public static void exceldata() throws IOException {
		String appendurl= "https://www1.nseindia.com/live_market/dynaContent/live_watch/get_quote/GetQuote.jsp?symbol=";
		FileInputStream fs = new FileInputStream(excelPath());
		XSSFWorkbook wb = new XSSFWorkbook(fs);
		XSSFSheet  sheet1 = wb.getSheetAt(0);
		int maxCol = sheet1.getRow(0).getPhysicalNumberOfCells();
		int maxRow = sheet1.getPhysicalNumberOfRows();
		System.out.println("max row" + maxRow);
		int expCol = 0;
		String toTest,url;
		String[] stringArray = new String[28];
		String[] colHeader = new String[maxCol-8];
		int colCounter=0;
		for(int z=8;z<=maxCol-1;z++) {
			colHeader[colCounter]=sheet1.getRow(0).getCell(z).getStringCellValue();
			System.out.println(colHeader[colCounter]);
			colCounter++;
			
		}
		for(int j=1;j<maxRow;j++) {
			url=appendurl + sheet1.getRow(j).getCell(0).getStringCellValue();
			System.out.println(url);
			stringArray=getMarketData(url,colHeader);
			System.out.println("maxCol:" + maxCol);
			
			for(int i=8;i<=maxCol-1;i++)
			{
				//System.out.println("maxCol:" + maxCol);
				sheet1.getRow(j).createCell(i);//stringArray(j));
				sheet1.getRow(j).getCell(i).setCellValue(stringArray[i-8]);
				//setCellValue(stringArray[i-9]);
				
			}
		}
		FileOutputStream outputStream = new FileOutputStream(excelPath());
		wb.write(outputStream);

		
		//FileOutputStream fileOut = new FileOutputStream(excelPath());
		//wb.write(fileOut);
		//wb.close();
		//fileOut.flush();
		//fileOut.close();
				//return sheet1.getRow(1).getCell(expCol).getStringCellValue();
		 
		
	}
	
	
	

	public static String[] getMarketData(String url, String[] listColumn) throws IOException {
		Document docoutput;
		System.out.println("length of array"+listColumn.length);
		String[] returnvalue = new String[listColumn.length];
		try {
				docoutput = Jsoup.connect(url).userAgent("Mozilla/5.0").timeout(10 * 1000).get();
			
			
			Gson gson =  new Gson();
			Element display = docoutput.getElementById("responseDiv");
			
			JsonParser parser = new JsonParser();
			String s = display.text().toString();
			JsonElement obj =   parser.parse(s);
		
			JsonArray jsonsubarray =  (JsonArray) obj.getAsJsonObject().get("data");
			String jsonsubstring = jsonsubarray.toString().substring(1, jsonsubarray.toString().length()-1);
			
			JsonElement jsonsubelement=  parser.parse(jsonsubstring);
			for(int i=0;i<listColumn.length;i++) {
				System.out.println(listColumn[i]);
				System.out.println(jsonsubelement.getAsJsonObject().get(listColumn[i]).toString());
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
