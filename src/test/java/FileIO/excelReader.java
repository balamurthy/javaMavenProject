package FileIO;


import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//How to read excel files using Apache POI
public class excelReader {

	static XSSFCell s_custid;
	static XSSFCell s_custname;



	public static void main (String [] args) throws IOException, InterruptedException{


		FileInputStream fis = new FileInputStream("D:\\blaptop\\ascendion\\Test.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		System.out.println(sheet.getLastRowNum());
		int rowcount = sheet.getLastRowNum();
		for (int i = 1; i < rowcount +1; i ++) {
			s_custid = sheet.getRow(i).getCell(0);
			s_custname = sheet.getRow(i).getCell(1);
			display(s_custid,s_custname);
		}
		//String cellval = cell.getStringCellValue();
		//System.out.println(cellv

	}	
	static void display(XSSFCell s_custid2,XSSFCell s_custname2) throws InterruptedException
	{

		System.out.println("Custid is:" + s_custid2 + " Custname is : " + s_custname2);
	}


}



