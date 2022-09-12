package excelreader;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCode

	{
		public static FileInputStream f;
		public static XSSFWorkbook w;
		public static XSSFSheet s;
		
		public static String readStringData(int row,int column) throws IOException
		{ 
			f=new FileInputStream("C:\\Users\\User\\Desktop\\samplefor java.xlsx"); // path of file
			w=new XSSFWorkbook(f);  //to take workbook from file f          
			s=w.getSheet("Sheet1");// to take sheet1 from workbook so no need to instiantiate
			Row r= s.getRow(row);
			Cell c=r.getCell(column);
			return c.getStringCellValue();         
				
		}
		
		public static String readIntegerData(int row,int column) throws IOException
		{
			f=new FileInputStream("C:\\Users\\User\\Desktop\\samplefor java.xlsx"); 
			w=new XSSFWorkbook(f);            
			s=w.getSheet("Sheet1");
			Row r= s.getRow(row);
			Cell c=r.getCell(column);
			int a=(int) c.getNumericCellValue();   //integer   
			return String.valueOf(a);                    // convert to string valueOf
			
		}
}
