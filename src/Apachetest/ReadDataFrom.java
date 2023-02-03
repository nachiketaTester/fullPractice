package Apachetest;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;


public class ReadDataFrom {
@Test
public void writeData() throws IOException
{
	FileInputStream fis=new FileInputStream("D://Pupun.xlsx");
	XSSFWorkbook wb = new XSSFWorkbook(fis) ;
		XSSFSheet s=wb.getSheet("TestData");
		XSSFRow r=s.getRow(3);
		XSSFCell c=s.getRow(3).getCell(2);
	System.out.println(c.toString());
	System.out.println(c.getStringCellValue());
	
}
}
