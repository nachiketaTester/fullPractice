package Apachetest;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class REadData {
@Test
public void readdata() throws IOException
{
	File src=new File("TestExcel//Book1.xlsx ");
	FileInputStream fis=new FileInputStream(src);
	XSSFWorkbook wb=new XSSFWorkbook(fis);
	System.out.println(wb.getSheet("TestData").getRow(3).getCell(1).toString());
}
}
