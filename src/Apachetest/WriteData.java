package Apachetest;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class WriteData {
@Test
public void writeData() throws IOException
{
	FileInputStream fis=new FileInputStream("D://Pupun.xlsx");
	XSSFWorkbook wb=new XSSFWorkbook(fis);
	//Sheet s=wb.getSheet("TestData");
	XSSFSheet s=wb.createSheet("Love2");
	XSSFRow r=s.createRow(6);
	XSSFCell c=r.createCell(1);
	c.setCellValue("Nachiketa");
	FileOutputStream fo=new FileOutputStream("D://Pupun.xlsx");
	wb.write(fo);
	//fo.close();
	
	
	
}
}
