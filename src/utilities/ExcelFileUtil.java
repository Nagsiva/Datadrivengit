package utilities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileUtil
{
 XSSFWorkbook wb;
 public ExcelFileUtil(String excelpath)throws Throwable 
 {
	 FileInputStream fi = new FileInputStream(excelpath);
	 wb = new XSSFWorkbook(fi);
 }
 public int rowcount(String sheetname)
 {
	 return wb.getSheet(sheetname).getLastRowNum();
 }
 public String getcelldata(String sheetname,int row,int column) 
 {
	String data = " ";
	if (wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC)
	{
	int celldata = (int) wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
	data = String.valueOf(celldata);
	}
	else
	{
	data = wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();	
	}
	return data;
}
 public void setcelldata(String sheetname,int row, int column,String status,String writeexcel) throws Throwable 
 {
	 XSSFSheet ws = wb.getSheet(sheetname);
	 XSSFRow rownum = ws.getRow(row);
	 XSSFCell cell = rownum.createCell(column);
	 cell.setCellValue(status);
	 if (status.equalsIgnoreCase("Pass"))
	 {
	XSSFCellStyle style = wb.createCellStyle();
	XSSFFont font = wb.createFont();
	font.setColor(IndexedColors.GREEN.getIndex());
	font.setBold(true);
	font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
	style.setFont(font);
	rownum.getCell(column).setCellStyle(style);
	}
	 else if (status.equalsIgnoreCase("Fail"))
	 {
		 XSSFCellStyle style = wb.createCellStyle();
			XSSFFont font = wb.createFont();
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
	 }
	 else if (status.equalsIgnoreCase("Blocked"))
	 {
		 XSSFCellStyle style = wb.createCellStyle();
			XSSFFont font = wb.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
	}
	FileOutputStream fo = new FileOutputStream(writeexcel);
	wb.write(fo);
	}
 public static void main(String[] args)throws Throwable
 {
ExcelFileUtil xl = new ExcelFileUtil("D://Nag.xlsx/");
int rc = xl.rowcount("EmpData");
System.out.println(rc);
for(int i=1;i<=rc;i++)
{
	String fname = xl.getcelldata("EmpData", i, 0);
	String mname = xl.getcelldata("EmpData", i, 1);
	String lname = xl.getcelldata("EmpData", i, 2);
	String eid = xl.getcelldata("EmpData", i, 3);
	System.out.println(fname+" "+mname+" "+lname+" "+eid);
	xl.setcelldata("EmpData",i, 4, "Pass", "D://amme.xlsx/");
	
}
 }
}