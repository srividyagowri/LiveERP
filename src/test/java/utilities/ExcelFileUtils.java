package utilities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import javax.imageio.stream.FileImageOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.w3c.dom.ls.LSOutput;

public class ExcelFileUtils {
	//global object -outside the method but inside the class11q4q
	Workbook wb;
	//constructor for reading excel path
	public ExcelFileUtils(String Excelpath) throws Throwable {
		FileInputStream fi =new FileInputStream(Excelpath);
		wb = WorkbookFactory.create(fi);

	}
	//counting number of rows in sheet
	public int rowCount(String SheetName)
	{
		return wb.getSheet(SheetName).getLastRowNum();
	}

	//read cell data 
	//for non static create reference obj and then call the methods
	public String getCellData(String SheetName,int row,int column) {
		//local method as it is defined inside the method
		String data ="";
		if (wb.getSheet(SheetName).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC)
		{
			//read integer type cell data
			int celldata =(int)wb.getSheet(SheetName).getRow(row).getCell(column).getNumericCellValue();
			System.out.println("Cell data " + celldata);
			//valueof---it converts int type to String type
			data = String.valueOf(celldata);
			System.out.println(" data " + data);
		}
		else
		{
			data = wb.getSheet(SheetName).getRow(row).getCell(column).getStringCellValue();

		}
		return data;

	}
	//method for set cell data
	public void setCelldata(String sheetName,int row,int coloumn,String status,String writeExcel)throws Throwable {

		//get sheet from wb
		Sheet ws =wb.getSheet(sheetName);
		Row rowNum =ws.getRow(row);
		//create cell in a row
		Cell cell=rowNum.createCell(coloumn);
		//write status
		cell.setCellValue(status);
		if(status.equalsIgnoreCase("pass"))
		{
			CellStyle style = wb.createCellStyle();
			Font font =wb.createFont();
			//colour text green
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			style.setFont(font);
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			rowNum.getCell(coloumn).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("fail"))


		{

			CellStyle style = wb.createCellStyle();
			Font font =wb.createFont();
			//colour text green
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			style.setFont(font);
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			rowNum.getCell(coloumn).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("blocked"))

		{
			CellStyle style = wb.createCellStyle();
			Font font =wb.createFont();
			//colour text green
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			style.setFont(font);
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			rowNum.getCell(coloumn).setCellStyle(style);
		}
		FileOutputStream fo =new FileOutputStream(writeExcel);
		wb.write(fo);
	}
	/*public static void main(String[] args) throws Throwable {
		ExcelFileUtils xl =new ExcelFileUtils("D:\\Sample.xlsx");
		int rc =xl.rowCount("Emp");
		System.out.println(rc);
		
		for (int i=1;i<=rc;i++)
		{
			String fname =xl.getCellData("Emp",i, 0);
			String mname =xl.getCellData("Emp",i, 1);
			String lname =xl.getCellData("Emp",i, 2);
			String eid =xl.getCellData("Emp",i, 3);
			System.out.println(fname+ "    "+mname+"    "+lname+"  "+eid);
			//xl.setCelldata("Emp",i, 4, "Pass","D://Result.xlsx" );
			//xl.setCelldata("Emp",i, 4, "Fail","D://Result1.xlsx" );
			xl.setCelldata("Emp",i, 4, "Blocked","D://Result5.xlsx" );
		}
	
		
		
	}*/
}
