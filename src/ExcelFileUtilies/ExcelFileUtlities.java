package ExcelFileUtilies;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ExcelFileUtlities {

Workbook wb;
	
	public ExcelFileUtlities() throws Exception{
		File src=new File("D:\\mahendar\\StockAccountingHybrid\\Testinput\\InputSheet.xlsx");
		FileInputStream fis=new FileInputStream(src);
		 wb=new XSSFWorkbook(fis);
		
	}
	public int rowCount(String sheetname){
		return wb.getSheet(sheetname).getLastRowNum();
	}
	public int colCount(String sheetname)
	{
		return wb.getSheet(sheetname).getRow(0).getLastCellNum();
	}
	public String getData(String sheetname,int row,int column){
		String data="";
		if(wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC)
	
			{
			int celldata=(int)wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
			data=String.valueOf(celldata);
			}
		else{
			data=wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();
		}
		return data;
		
		}
	public void setData(String sheetname,int row,int column, String status) throws Exception{
		Sheet sh=wb.getSheet(sheetname);
		Row rownum=sh.getRow(row);
		Cell cell=rownum.createCell(column);
		cell.setCellValue(status);
		if(status.equalsIgnoreCase("pass")){
			
		
		//Create a cell style
		CellStyle style=wb.createCellStyle();
		//create a font
		Font font=wb.createFont();
		//apply color to the text
		font.setColor(IndexedColors.GREEN.getIndex());
		//apply bold to text
		font.setBold(true);
		style.setFont(font);
		rownum.getCell(column).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("fail")){
			
			
			//Create a cell style
			CellStyle style=wb.createCellStyle();
			//create a font
			Font font=wb.createFont();
			//apply color to the text
			font.setColor(IndexedColors.RED.getIndex());
			//apply bold to text
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
			}
else if(status.equalsIgnoreCase("not executed")){
			
			
			//Create a cell style
			CellStyle style=wb.createCellStyle();
			//create a font
			Font font=wb.createFont();
			//apply color to the text
			font.setColor(IndexedColors.BLUE.getIndex());
			//apply bold to text
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
			}
	FileOutputStream fos=new FileOutputStream("D:\\mahendar\\StockAccountingHybrid\\Testoutput\\hybrid.xlsx");
       wb.write(fos);
       fos.close();
}
}


