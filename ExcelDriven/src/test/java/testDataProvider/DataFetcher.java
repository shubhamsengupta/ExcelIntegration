package testDataProvider;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataFetcher {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		FileInputStream fs = new FileInputStream("D:/Coding/Eclipse/Workspace/ExcelIntegration/TestDataDemo.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fs);
		XSSFSheet sheet = workbook.getSheet("DataSheet");
		int rows = sheet.getLastRowNum()+1;
		short top = sheet.getTopRow();
		XSSFRow rowHeader = sheet.getRow(top);
		int cols = rowHeader.getLastCellNum();
		for(int i=0; i<rows; i++) {
			XSSFRow row = sheet.getRow(i);
			for(int j=0; j<cols; j++) {
				XSSFCell cell = row.getCell(j);
				System.out.print(cell.getStringCellValue()+"\t");
			}
			System.out.println("");
		}
		workbook.close();

	}

}
