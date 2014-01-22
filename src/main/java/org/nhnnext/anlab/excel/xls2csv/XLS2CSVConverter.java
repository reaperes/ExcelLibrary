package org.nhnnext.anlab.excel.xls2csv;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class XLS2CSVConverter {
	public void convertToCSV(String sourceFile, int sheetIndex, String targetFile) throws Exception {
		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File(sourceFile)));
		XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
		
		StringBuffer buffer = new StringBuffer();
		
		int maxRow = sheet.getPhysicalNumberOfRows()-1;
		for (int i=0; i<maxRow; i++) {
			XSSFRow row = sheet.getRow(i);
			
			int maxCell = row.getPhysicalNumberOfCells();
			for (int j=0; j<maxCell; j++) {
				XSSFCell cell = row.getCell(j);

				switch (cell.getCellType()) {
				case XSSFCell.CELL_TYPE_STRING:
					buffer.append(cell.getStringCellValue());
					break;
					
				case XSSFCell.CELL_TYPE_NUMERIC:
					buffer.append(Math.round(cell.getNumericCellValue()));
					break;
					
				default:
					new IllegalArgumentException(String.valueOf(cell.getCellType()));
				}
				
				if (j != maxCell-1)
					buffer.append("$");
			}
			
			if (i != maxCell-1)
				buffer.append(System.lineSeparator());
		}
		
		Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(targetFile), "UTF8"));
		writer.write(buffer.toString());
		writer.flush();
		writer.close();
	}
}
