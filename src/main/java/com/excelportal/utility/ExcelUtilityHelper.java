package com.excelportal.utility;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ExcelUtilityHelper {
	
	public static Map<String, Integer> mapColumnNamesToIndex(Row row) {
		Map<String, Integer> columnNameMap = new HashMap<>();
		short minColumnIndex = row.getFirstCellNum();
		short maxColumnIndex = row.getLastCellNum();
		for(short colIndex = minColumnIndex; colIndex < maxColumnIndex; colIndex++) {
			Cell currentCell = row.getCell(colIndex);
			columnNameMap.put(currentCell.getStringCellValue(), currentCell.getColumnIndex());
		}
		return columnNameMap;
	}
	
	public static void filterForDriversWithOccurencesOfThreeOrMore(Sheet sheet, Map<String, Integer> driverMap, Map<String, Integer> columnNameMap) {
		for(int rowIndex = 1; rowIndex < sheet.getLastRowNum(); rowIndex++) {
			if(sheet.getRow(rowIndex) == null) {
				continue;
			} else {
				Row currentRow = sheet.getRow(rowIndex);
				int indexOfDriverNameColumn = columnNameMap.get("Driver Name");
				Cell driverNameCell = currentRow.getCell(indexOfDriverNameColumn);
				String driverName = driverNameCell.getStringCellValue();
				if(driverMap.get(driverName) < 3) {
					sheet.removeRow(currentRow);
				}
			}
		}
	}
	
	public static void removeEmptyRows(Sheet sheet) {
		Boolean isRowEmpty = Boolean.FALSE;
		for (int i = 0; i <= sheet.getLastRowNum(); i++) {
			if (sheet.getRow(i) == null) {
				isRowEmpty = true;
				sheet.shiftRows(i + 1, sheet.getLastRowNum() + 1, -1);
				i--;
				continue;
			}
			for (int j = 0; j < sheet.getRow(i).getLastCellNum(); j++) {
				if (sheet.getRow(i).getCell(j) == null || sheet.getRow(i).getCell(j).toString().trim().equals("")) {
					isRowEmpty = true;
				} else {
					isRowEmpty = false;
					break;
				}
			}
			if (isRowEmpty == true) {
				sheet.shiftRows(i + 1, sheet.getLastRowNum() + 1, -1);
				i--;
			}
		}
	}
}
