package com.excelportal.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jboss.logging.Logger;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class DriverDataService {
	
	private final Logger LOGGER = Logger.getLogger(getClass());
	
	public ByteArrayInputStream parseForDriverOverrideMiles(MultipartFile excelFile) throws IOException {
		
		// workbook to use
		XSSFWorkbook workbook = null;
		
		// map to keep track of how many times a driver appears on the list
		Map<String, Integer> driverMap = new HashMap<>();
		
		// create empty map to later store column names
		Map<String, Integer> columnNameMap;
		
		try {
			// read file into workbook
			workbook = new XSSFWorkbook(excelFile.getInputStream());
			if(workbook != null) {
				LOGGER.info("Workbook received!!!!!!!!");
			}
			// we know this workbook will only have one sheet, so get it by using index of 0
			Sheet sheet = workbook.getSheetAt(0);
			
			
			// get the first row to be used to map the column names
			Row firstRow = sheet.getRow(0);
			
			// then map the column names of original sheet using our method
			columnNameMap = mapColumnNamesToIndex(firstRow);
			

			// total number of rows to use in our loop
			int totalRows = sheet.getPhysicalNumberOfRows();
			
			// loop through the rows
			for(int indexOfRow = 1; indexOfRow < totalRows; indexOfRow++) {				
				// current row in the loop
				Row currentRow = sheet.getRow(indexOfRow);	
				
				// we only care about drivers with system miles > 0 so start there
				int indexOfSystemMilesColumn = columnNameMap.get("System Miles");
				
				Cell systemMilesCell = currentRow.getCell(indexOfSystemMilesColumn);
				
				if(systemMilesCell.getNumericCellValue() == 0) {
					sheet.removeRow(currentRow);
				} else {
					// get the name of the driver
					int indexOfDriverNameColumn = columnNameMap.get("Driver Name");
					
					Cell driverNameCell = currentRow.getCell(indexOfDriverNameColumn);
					
					String driverName = driverNameCell.getStringCellValue();
					
					// check if the map already contains the name of the driver, then either add the new driver to the map or update the driver's count
					if(!driverMap.containsKey(driverName)) {
						driverMap.put(driverName, 1);
					} else {
						driverMap.put(driverName, driverMap.get(driverName) + 1);
					}
				}
	
			}
			
			for(Map.Entry<String, Integer> entry : driverMap.entrySet()) {
				LOGGER.info("DRIVER MAP =========> " + entry);
			}
			
			// filter drivers
			filterForDriversWithOccurencesOfThreeOrMore(sheet, driverMap, columnNameMap);
			
			// now remove empty rows
			removeEmptyRows(sheet);
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		workbook.write(outputStream);
		return new ByteArrayInputStream(outputStream.toByteArray());
	}
	
	private Map<String, Integer> mapColumnNamesToIndex(Row row) {
		Map<String, Integer> columnNameMap = new HashMap<>();
		short minColumnIndex = row.getFirstCellNum();
		short maxColumnIndex = row.getLastCellNum();
		for(short colIndex = minColumnIndex; colIndex < maxColumnIndex; colIndex++) {
			Cell currentCell = row.getCell(colIndex);
			columnNameMap.put(currentCell.getStringCellValue(), currentCell.getColumnIndex());
		}
		return columnNameMap;
	}
	
	private void filterForDriversWithOccurencesOfThreeOrMore(Sheet sheet, Map<String, Integer> driverMap, Map<String, Integer> columnNameMap) {
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
	
	private void removeEmptyRows(Sheet sheet) {
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
