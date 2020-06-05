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

import com.excelportal.utility.ExcelUtilityHelper;

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
			columnNameMap = ExcelUtilityHelper.mapColumnNamesToIndex(firstRow);
			

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
			ExcelUtilityHelper.filterForDriversWithOccurencesOfThreeOrMore(sheet, driverMap, columnNameMap);
			
			// now remove empty rows
			ExcelUtilityHelper.removeEmptyRows(sheet);
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		workbook.write(outputStream);
		return new ByteArrayInputStream(outputStream.toByteArray());
	}
	
}
