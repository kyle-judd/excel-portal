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

		XSSFWorkbook workbook = null;

		Map<String, Integer> driverMap = new HashMap<>();

		Map<String, Integer> columnNameMap;

		try {

			workbook = new XSSFWorkbook(excelFile.getInputStream());

			Sheet sheet = workbook.getSheetAt(0);

			Row firstRow = sheet.getRow(0);

			int totalRows = sheet.getLastRowNum();

			columnNameMap = ExcelUtilityHelper.mapColumnNamesToIndex(firstRow);

			for (int indexOfRow = 1; indexOfRow < totalRows; indexOfRow++) {

				Row currentRow = sheet.getRow(indexOfRow);

				if (currentRow.getCell(columnNameMap.get("System Miles")).getNumericCellValue() != 0) {

					Cell driverNameCell = currentRow.getCell(columnNameMap.get("Driver Name"));

					String driverName = driverNameCell.getStringCellValue();

					if (!driverMap.containsKey(driverName)) {

						driverMap.put(driverName, 1);

					} else {

						driverMap.put(driverName, driverMap.get(driverName) + 1);
					}

				} else {

					sheet.removeRow(currentRow);

				}
			}
			
			filterForDriversWithOccurencesOfThreeOrMore(sheet, driverMap, columnNameMap);
			
			ExcelUtilityHelper.removeEmptyRows(sheet);

		} catch (IOException e) {
			e.printStackTrace();
		}

		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		workbook.write(outputStream);
		return new ByteArrayInputStream(outputStream.toByteArray());
	}
	
	private void filterForDriversWithOccurencesOfThreeOrMore(Sheet sheet, Map<String, Integer> driverMap,
			Map<String, Integer> columnNameMap) {
		
		for (int indexOfRow = 1; indexOfRow < sheet.getLastRowNum(); indexOfRow++) {
			
			Row currentRow = sheet.getRow(indexOfRow);
			
			if (currentRow == null) {
				
				continue;
				
			}
			
			Cell driverNameCell = currentRow.getCell(columnNameMap.get("Driver Name"));
			
			String driverName = driverNameCell.getStringCellValue();
			
			if (driverMap.get(driverName) < 3) {
				
				sheet.removeRow(currentRow);
			}
		}
	}

}


