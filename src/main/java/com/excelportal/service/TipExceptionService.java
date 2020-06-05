package com.excelportal.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jboss.logging.Logger;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.excelportal.utility.ExcelUtilityHelper;

@Service
public class TipExceptionService {
	
	// highlight yellow if $20 or more and greater than 50% of the ticket
	
	// hightlight blue if $10 > or more and greater than 100% of the ticket
	
	// add area coach
	
	// sort by area coach
	
	// then sort by store
	
	private final Logger LOGGER = Logger.getLogger(this.getClass());
	
	public ByteArrayInputStream parseTipException(MultipartFile tipExceptionReport) throws IOException {
		
		XSSFWorkbook workbook = null;
		
		Map<String, Integer> columnNameMap = new HashMap<>();
		
		try {
			
			workbook = new XSSFWorkbook(tipExceptionReport.getInputStream());
			
			Sheet sheet = workbook.getSheetAt(0);
			
			Row headerRow = sheet.getRow(3);
			
			columnNameMap = ExcelUtilityHelper.mapColumnNamesToIndex(headerRow);
			
			int totalRows = sheet.getPhysicalNumberOfRows();
			
			for(int indexOfRow = 5; indexOfRow < totalRows; indexOfRow++) {
				
				Row currentRow = sheet.getRow(indexOfRow);
				
				int indexOfTipColumn = columnNameMap.get("Tip");
				
				Cell tipCell = currentRow.getCell(indexOfTipColumn);
				
				// look for a better way to do this
				if(tipCell.getCellType() == CellType.STRING) {
					sheet.removeRow(currentRow);
					continue;
				}
				
				double tipValue = tipCell.getNumericCellValue();
	
				if(tipValue < 10) {
					sheet.removeRow(currentRow);
					continue;
				}
				
				LOGGER.warn("TIP VALUE IS ======> " + tipValue);
				
				int indexOfTipPercentColumn = columnNameMap.get("Tip Pct");
				
				Cell tipPercentCell = currentRow.getCell(indexOfTipPercentColumn);
				
				double tipPercent = tipPercentCell.getNumericCellValue();
				
				LOGGER.warn("TIP PERCENT IS =====> " + tipPercent);
				
				if((tipValue >= 20 && tipPercent < 0.5) || ((tipValue >= 10 && tipValue < 20) && tipPercent < 1)) {
					sheet.removeRow(currentRow);
				}

			}
	
			ExcelUtilityHelper.removeEmptyRows(sheet);
			
			int totalRowsAfterRemovingEmptyRows = sheet.getPhysicalNumberOfRows();
			
			for(int indexOfRow = 5; indexOfRow < totalRowsAfterRemovingEmptyRows; indexOfRow++) {
				
				Row currentRow = sheet.getRow(indexOfRow);
				
				int indexOfTipColumn = columnNameMap.get("Tip");
				
				Cell tipCell = currentRow.getCell(indexOfTipColumn);
				
				double tipValue = tipCell.getNumericCellValue();

				if(tipValue >= 10 && tipValue < 20) {
					CellStyle backgroundStyle = workbook.createCellStyle();
					CellStyle alignmentStyle = workbook.createCellStyle();
					alignmentStyle.setAlignment(HorizontalAlignment.CENTER);
					backgroundStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
					backgroundStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					for(int indexOfCell = 0; indexOfCell < currentRow.getPhysicalNumberOfCells(); indexOfCell++) {
						currentRow.getCell(indexOfCell).setCellStyle(backgroundStyle);
						for(int nonStoreCell = 1; nonStoreCell < currentRow.getPhysicalNumberOfCells(); nonStoreCell++) {
							currentRow.getCell(nonStoreCell).setCellStyle(alignmentStyle);
						}
					}
					
				} else {
					CellStyle style = workbook.createCellStyle();
					style.setAlignment(HorizontalAlignment.CENTER);
					style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					for(int indexOfCell = 0; indexOfCell < currentRow.getPhysicalNumberOfCells(); indexOfCell++) {
						currentRow.getCell(indexOfCell).setCellStyle(style);
					}
				}
			}
			
		} catch(Exception e) {
			e.printStackTrace();
		}
		
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		workbook.write(outputStream);
		return new ByteArrayInputStream(outputStream.toByteArray());
	}
}
