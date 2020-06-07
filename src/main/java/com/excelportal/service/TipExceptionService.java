package com.excelportal.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.HashMap;
import java.util.Map;

import javax.persistence.Column;

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

import ch.qos.logback.classic.db.names.ColumnName;

@Service
public class TipExceptionService {
	
	// highlight yellow if $20 or more and greater than 50% of the ticket
	
	// hightlight blue if $10 > or more and greater than 100% of the ticket
	
	// add area coach
	
	// sort by area coach
	
	// then sort by store
	
	private XSSFWorkbook workbook;
	
	private final Logger LOGGER = Logger.getLogger(this.getClass());
	
	public ByteArrayInputStream parseTipException(MultipartFile tipExceptionReport) throws IOException {
		
		
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

				int indexOfTipPercentColumn = columnNameMap.get("Tip Pct");
				
				Cell tipPercentCell = currentRow.getCell(indexOfTipPercentColumn);
				
				double tipPercent = tipPercentCell.getNumericCellValue();
				
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

				styleCells(currentRow, tipValue, columnNameMap);
				
			}
			
			/*for(int indexOfRow = 5; indexOfRow < totalRowsAfterRemovingEmptyRows; indexOfRow++) {
				Row currentRow = sheet.getRow(indexOfRow);
				int indexOfTipPercentColumn = columnNameMap.get("Tip Pct");
				Cell tipPercentCell = currentRow.getCell(indexOfTipPercentColumn);
				tipPercentCell.setCellValue(round(tipPercentCell.getNumericCellValue() * 100, 2));
				CellStyle style = workbook.createCellStyle();
				style.setDataFormat(workbook.createDataFormat().getFormat("0.000%"));
				tipPercentCell.setCellStyle(style);
			}*/
			
		} catch(Exception e) {
			e.printStackTrace();
		}
		
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		workbook.write(outputStream);
		return new ByteArrayInputStream(outputStream.toByteArray());
	}
	
	private void styleCells(Row currentRow, double tipValue, Map<String, Integer> columnNameMap) {
		if(tipValue >= 10 && tipValue < 20) {
			IndexedColors color = IndexedColors.PALE_BLUE;
			CellStyle storeStyle = createStoreCellStyle(color);
			currentRow.getCell(0).setCellStyle(storeStyle);
			for(int indexOfFirstNonStoreCell = 1; indexOfFirstNonStoreCell < currentRow.getPhysicalNumberOfCells(); indexOfFirstNonStoreCell++) {
				if(!(indexOfFirstNonStoreCell == columnNameMap.get("Tip Pct"))) {
					CellStyle nonStoreStyle = createNonStoreCellStyle(color);
					currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(nonStoreStyle);
				} else {					
					CellStyle tipPercentageStyle = createTipPercentageCellStyle(color);
					currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(tipPercentageStyle);
				}			
			}
			
		} else {
			IndexedColors color = IndexedColors.YELLOW;
			CellStyle storeStyle = createStoreCellStyle(color);
			currentRow.getCell(0).setCellStyle(storeStyle);
			for(int indexOfFirstNonStoreCell = 1; indexOfFirstNonStoreCell < currentRow.getPhysicalNumberOfCells(); indexOfFirstNonStoreCell++) {
				if(!(indexOfFirstNonStoreCell == columnNameMap.get("Tip Pct"))) {
					CellStyle nonStoreStyle = createNonStoreCellStyle(color);
					currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(nonStoreStyle);
				} else {					
					CellStyle tipPercentageStyle = createTipPercentageCellStyle(color);
					currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(tipPercentageStyle);
				}	
			}
		}
	}
	
	private CellStyle createStoreCellStyle(IndexedColors color) {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(color.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		return style;
	}
	
	private CellStyle createNonStoreCellStyle(IndexedColors color) {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(color.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		return style;
	}
	
	private CellStyle createTipPercentageCellStyle(IndexedColors color) {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(color.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setDataFormat(workbook.createDataFormat().getFormat("00.00%"));
		return style;
	}

}
