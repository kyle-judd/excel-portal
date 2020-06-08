package com.excelportal.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import javax.persistence.Column;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
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
			// read multipart file into workbook
			workbook = new XSSFWorkbook(tipExceptionReport.getInputStream());
			
			// tip exception just has one sheet
			Sheet sheet = workbook.getSheetAt(0);
			
			// row where column headers are located
			Row headerRow = sheet.getRow(3);
			
			// map the column header string to their indexes for easy reference
			columnNameMap = ExcelUtilityHelper.mapColumnNamesToIndex(headerRow);
			
			// total rows to loop through
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
			
			// remove title rows
			for(int indexOfRow = 0; indexOfRow < 3; indexOfRow++) {
				sheet.removeRow(sheet.getRow(indexOfRow));
			}
			
			// remove the empty rows
			ExcelUtilityHelper.removeEmptyRows(sheet);
			
			int totalRowsAfterRemovingEmptyRows = sheet.getPhysicalNumberOfRows();		
			
			// styles all the cells accordingly
			for(int indexOfRow = 2; indexOfRow < totalRowsAfterRemovingEmptyRows; indexOfRow++) {
				
				Row currentRow = sheet.getRow(indexOfRow);
				
				LOGGER.warn("BEFORE STYLING BUSINESS DATE =====> " + currentRow.getCell(columnNameMap.get("Business Date")).getNumericCellValue());
				
				int indexOfTipColumn = columnNameMap.get("Tip");
				
				Cell tipCell = currentRow.getCell(indexOfTipColumn);
				
				double tipValue = tipCell.getNumericCellValue();

				styleCells(currentRow, tipValue, columnNameMap);
				
			}
			
			for(int i = 0; i < 4; i++) {
				for(int indexOfColumnToBeDeleted = columnNameMap.get("Card Type"); indexOfColumnToBeDeleted < sheet.getRow(0).getPhysicalNumberOfCells(); indexOfColumnToBeDeleted++) {
					ExcelUtilityHelper.deleteColumn(sheet, indexOfColumnToBeDeleted);
				}
			}
			
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
					if(!(indexOfFirstNonStoreCell == columnNameMap.get("Business Date"))) {
						CellStyle nonStoreStyle = createNonStoreCellStyle(color);
						currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(nonStoreStyle);
					} else {
						CellStyle businessDateStyle = createBusinessDateCellStyle(color);
						currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(businessDateStyle);
					}
					
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
					if(!(indexOfFirstNonStoreCell == columnNameMap.get("Business Date"))) {
						CellStyle nonStoreStyle = createNonStoreCellStyle(color);
						currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(nonStoreStyle);
					} else {
						CellStyle businessDateStyle = createBusinessDateCellStyle(color);
						currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(businessDateStyle);
					}
					
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
	
	private CellStyle createBusinessDateCellStyle(IndexedColors color) {
		CellStyle style = workbook.createCellStyle();
		CreationHelper helper = workbook.getCreationHelper();
		style.setDataFormat(helper.createDataFormat().getFormat("mm/dd/yyyy"));
		style.setFillForegroundColor(color.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
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
