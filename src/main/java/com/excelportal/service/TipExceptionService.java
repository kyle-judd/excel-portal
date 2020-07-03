package com.excelportal.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.concurrent.ConcurrentHashMap.KeySetView;
import java.util.stream.IntStream;

import javax.persistence.Column;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
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
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jboss.logging.Logger;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.excelportal.utility.ExcelUtilityHelper;

import ch.qos.logback.classic.db.names.ColumnName;

@Service
public class TipExceptionService {
	
	private XSSFWorkbook workbook;
	
	private final Logger LOGGER = Logger.getLogger(this.getClass());
	
	public ByteArrayInputStream parseTipException(MultipartFile tipExceptionReport) throws IOException {
		
		Map<String, Integer> columnNameMap;
		
		Map<String, List<String>> areaCoachMap;
		
		try {
			
			workbook = new XSSFWorkbook(tipExceptionReport.getInputStream());
			
			Sheet sheet = workbook.getSheetAt(0);
			
			ExcelUtilityHelper.insertNewColumnBefore(workbook, 0, 1);
			
			sheet.getRow(0).getCell(1).setCellValue("Area Coach");

			Row headerRow = sheet.getRow(3);
			
			columnNameMap = ExcelUtilityHelper.mapColumnNamesToIndex(headerRow);
			
			areaCoachMap = getAreaCoachMap(columnNameMap, sheet);

			int totalRows = sheet.getPhysicalNumberOfRows();
			
			for(int indexOfRow = 4; indexOfRow < totalRows; indexOfRow++) {
				
				Row currentRow = sheet.getRow(indexOfRow);
								
				if(currentRow.getRowNum() == 4) {
		
					sheet.removeRow(currentRow);
					
					continue;
				}
				
				int indexOfTipColumn = columnNameMap.get("Tip");
				
				Cell tipCell = currentRow.getCell(indexOfTipColumn);
				
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
			
			for(int indexOfRow = 0; indexOfRow < 3; indexOfRow++) {

				sheet.removeRow(sheet.getRow(indexOfRow));
				
			}
			
			ExcelUtilityHelper.removeEmptyRows(sheet);
			
			setAreaCoachColumnValues(areaCoachMap, columnNameMap, sheet);
			
			styleCells(sheet, columnNameMap);
			
			for(int i = 0; i < 4; i++) {
				
				for(int indexOfColumnToBeDeleted = columnNameMap.get("Card Type"); indexOfColumnToBeDeleted < sheet.getRow(0).getPhysicalNumberOfCells(); indexOfColumnToBeDeleted++) {
					
					ExcelUtilityHelper.deleteColumn(sheet, indexOfColumnToBeDeleted);
					
				}
			}
			
			
			IntStream.range(0, sheet.getRow(0).getPhysicalNumberOfCells()).forEach(columnIndex -> sheet.autoSizeColumn(columnIndex));
	
			
		} catch(IOException e) {
			e.printStackTrace();
		}
		
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		
		workbook.write(outputStream);
		
		return new ByteArrayInputStream(outputStream.toByteArray());
	}
	
	private void styleCells(Sheet sheet, Map<String, Integer> columnNameMap) {
		
		for(int indexOfRow = 0; indexOfRow < sheet.getPhysicalNumberOfRows(); indexOfRow++) {
			
			Row currentRow = sheet.getRow(indexOfRow);
			
			if(currentRow.getRowNum() == 0) {
				
				for(int cellIndex = 0; cellIndex < currentRow.getPhysicalNumberOfCells(); cellIndex++) {

					currentRow.getCell(cellIndex).setCellStyle(createHeaderStyle());
					
				}
				
				continue;
			}
			
			int indexOfTipColumn = columnNameMap.get("Tip");
			
			Cell tipCell = currentRow.getCell(indexOfTipColumn);
			
			double tipValue = tipCell.getNumericCellValue();
			
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
	
	private CellStyle createHeaderStyle() {
		
		CellStyle style = workbook.createCellStyle();
		
		Font font = workbook.createFont();
		
		font.setFontHeightInPoints((short) 15);
		
		font.setBold(true);
		
		style.setFont(font);
		
		style.setAlignment(HorizontalAlignment.CENTER);
		
		style.setBorderBottom(BorderStyle.MEDIUM);
		
		return style;
	}
	
	private Map<String, List<String>> getAreaCoachMap(Map<String, Integer> columnNameMap, Sheet sheet) {
		
		Map<String, List<String>> coachStoreMap = new HashMap<>();
		
		int totalRows = sheet.getPhysicalNumberOfRows();
		
		Row firstAreaCoachRow = sheet.getRow(4);
		
		String areaCoachName = firstAreaCoachRow.getCell(columnNameMap.get("Store")).getStringCellValue();
		
		coachStoreMap.put(areaCoachName, new ArrayList<String>());
		
		for(int indexOfRow = 5; indexOfRow < totalRows; indexOfRow++) {
		
			Row currentRow = sheet.getRow(indexOfRow);
			
			Cell businessDateCell = currentRow.getCell(columnNameMap.get("Business Date"));
			
			if(businessDateCell.getCellType() == CellType.BLANK) {
			
				areaCoachName = currentRow.getCell(columnNameMap.get("Store")).getStringCellValue();
				
				coachStoreMap.put(areaCoachName, new ArrayList<String>());
			}
			
			coachStoreMap.get(areaCoachName).add(currentRow.getCell(columnNameMap.get("Store")).getStringCellValue());
		}
		
		return coachStoreMap;
	}
	
	private void setAreaCoachColumnValues(Map<String, List<String>> areaCoachMap, Map<String, Integer> columnNameMap, Sheet sheet) {
		
		for(int indexOfRow = 1; indexOfRow < sheet.getPhysicalNumberOfRows(); indexOfRow++) {
		
			Row currentRow = sheet.getRow(indexOfRow);
			
			Cell storeCell = currentRow.getCell(columnNameMap.get("Store"));

			for(Entry<String, List<String>> entrySet : areaCoachMap.entrySet()) {

				if(entrySet.getValue().contains(storeCell.getStringCellValue())) {
				
					String nameKey = entrySet.getKey();
					
					if(nameKey.contains("Andre")) {
						currentRow.getCell(1).setCellValue("Arlee");
					
					} else if(nameKey.contains("Hill")) {
						currentRow.getCell(1).setCellValue("Hannah");
					
					} else if(nameKey.contains("Jordan")) {
						currentRow.getCell(1).setCellValue("Harvey");
					
					} else if(nameKey.contains("Holden")) {
						currentRow.getCell(1).setCellValue("Joe");
					
					} else if(nameKey.contains("Welsh")) {
						currentRow.getCell(1).setCellValue("Jen");
					
					} else if(nameKey.contains("Johnson")) {
						currentRow.getCell(1).setCellValue("Monica");
					
					} else if(nameKey.contains("Lewis")) {
						currentRow.getCell(1).setCellValue("Michelle");
					
					} else if(nameKey.contains("Comer")) {
						currentRow.getCell(1).setCellValue("Robert");
					
					} else if(nameKey.contains("Sanchez")) {
						currentRow.getCell(1).setCellValue("Rachel");
					
					} else if(nameKey.contains("Traill")) {
						currentRow.getCell(1).setCellValue("Rumone");
					
					} else if(nameKey.contains("Evans")) {
						currentRow.getCell(1).setCellValue("Theresa");
					
					} else if(nameKey.contains("Shreve")) {
						currentRow.getCell(1).setCellValue("Barbara");
					
					} else if(nameKey.contains("Hauert")) {
						currentRow.getCell(1).setCellValue("Brandon");
					
					} else if(nameKey.contains("Berrios")) {
						currentRow.getCell(1).setCellValue("Hector");
					}
				}
			}
		}
	}

}
