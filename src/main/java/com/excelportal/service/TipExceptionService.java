package com.excelportal.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.stream.IntStream;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jboss.logging.Logger;
import org.jboss.logging.Logger.Level;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.excelportal.utility.ExcelUtilityHelper;

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

			Row headerRow = sheet.getRow(3);

			columnNameMap = ExcelUtilityHelper.mapColumnNamesToIndex(headerRow);

			areaCoachMap = getAreaCoachMap(columnNameMap, sheet);			

			for (int indexOfRow = 0; indexOfRow < 3; indexOfRow++) {

				sheet.removeRow(sheet.getRow(indexOfRow));

			}
			
			ExcelUtilityHelper.removeEmptyRows(sheet);
			
			int totalRows = sheet.getPhysicalNumberOfRows();

			for (int indexOfRow = 1; indexOfRow < totalRows; indexOfRow++) {
				
				Row currentRow = sheet.getRow(indexOfRow);

				int indexOfTipColumn = columnNameMap.get("Tip");

				Cell tipCell = currentRow.getCell(indexOfTipColumn);

				if (tipCell.getCellType() == CellType.STRING) {

					sheet.removeRow(currentRow);

					continue;
				}

				double tipValue = tipCell.getNumericCellValue();

				if (tipValue < 10) {

					sheet.removeRow(currentRow);

					continue;
				}

				int indexOfTipPercentColumn = columnNameMap.get("Tip Pct");

				Cell tipPercentCell = currentRow.getCell(indexOfTipPercentColumn);

				double tipPercent = tipPercentCell.getNumericCellValue();

				if ((tipValue >= 20 && tipPercent < 0.5) || ((tipValue >= 10 && tipValue < 20) && tipPercent < 1)) {

					sheet.removeRow(currentRow);

				}

			}

			ExcelUtilityHelper.removeEmptyRows(sheet);

			setAreaCoachColumnValues(areaCoachMap, columnNameMap, sheet);
			
			int lastRowNumberBeforeSorting = sheet.getLastRowNum();

			ExcelUtilityHelper.sortSheet(sheet, columnNameMap);
			
			for (int i = 1; i < lastRowNumberBeforeSorting; i++) {
				
				sheet.removeRow(sheet.getRow(i));
			}
			
			ExcelUtilityHelper.removeEmptyRows(sheet);

			styleCells(sheet, columnNameMap);

			for (int i = 0; i < 4; i++) {

				for (int indexOfColumnToBeDeleted = columnNameMap.get("Card Type"); indexOfColumnToBeDeleted < sheet
						.getRow(0).getPhysicalNumberOfCells(); indexOfColumnToBeDeleted++) {

					ExcelUtilityHelper.deleteColumn(sheet, indexOfColumnToBeDeleted);

				}
			}

			IntStream.range(0, sheet.getRow(0).getPhysicalNumberOfCells())
					.forEach(columnIndex -> sheet.autoSizeColumn(columnIndex));

		} catch (IOException e) {
			e.printStackTrace();
		}

		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

		workbook.write(outputStream);

		return new ByteArrayInputStream(outputStream.toByteArray());
	}

	private void styleCells(Sheet sheet, Map<String, Integer> columnNameMap) {

		String currentAreaCoach;

		String nextAreaCoach;

		for (int indexOfRow = 0; indexOfRow < sheet.getPhysicalNumberOfRows(); indexOfRow++) {

			Row currentRow = sheet.getRow(indexOfRow);

			if (currentRow.getRowNum() == 0) {

				for (int cellIndex = 0; cellIndex < currentRow.getPhysicalNumberOfCells(); cellIndex++) {

					currentRow.getCell(1).setCellValue("Area Coach");

					CellStyle headerStyle = createHeaderStyle();

					if (cellIndex == 0) {

						headerStyle.setBorderLeft(BorderStyle.MEDIUM);
					}

					if (cellIndex == currentRow.getPhysicalNumberOfCells() - 1) {

						headerStyle.setBorderRight(BorderStyle.MEDIUM);

					}

					currentRow.getCell(cellIndex).setCellStyle(headerStyle);

				}

				continue;
			}

			currentAreaCoach = currentRow.getCell(1).getStringCellValue();

			int nextAreaCoachIndex;

			if (indexOfRow == sheet.getPhysicalNumberOfRows() - 1) {

				nextAreaCoachIndex = indexOfRow;

			} else {

				nextAreaCoachIndex = indexOfRow + 1;

			}

			nextAreaCoach = sheet.getRow(nextAreaCoachIndex).getCell(1).getStringCellValue();

			LOGGER.log(Level.INFO, nextAreaCoach);

			int indexOfTipColumn = columnNameMap.get("Tip");

			Cell tipCell = currentRow.getCell(indexOfTipColumn);

			double tipValue = tipCell.getNumericCellValue();

			if (tipValue >= 10 && tipValue < 20) {

				IndexedColors color = IndexedColors.AQUA;

				CellStyle storeStyle = createStoreCellStyle();

				if (!currentAreaCoach.equals(nextAreaCoach)) {

					storeStyle.setBorderBottom(BorderStyle.MEDIUM);
				}

				currentRow.getCell(0).setCellStyle(storeStyle);

				for (int indexOfFirstNonStoreCell = 1; indexOfFirstNonStoreCell < currentRow
						.getPhysicalNumberOfCells(); indexOfFirstNonStoreCell++) {

					if (!(indexOfFirstNonStoreCell == columnNameMap.get("Tip Pct"))) {

						if (!(indexOfFirstNonStoreCell == columnNameMap.get("Business Date"))) {

							CellStyle nonStoreStyle = createNonStoreCellStyle();

							if (!currentAreaCoach.equals(nextAreaCoach)) {

								nonStoreStyle.setBorderBottom(BorderStyle.MEDIUM);
							}

							currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(nonStoreStyle);

						} else {

							CellStyle businessDateStyle = createBusinessDateCellStyle();

							if (!currentAreaCoach.equals(nextAreaCoach)) {

								businessDateStyle.setBorderBottom(BorderStyle.MEDIUM);
							}

							currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(businessDateStyle);
						}

					} else {

						CellStyle tipPercentageStyle = createTipPercentageCellStyle();

						if (!currentAreaCoach.equals(nextAreaCoach)) {

							tipPercentageStyle.setBorderBottom(BorderStyle.MEDIUM);
						}

						currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(tipPercentageStyle);
					}
				}

			} else {

				IndexedColors color = IndexedColors.YELLOW;

				CellStyle storeStyle = createStoreCellStyle();

				storeStyle.setFillForegroundColor(color.getIndex());

				if (!currentAreaCoach.equals(nextAreaCoach)) {

					storeStyle.setBorderBottom(BorderStyle.MEDIUM);

				}

				currentRow.getCell(0).setCellStyle(storeStyle);

				for (int indexOfFirstNonStoreCell = 1; indexOfFirstNonStoreCell < currentRow
						.getPhysicalNumberOfCells(); indexOfFirstNonStoreCell++) {

					if (!(indexOfFirstNonStoreCell == columnNameMap.get("Tip Pct"))) {

						if (!(indexOfFirstNonStoreCell == columnNameMap.get("Business Date"))) {

							CellStyle nonStoreStyle = createNonStoreCellStyle();

							nonStoreStyle.setFillForegroundColor(color.getIndex());

							if (!currentAreaCoach.equals(nextAreaCoach)) {

								nonStoreStyle.setBorderBottom(BorderStyle.MEDIUM);

							}

							currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(nonStoreStyle);

						} else {

							CellStyle businessDateStyle = createBusinessDateCellStyle();

							businessDateStyle.setFillForegroundColor(color.getIndex());

							if (!currentAreaCoach.equals(nextAreaCoach)) {

								businessDateStyle.setBorderBottom(BorderStyle.MEDIUM);
							}

							currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(businessDateStyle);
						}

					} else {

						CellStyle tipPercentageStyle = createTipPercentageCellStyle();

						tipPercentageStyle.setFillForegroundColor(color.getIndex());

						if (!currentAreaCoach.equals(nextAreaCoach)) {

							tipPercentageStyle.setBorderBottom(BorderStyle.MEDIUM);

						}

						currentRow.getCell(indexOfFirstNonStoreCell).setCellStyle(tipPercentageStyle);
					}
				}
			}
		}

	}

	private CellStyle createStoreCellStyle() {

		CellStyle style = workbook.createCellStyle();

		style.setBorderLeft(BorderStyle.MEDIUM);

		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		if (style instanceof XSSFCellStyle) {

			XSSFCellStyle xssfcellcolorstyle = (XSSFCellStyle) style;

			xssfcellcolorstyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(56, 163, 237)));

		}

		return style;
	}

	private CellStyle createBusinessDateCellStyle() {

		CellStyle style = workbook.createCellStyle();

		CreationHelper helper = workbook.getCreationHelper();

		style.setDataFormat(helper.createDataFormat().getFormat("mm/dd/yyyy"));

		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		style.setAlignment(HorizontalAlignment.CENTER);

		if (style instanceof XSSFCellStyle) {

			XSSFCellStyle xssfcellcolorstyle = (XSSFCellStyle) style;

			xssfcellcolorstyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(56, 163, 237)));

		}

		return style;
	}

	private CellStyle createNonStoreCellStyle() {

		CellStyle style = workbook.createCellStyle();

		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		style.setAlignment(HorizontalAlignment.CENTER);

		if (style instanceof XSSFCellStyle) {

			XSSFCellStyle xssfcellcolorstyle = (XSSFCellStyle) style;

			xssfcellcolorstyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(56, 163, 237)));

		}

		return style;
	}

	private CellStyle createTipPercentageCellStyle() {

		CellStyle style = workbook.createCellStyle();

		style.setBorderRight(BorderStyle.MEDIUM);

		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		style.setAlignment(HorizontalAlignment.CENTER);

		style.setDataFormat(workbook.createDataFormat().getFormat("00.00%"));

		if (style instanceof XSSFCellStyle) {

			XSSFCellStyle xssfcellcolorstyle = (XSSFCellStyle) style;

			xssfcellcolorstyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(56, 163, 237)));

		}

		return style;
	}

	private CellStyle createHeaderStyle() {

		CellStyle style = workbook.createCellStyle();

		Font font = workbook.createFont();

		font.setFontHeightInPoints((short) 11);

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

		for (int indexOfRow = 5; indexOfRow < totalRows; indexOfRow++) {

			Row currentRow = sheet.getRow(indexOfRow);

			Cell businessDateCell = currentRow.getCell(columnNameMap.get("Business Date"));

			if (businessDateCell.getCellType() == CellType.BLANK) {

				areaCoachName = currentRow.getCell(columnNameMap.get("Store")).getStringCellValue();

				coachStoreMap.put(areaCoachName, new ArrayList<String>());
			}

			coachStoreMap.get(areaCoachName).add(currentRow.getCell(columnNameMap.get("Store")).getStringCellValue());
		}

		return coachStoreMap;
	}

	private void setAreaCoachColumnValues(Map<String, List<String>> areaCoachMap, Map<String, Integer> columnNameMap,
			Sheet sheet) {

		for (int indexOfRow = 1; indexOfRow < sheet.getPhysicalNumberOfRows(); indexOfRow++) {

			Row currentRow = sheet.getRow(indexOfRow);

			Cell storeCell = currentRow.getCell(columnNameMap.get("Store"));

			for (Entry<String, List<String>> entrySet : areaCoachMap.entrySet()) {

				if (entrySet.getValue().contains(storeCell.getStringCellValue())) {

					String nameKey = entrySet.getKey();

					if (nameKey.contains("Andre")) {
						currentRow.getCell(1).setCellValue("Arlee");

					} else if (nameKey.contains("Hill")) {
						currentRow.getCell(1).setCellValue("Hannah");

					} else if (nameKey.contains("Jordan")) {
						currentRow.getCell(1).setCellValue("Harvey");

					} else if (nameKey.contains("Holden")) {
						currentRow.getCell(1).setCellValue("Joe");

					} else if (nameKey.contains("Welsh")) {
						currentRow.getCell(1).setCellValue("Jen");

					} else if (nameKey.contains("Johnson")) {
						currentRow.getCell(1).setCellValue("Monica");

					} else if (nameKey.contains("Lewis")) {
						currentRow.getCell(1).setCellValue("Michelle");

					} else if (nameKey.contains("Comer")) {
						currentRow.getCell(1).setCellValue("Robert");

					} else if (nameKey.contains("Sanchez")) {
						currentRow.getCell(1).setCellValue("Rachael");

					} else if (nameKey.contains("Traill")) {
						currentRow.getCell(1).setCellValue("Rumone");

					} else if (nameKey.contains("Evans")) {
						currentRow.getCell(1).setCellValue("Theresa");

					} else if (nameKey.contains("Shreve")) {
						currentRow.getCell(1).setCellValue("Barbara");

					} else if (nameKey.contains("Hauert")) {
						currentRow.getCell(1).setCellValue("Brandon");

					} else if (nameKey.contains("Berrios")) {
						currentRow.getCell(1).setCellValue("Hector");
					}
				}
			}
		}
	}

}
