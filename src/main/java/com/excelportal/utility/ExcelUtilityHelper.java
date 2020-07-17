package com.excelportal.utility;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.compress.utils.Lists;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jboss.logging.Logger;

public class ExcelUtilityHelper {

	private static final char LETTERS_IN_EN_ALFABET = 26;
	private static final char BASE = LETTERS_IN_EN_ALFABET;
	private static final char A_LETTER = 65;
	private static final Logger LOGGER = Logger.getLogger(ExcelUtilityHelper.class.getName());

	public static Map<String, Integer> mapColumnNamesToIndex(Row row) {

		Map<String, Integer> columnNameMap = new HashMap<>();

		short minColumnIndex = row.getFirstCellNum();

		short maxColumnIndex = row.getLastCellNum();

		for (short colIndex = minColumnIndex; colIndex < maxColumnIndex; colIndex++) {

			Cell currentCell = row.getCell(colIndex);

			columnNameMap.put(currentCell.getStringCellValue(), currentCell.getColumnIndex());

		}

		return columnNameMap;
	}

	public static void filterForDriversWithOccurencesOfThreeOrMore(Sheet sheet, Map<String, Integer> driverMap,
			Map<String, Integer> columnNameMap) {

		for (int rowIndex = 1; rowIndex < sheet.getLastRowNum(); rowIndex++) {

			if (sheet.getRow(rowIndex) == null) {

				continue;

			} else {

				Row currentRow = sheet.getRow(rowIndex);

				int indexOfDriverNameColumn = columnNameMap.get("Driver Name");

				Cell driverNameCell = currentRow.getCell(indexOfDriverNameColumn);

				String driverName = driverNameCell.getStringCellValue();

				if (driverMap.get(driverName) < 3) {

					sheet.removeRow(currentRow);

				}
			}
		}
	}
	
	/* private static void removeRows(Sheet sheet) {

		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
			sheet.removeRow(sheet.getRow(i));
		}
	}*/

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

	public static void deleteColumn(Sheet sheet, int columnToDelete) {

		for (int indexOfRow = 0; indexOfRow < sheet.getPhysicalNumberOfRows(); indexOfRow++) {

			Row currentRow = sheet.getRow(indexOfRow);

			for (int indexOfColumn = columnToDelete; indexOfColumn < currentRow
					.getPhysicalNumberOfCells(); indexOfColumn++) {

				Cell oldCell = currentRow.getCell(indexOfColumn);

				if (oldCell != null) {

					currentRow.removeCell(oldCell);

				}

				Cell nextCell = currentRow.getCell(indexOfColumn + 1);

				if (nextCell != null) {

					Cell newCell = currentRow.createCell(indexOfColumn, nextCell.getCellType());

					cloneCell(newCell, nextCell);

					// Set the column width only on the first row.

					// Other wise the second row will overwrite the original column width set
					// previously.

					if (indexOfRow == 0) {

						sheet.setColumnWidth(indexOfColumn, sheet.getColumnWidth(indexOfColumn + 1));

					}
				}
			}
		}
	}

	private static void cloneCell(Cell newCell, Cell oldCell) {

		if (CellType.BOOLEAN == oldCell.getCellType()) {

			newCell.setCellValue(oldCell.getBooleanCellValue());

		} else if (CellType.NUMERIC == oldCell.getCellType()) {

			newCell.setCellValue(oldCell.getNumericCellValue());

		} else if (CellType.STRING == oldCell.getCellType()) {

			newCell.setCellValue(oldCell.getStringCellValue());

		} else if (CellType.ERROR == oldCell.getCellType()) {

			newCell.setCellValue(oldCell.getErrorCellValue());

		} else if (CellType.FORMULA == oldCell.getCellType()) {

			newCell.setCellValue(oldCell.getCellFormula());

		}
		
		newCell.setCellComment(oldCell.getCellComment());

		newCell.setCellStyle(oldCell.getCellStyle());
	}

	public static void insertNewColumnBefore(XSSFWorkbook workbook, int sheetIndex, int columnIndex) {

		assert workbook != null;

		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

		evaluator.clearAllCachedResultValues();

		Sheet sheet = workbook.getSheetAt(sheetIndex);

		int totalRows = sheet.getPhysicalNumberOfRows();

		int totalColumns = sheet.getRow(0).getLastCellNum();

		for (int indexOfRow = 0; indexOfRow < totalRows; indexOfRow++) {

			Row currentRow = sheet.getRow(indexOfRow);

			if (currentRow == null) {

				continue;

			}

			for (int column = totalColumns; column > columnIndex; column--) {

				Cell rightCell = currentRow.getCell(column);

				if (rightCell != null) {

					currentRow.removeCell(rightCell);

				}

				Cell leftCell = currentRow.getCell(column - 1);

				if (leftCell != null) {

					Cell newCell = currentRow.createCell(column, leftCell.getCellType());

					cloneCell(newCell, leftCell);

					if (newCell.getCellType() == CellType.FORMULA) {

						newCell.setCellFormula(updateFormula(newCell.getCellFormula(), columnIndex));

						evaluator.notifySetFormula(newCell);

						CellValue cellValue = evaluator.evaluate(newCell);

						evaluator.evaluateFormulaCell(newCell);

					}
				}
			}

			CellType cellType = CellType.BLANK;

			Cell currentEmptyWeekCell = currentRow.getCell(columnIndex);

			if (currentEmptyWeekCell != null) {

				currentRow.removeCell(currentEmptyWeekCell);
			}

			currentRow.createCell(columnIndex, cellType);
		}

		XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
	}

	public static String updateFormula(String cellFormula, int columnIndex) {

		String existingColName = getReferenceForColumnIndex(columnIndex);

		String newColName = getReferenceForColumnIndex(columnIndex + 1);

		String newCellFormula = cellFormula.replace(existingColName, newColName);

		return newCellFormula;

	}

	private static String getReferenceForColumnIndex(int columnIndex) {

		StringBuilder sb = new StringBuilder();

		while (columnIndex >= 0) {

			if (columnIndex == 0) {

				sb.append((char) A_LETTER);

				break;
			}

			char code = (char) (columnIndex % BASE);

			char letter = (char) (code + A_LETTER);

			sb.append(letter);

			columnIndex /= BASE;

			columnIndex -= 1;
		}

		return sb.reverse().toString();

	}
	
	public static void sortSheet(Workbook workbook, Sheet sheet, Map<String, Integer> columnNameMap) {

		List<Row> rows = Lists.newArrayList(sheet.rowIterator());

		int indexOfAreaCoachColumn = 1;

		int indexOfStoreColumn = columnNameMap.get("Store");

		rows.sort((row1, row2) -> {

			if (row1.getCell(indexOfAreaCoachColumn).getStringCellValue()
					.equals(row2.getCell(indexOfAreaCoachColumn).getStringCellValue())) {

				return row1.getCell(indexOfStoreColumn).getStringCellValue()
						.compareTo(row2.getCell(indexOfStoreColumn).getStringCellValue());

			} else {

				return row1.getCell(indexOfAreaCoachColumn).getStringCellValue()
						.compareTo(row2.getCell(indexOfAreaCoachColumn).getStringCellValue());

			}
		});
		
		/* rows.forEach(row -> LOGGER.warn("AFTER SORTING ---> " + row.getCell(0).getStringCellValue()));
		
		   rows.forEach(row -> LOGGER.warn("AFTER SORTING ---> " + row.getCell(1).getStringCellValue()));

			removeRows(sheet);
			
			rows.forEach(row -> LOGGER.warn("AFTER removeRows() method ---> " + row.getCell(0).getStringCellValue()));
			
			rows.forEach(row -> LOGGER.warn("AFTER removeRows() method ---> " + row.getCell(1).getStringCellValue()));
		
		*/

		for (int i = 1; i < rows.size(); i++) {

			// rows.forEach(row -> row.forEach(cell -> LOGGER.info("IN LOOP OF CLONING CELLS ----> " + cell.toString())));

			Row rowToOverwrite = sheet.getRow(i);

			Row sortedRow = rows.get(i);
			
			// sortedRow.forEach(cell -> LOGGER.warn("SORTED ROW IN LOOP " + cell.toString()));

			for (int j = 0; j < sortedRow.getPhysicalNumberOfCells(); j++) {

				Cell cellToOverwrite = rowToOverwrite.getCell(j);

				// LOGGER.info("Old cell is ->>>>> " + oldCell.toString());

				Cell sortedCell = sortedRow.getCell(j);

				cloneCell(cellToOverwrite, sortedCell);

				// LOGGER.info("New cell after cloning is ->>>>>> " + newCell.toString());
			}
		}
	}

}
