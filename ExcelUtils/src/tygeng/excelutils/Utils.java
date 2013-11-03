/**
 * 
 */
package tygeng.excelutils;

import java.util.Date;

import java.text.SimpleDateFormat;
import org.apache.poi.ss.usermodel.CellStyle;
import java.io.FileOutputStream;
import java.io.BufferedOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.io.File;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 
 * @author Tianyu Geng
 * @version Oct 31, 2013
 */
public class Utils {
	public static Workbook getWorkbook(File wbFile)
			throws InvalidFormatException, IOException {
		return WorkbookFactory.create(wbFile);

	}

	/**
	 * Get the index of the first row with data.
	 * 
	 * @param sheet
	 * @return
	 */
	public static int getDataStartRow(Sheet sheet) {
		return 4;
	}

	public static String normalizeFileName(String raw, String extension) {
		if (raw == null) {
			return null;
		}
		if (raw.endsWith("." + extension)) {
			return raw;
		} else {
			return raw + "." + extension;
		}
	}

	/**
	 * Return the row index after the last data row.
	 * 
	 * @param sheet
	 * @return
	 */
	public static int getDataEndRow(Sheet sheet) {
		int size = sheet.getLastRowNum() + 1;
		int i;
		for (i = getDataStartRow(sheet); i < size; i++) {
			Row currentRow = sheet.getRow(i);
			if (currentRow == null) {
				break;
			}
			Cell c0 = currentRow.getCell(0);
			Cell c1 = currentRow.getCell(1);
			if (c0 == null || c1 == null || c0.toString().isEmpty()
					|| c1.toString().isEmpty()) {
				break;
			}
		}
		return i;
	}

	public static int getNonemptyRowSince(Sheet sheet, int since) {
		int size = sheet.getLastRowNum() + 1;
		int i;
		int counter = 0;
		for (i = since; i < size; i++) {
			Row r = sheet.getRow(i);
			if (r != null) {
				Cell c0 = r.getCell(0);
				Cell c1 = r.getCell(1);
				if (c0 != null && c1 != null && !c0.toString().isEmpty()
						&& !c1.toString().isEmpty()) {
					counter++;
				}

			}
		}
		return counter;
	}

	public static void write(Workbook target, File targetFile)
			throws IOException {
		BufferedOutputStream out = new BufferedOutputStream(
				new FileOutputStream(targetFile));
		target.write(out);
		out.close();
	}

	public static String getStringRepresentation(Cell cell) {

		if (cell != null) {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				return cell.getStringCellValue();

			case Cell.CELL_TYPE_NUMERIC:
				return Double.toString(cell.getNumericCellValue());
			}
		}
		return "";
	}

	public static void copyCell(Cell targetCell, Cell currentCell,
			boolean isDate, CellStyle dateStyle) {
		switch (currentCell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			targetCell.setCellValue(currentCell.getStringCellValue());
			break;
		case Cell.CELL_TYPE_NUMERIC:
			targetCell.setCellValue(currentCell.getNumericCellValue());
			if (isDate) {
				targetCell.setCellStyle(dateStyle);
			}

			break;

		case Cell.CELL_TYPE_BOOLEAN:
			targetCell.setCellValue(currentCell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_FORMULA:
			targetCell.setCellValue("="+currentCell.getCellFormula());
			break;
		}
	}

	public static String getOutputName(String target, String postfix) {
		int dotPos = target.lastIndexOf('.');
		int dateLength = 18;
		String ext, base;

		if (dotPos == -1) {
			ext = "";
			base = target;
		} else {
			base = target.substring(0, dotPos);
			ext = target.substring(dotPos);
		}
		if (!base.endsWith("-" + postfix)) {
			base += "-" + postfix;
		}
		if(base.length()>dateLength) {
			String mayBeDate = base.substring(0,dateLength);
			if(mayBeDate.matches("\\d{4}-\\d{2}-\\d{2}-\\d{6}-")) {
				base = base.substring(dateLength);
			}
		}
		String dateString = new SimpleDateFormat("yyyy-MM-dd-HHmmss-")
				.format(new Date());
		return dateString + base + ext;
	}
}
