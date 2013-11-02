/**
 * 
 */
package tygeng.excelutils;

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

	/**
	 * Return the row index after the last data row.
	 * 
	 * @param sheet
	 * @return
	 */
	public static int getDataEndRow(Sheet sheet) {
		int size = sheet.getLastRowNum()+1;
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
		int size = sheet.getLastRowNum()+1;
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
	public static void write(Workbook target,File targetFile) throws IOException {
		BufferedOutputStream out = new BufferedOutputStream(
				new FileOutputStream(targetFile));
		target.write(out);
		out.close();
	}
}
