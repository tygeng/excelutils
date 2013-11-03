/**
 * 
 */
package tygeng.excelutils;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import java.io.File;
import java.util.HashMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import java.util.Map;
import org.apache.poi.ss.usermodel.Workbook;
import static tygeng.common.utils.string.StringUtils.normalize4Hash;

/**
 * 
 * @author Tianyu Geng
 * @version Oct 30, 2013
 */
/**
 * 
 * @author Tianyu Geng
 * @version Oct 31, 2013
 */
/**
 * 
 * @author Tianyu Geng
 * @version Oct 31, 2013
 */
public class Merger {
	private static class ColumnProperty {
		int index;
		boolean isDate;

		ColumnProperty(int index, boolean isDate) {
			this.index = index;
			this.isDate = isDate;
		}
	}
	private static CellStyle dateStyle;
	private static CreationHelper createHelper;

	private Map<String, Map<String, ColumnProperty>> headerMaps;
	private Map<String, Integer> sheetIndex;
	private Workbook target;
	private Logger log;
//	private Workbook config;

	public Merger(Workbook target, Logger log, Workbook config)
			throws IllegalSpreadSheetException {
		createHelper = target.getCreationHelper();
		dateStyle = target.createCellStyle();
		dateStyle.setDataFormat(createHelper.createDataFormat().getFormat(
				"m/d/yyyy"));
		this.log = log;
//		this.config = config;
		sheetIndex = new HashMap<String, Integer>();
		int numSheets = target.getNumberOfSheets();
		headerMaps = new HashMap<String, Map<String, ColumnProperty>>();
		Sheet configSheet = null;
		if (config != null) {
			configSheet = config.getSheetAt(0);
		}
		for (int i = 0; i < numSheets; i++) {
			Sheet currentSheet = target.getSheetAt(i);
			if (currentSheet.getSheetName().contains("Guidance")
					|| currentSheet.getSheetName().contains("Sample Order")) {
				continue;
			}
			sheetIndex.put(normalize4Hash(currentSheet.getSheetName()), i);

			headerMaps.put(normalize4Hash(currentSheet.getSheetName()),
					getHeaderMap(currentSheet, configSheet, i));
		}
		this.target = target;

	}

	/**
	 * Merge a spread sheet file to the target file.
	 * 
	 * @param wbFile
	 */
	public void merge(File wbFile) {
		Workbook wb;
		try {
			wb = Utils.getWorkbook(wbFile);
			int numSheets = wb.getNumberOfSheets();
			for (int i = 0; i < numSheets; i++) {
				Sheet currentSheet = wb.getSheetAt(i);
				if (currentSheet.getSheetName().contains("Guidance")
						|| currentSheet.getSheetName().contains("Sample Order")) {
					continue;
				}
				int endRow = Utils.getDataEndRow(currentSheet);
				int startRow = Utils.getDataStartRow(currentSheet);
				if (startRow >= endRow) {
					continue;
				}
				log.s("   >>> " + currentSheet.getSheetName());
				if (Utils.getNonemptyRowSince(currentSheet, endRow) > 0) {
					log.s("#### Sheet \""
							+ currentSheet.getSheetName()
							+ "\" may contain extra rows that is ill formated. Skipped.");
					continue;
				}
				// System.out.println("Sheet " + i + ":"
				// + currentSheet.getSheetName());
				Map<String, ColumnProperty> targetSheetHeader = headerMaps
						.get(normalize4Hash(currentSheet.getSheetName()));
				if (targetSheetHeader == null) {
					log.s("#### Ignore sheet \"" + currentSheet.getSheetName()
//							+ "\" in file \"" + wbFile.getName()
							+ "\" because it doesn't exist in the target.");
				} else {
					Row r2 = currentSheet.getRow(1);
					Row r3 = currentSheet.getRow(2);
					Row r4 = currentSheet.getRow(3);
					if (r2 == null || r3 == null || r4 == null) {
						log.s("#### Ignore sheet \"" + currentSheet.getSheetName()
//								+ "\" in file \"" + wbFile.getName()
								+ "\" because its header is less than 3 rows.");
					}
					String r2State = null;
					int headerSize = r4.getLastCellNum();
					ColumnProperty[] indexCorrespondence = new ColumnProperty[headerSize];
					for (int j = 0; j < indexCorrespondence.length; j++) {
						indexCorrespondence[j] = null;
					}
					// Match the header index by header content
					for (int j = 0; j < headerSize; j++) {
						Cell r2Cell = r2.getCell(j);
						Cell r3Cell = r3.getCell(j);
						Cell r4Cell = r4.getCell(j);
						String r4State = null;
						if (r2Cell != null) {
							String r2Contents = r2Cell.getStringCellValue();
							if (!r2Contents.isEmpty()) {
								r2State = r2Contents;
							}
						}
						if (r4Cell != null) {
							r4State = r4Cell.getStringCellValue();
						} else {
							r4State = "";
						}
						ColumnProperty correspondingTargetColumnProperty;

						if (r3Cell == null || r3Cell.toString().isEmpty()) {
							correspondingTargetColumnProperty = targetSheetHeader
									.get(normalize4Hash(r4State));

						} else {
							correspondingTargetColumnProperty = targetSheetHeader
									.get(normalize4Hash(r2State + r4State));
						}
						if (correspondingTargetColumnProperty != null) {
							indexCorrespondence[j] = correspondingTargetColumnProperty;
						} else {
							log.m("Header \"" + r2State + " "
									+ r3Cell.toString() + " " + r4State
									+ "\" is not in target.");
						}
					}
					log.flush();
					// Copy cell contents
					Integer sheetIndex = this.sheetIndex
							.get(normalize4Hash(currentSheet.getSheetName()));
					Sheet targetSheet = target.getSheetAt(sheetIndex);
					int nextRowIndex = Utils.getDataEndRow(targetSheet);

					// start with row 5 where actual contents are
					for (int j = startRow; j < endRow; j++) {
						Row targetRow = targetSheet.createRow(nextRowIndex++);
						Row currentRow = currentSheet.getRow(j);
						for (int k = 0; k < headerSize; k++) {
							if (indexCorrespondence[k] == null) {
								continue;
							}
							Cell currentCell = currentRow.getCell(k);
							if (currentCell != null) {
								Cell targetCell = targetRow
										.createCell(indexCorrespondence[k].index);
								boolean isDate = indexCorrespondence[k].isDate;
								Utils.copyCell(targetCell, currentCell, isDate, dateStyle);
							}
						}
					}

				}
			}
		} catch (InvalidFormatException e) {
			log.s("Spread sheet " + wbFile.getName() + " is corrupted.");
			log.flush();
		} catch (IOException e) {
			log.s("Cannot read spread sheet " + wbFile.getName() + ".");
			log.flush();
		}
	}

 Map<String, ColumnProperty> getHeaderMap(Sheet sheet, Sheet config,
			int sheetIndex) throws IllegalSpreadSheetException {
		Map<String, ColumnProperty> result = new HashMap<String, ColumnProperty>();
		Row r2 = sheet.getRow(1);
		Row r3 = sheet.getRow(2);
		Row r4 = sheet.getRow(3);
		if (r2 == null || r3 == null || r4 == null) {
			throw new IllegalSpreadSheetException(
					"Illegal target spread sheet header. Row 2 or row4 is empty.");
		}

		String r2State = null;
		int headerSize = r4.getLastCellNum();
		for (int i = 0; i < headerSize; i++) {
			Cell r2Cell = r2.getCell(i);
			Cell r3Cell = r3.getCell(i);
			Cell r4Cell = r4.getCell(i);
			String r4State = null;
			if (r2Cell != null) {
				String r2Contents = Utils.getStringRepresentation(r2Cell);
				if (!r2Contents.isEmpty()) {
					r2State = r2Contents;
				}
			}
			boolean isDate = false;
			if (r4Cell != null) {
				r4State = Utils.getStringRepresentation(r4Cell);
				// } else {
				// r4State = "";
				if (r4State != null && r4State.contains("Date")) {
					isDate = true;
				}

			} else {
				break;
			}

			if (r3Cell == null || r3Cell.toString().isEmpty()) {
				result.put(normalize4Hash(r4State), new ColumnProperty(i,
						isDate));
			} else {
				result.put(normalize4Hash(r2State + r4State),
						new ColumnProperty(i, isDate));
			}

		}
		if (config != null) {
			int size = config.getLastRowNum() + 1;
			for (int j = 1; j < size; j++) {
				Row r = config.getRow(j);
				if (r != null) {
					Cell c0 = r.getCell(0);

					if (c0 != null) {
						String c0Content = c0.getStringCellValue();
						ColumnProperty cp = result
								.get(normalize4Hash(c0Content));
						if (cp != null) {

							for (int k = 1; k < r.getLastCellNum(); k++) {
								Cell c = r.getCell(k);
								if (c != null) {
									result.put(normalize4Hash(c
											.getStringCellValue()), cp);
								}
							}
						} else {
							log.s("Header " + c0Content + " is not in target.");
						}
					}

				}

			}
		}
		return result;
	}

}
