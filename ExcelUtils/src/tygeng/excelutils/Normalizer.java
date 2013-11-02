/**
 * 
 */
package tygeng.excelutils;

import java.io.IOException;

import java.io.FileOutputStream;
import java.io.File;
import java.io.BufferedOutputStream;
import javax.xml.bind.annotation.adapters.NormalizedStringAdapter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Workbook;
import static tygeng.common.utils.string.StringUtils.normalize4Hash;

;

/**
 * 
 * @author Tianyu Geng
 * @version Nov 1, 2013
 */
public class Normalizer {

	private Workbook configs;
	private Map<String, Map<String, Cell>> dupMaps;
	private Logger log;

	public Normalizer(Workbook configs, Logger log) {
		this.configs = configs;
		this.log = log;
		dupMaps = new HashMap<String, Map<String, Cell>>();
		int numSheet = configs.getNumberOfSheets();
		for (int i = 0; i < numSheet; i++) {
			Sheet config = configs.getSheetAt(i);
			dupMaps.put(normalize4Hash(config.getSheetName()),
					getDupMap(config));
		}

	}

	public void normalize(Workbook target) {
		int numSheets = target.getNumberOfSheets();
		for (int i = 0; i < numSheets; i++) {
			Sheet sheet = target.getSheetAt(i);
			log.s(">>> Sheet " + sheet.getSheetName());
			int start = Utils.getDataStartRow(sheet);
			int end = Utils.getDataEndRow(sheet);
			Row r4 = sheet.getRow(3);
			if (r4 == null) {
				continue;
			}
			Row[] rows = new Row[end];
			for (int j = start; j < end; j++) {
				rows[j] = sheet.getRow(j);
			}
			int lastColumn = r4.getLastCellNum();
			if (r4 != null) {
				for (int j = 0; j < lastColumn; j++) {

					String header = Utils
							.getStringRepresentation(r4.getCell(j));
					Map<String, Cell> columnMap = dupMaps
							.get(normalize4Hash(header));
					if (columnMap == null) {
						continue;
					}
					log.m("Column " + header);
					log.flush();
					for (int k = start; k < end; k++) {
						if (rows[k] != null) {
							Cell cj = rows[k].getCell(j);
							if (cj != null) {
								try {
									String content = Utils
											.getStringRepresentation(cj);
									Cell replaced = columnMap
											.get(normalize4Hash(content));
									if (replaced != null) {
										Utils.copyCell(cj, replaced, false,
												null);
									}
								} catch (Exception e) {
									log.m("Cell (" + (k + 1) + "," + (j + 1)
											+ ") is not a string.");
								}
							}

						}
					}
				}
			}
		}
	}

	private Map<String, Cell> getDupMap(Sheet config) {

		Map<String, Cell> result = new HashMap<String, Cell>();
		int size = config.getLastRowNum() + 1;
		for (int j = 1; j < size; j++) {
			Row r = config.getRow(j);
			if (r != null) {
				Cell c0 = r.getCell(0);

				if (c0 != null) {
					try {
						String c0Content = Utils.getStringRepresentation(c0);
						if (!c0Content.isEmpty()) {

							for (int k = 1; k < r.getLastCellNum(); k++) {
								Cell c = r.getCell(k);
								if (c != null) {
									try {
										String ckContent = Utils
												.getStringRepresentation(c);
										if (!ckContent.isEmpty()) {
											result.put(
													normalize4Hash(ckContent),
													c0);
										}
									} catch (Exception e) {
										log.m("Ignore non string value at ("
												+ (j + 1) + "," + (k + 1)
												+ ") in sheet "
												+ config.getSheetName() + ".");
									}
								}
							}
						} else {
							log.s("Header " + c0Content + " is not in target.");
						}
					} catch (Exception e) {
						log.m("Ignore row " + (j + 1) + " in sheet "
								+ config.getSheetName()
								+ " because of non string value.");

					}
				}

			}

		}
		return result;

	}

}
