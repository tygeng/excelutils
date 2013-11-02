/**
 * 
 */
package tygeng.excelutils;

import static org.junit.Assert.*;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import java.io.File;
import org.junit.Test;

/**
 * 
 * @author Tianyu Geng
 * @version Oct 31, 2013
 */
public class MergerTest {

	@Test
	public void testMerge() {
		try {
			Logger log = new Logger();
			Workbook target = Utils.getWorkbook(new File("/home/tony1/tmp/excelutils-test/original.xlsx"));
			Merger m = new Merger(target,  log, null);
			m.merge(new File("/home/tony1/tmp/excelutils-test/test1.xlsx"));
			m.merge(new File("/home/tony1/tmp/excelutils-test/test2.xlsx"));
			m.merge(new File("/home/tony1/tmp/excelutils-test/test3.xls"));
			m.merge(new File("/home/tony1/tmp/excelutils-test/test4.xlsx"));
			Utils.write(target, new File("/home/tony1/tmp/excelutils-test/merged.xlsx"));
			log.close();

		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalSpreadSheetException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
