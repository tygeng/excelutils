/**
 * 
 */
package tygeng.excelutils;

import static org.junit.Assert.*;

import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import tygeng.excelutils.Merger.IllegalSpreadSheetException;

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
			Merger m = new Merger(Utils.getWorkbook(new File("/home/tony1/tmp/excelutils-test/original.xlsx")), log);
			m.merge(new File("/home/tony1/tmp/excelutils-test/test1.xlsx"));
			m.merge(new File("/home/tony1/tmp/excelutils-test/test2.xlsx"));
			m.merge(new File("/home/tony1/tmp/excelutils-test/test3.xls"));
			m.merge(new File("/home/tony1/tmp/excelutils-test/test4.xlsx"));
			m.write(new File("/home/tony1/tmp/excelutils-test/merged.xlsx"));
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
