/**
 * 
 */
package tygeng.excelutils;

import org.apache.poi.hssf.util.CellReference;
import org.junit.Test;

/**
 * 
 * @author Tianyu Geng
 * @version Nov 2, 2013
 */
public class UtilsTest {

	@Test
	public void test() {
		System.out.println(Utils.getOutputName("target.xlsx", "merge"));
		System.out.println(Utils.getOutputName("2013-03-20-target.xlsx", "merge"));
		System.out.println(Utils.getOutputName("2013-03-20-target-merge.xlsx", "merge"));
        System.out.println(CellReference.convertColStringToIndex("'BU'"));
	}

}
