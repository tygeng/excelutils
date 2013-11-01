/**
 * 
 */
package tygeng.excelutils;

import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.usermodel.Sheet;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import java.io.File;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Workbook;


/**
 * 
 * @author Tianyu Geng
 * @version Oct 30, 2013
 */
public class Tester {

	public static void main(String[] args) {
		try {
			Workbook wb = WorkbookFactory.create(new File("/home/tony1/tmp/excelutils-test/test1.xlsx"));
			Sheet s = wb.getSheet("零担&整车运输 LTL&FTL");
			Row r4=s.getRow(4);
			System.out.println(r4.getCell(2));
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
}
