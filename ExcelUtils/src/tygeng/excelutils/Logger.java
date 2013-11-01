/**
 * 
 */
package tygeng.excelutils;

import java.io.IOException;

import java.io.File;
import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.OutputStreamWriter;
import java.io.OutputStream;
import java.io.PrintWriter;

/**
 * 
 * @author Tianyu Geng
 * @version Oct 31, 2013
 */
public class Logger {
	private PrintWriter writer;

	public Logger() {
		writer = new PrintWriter(new OutputStreamWriter(System.err));

	}

	public Logger(File logFile) throws IOException {
		if (logFile == null) {
			writer = new PrintWriter(new OutputStreamWriter(System.err));
		} else {

			writer = new PrintWriter(
					new BufferedWriter(new FileWriter(logFile)));
		}
	}

	public void m(String message) {
		writer.println("    " + message);
	}

	public void s(String sheetMessage) {
		writer.println(sheetMessage);
	}
	public void close() {
		writer.close();
	}

	public void flush() {
		writer.flush();
	}
}
