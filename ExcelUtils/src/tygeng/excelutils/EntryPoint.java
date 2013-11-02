/**
 * 
 */
package tygeng.excelutils;

import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.io.File;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 
 * @author Tianyu Geng
 * @version Oct 31, 2013
 */
public class EntryPoint {
	public static final String USAGE = "Usage: ExcelUtils <action> -t <target> [-d <directory>] [-l <log>] [-o <output>] [<files> ...] [-c <config>]\n"
			+ "<action> \tthe action to perform. Either 'merge' (m) or 'normalize' (n) or 'analyze' (a)\n"
			+ "<target> \tthe output template file\n"
			+ "<directory> \tthe directory(ies) containing input excel files\n"
			+ "<log> \tthe log file\n"
			+ "<output> \tthe output file. Default to <target>-<action>d\n"
			+ "<file> \tother input files\n"
			+ "<config> \ta config file for merge or normalize";

	public static enum Action {

		NONE, MERGE, NORMALIZE, ANALYZE
	}

	public static void main(String[] args) {
		Logger log = null;
		try {

			ArrayList<File> inputFiles = new ArrayList<File>();
			File targetFile = null;
			File dirFile = null;
			File outputFile = null;
			File logFile = null;
			File configFile = null;

			Action action = Action.NONE;
			try {
				if ("merge".equals(args[0]) || "m".equals(args[0])) {
					action = Action.MERGE;
				} else if ("normalize".equals(args[0]) || "n".equals(args[0])) {
					action = Action.NORMALIZE;
				} else if ("analyze".equals(args[0]) || "a".equals(args[0])) {
					action = Action.ANALYZE;
				} else {
					System.err.println("You need to specify <action>.");
				}

				for (int i = 1; i < args.length; i++) {

					if ("-t".equals(args[i])) {
						i++;
						targetFile = new File(args[i]);
					} else if ("-d".equals(args[i])) {
						i++;
						dirFile = new File(args[i]);
						if (dirFile.isDirectory() && dirFile.canRead()) {
							File[] dirFiles = dirFile.listFiles();
							for (File f : dirFiles) {
								if (f.canRead()
										&& (f.getName().endsWith("xls") || f
												.getName().endsWith("xlsx"))) {
									inputFiles.add(f);
								}
							}
						} else {
							System.err.println(dirFile.getName()
									+ " is not a readable directory.");
						}
					} else if ("-l".equals(args[i])) {
						i++;
						String name = args[i];
						if (!name.endsWith(".txt")) {
							name = name + ".txt";
						}
						logFile = new File(name);
					} else if ("-o".equals(args[i])) {
						i++;
						String name = args[i];
						if (!name.endsWith(".xlsx") && !name.endsWith(".xls")) {
							name = name + ".xlsx";
						}
						outputFile = new File(name);

					} else if ("-c".equals(args[i])) {
						i++;
						configFile = new File(args[i]);
					} else {
						File f = new File(args[i]);
						if (f.canRead()
								&& (f.getName().endsWith("xls") || f.getName()
										.endsWith("xlsx"))) {
							inputFiles.add(f);
						} else {
							System.err.println(f.getName() + " is ignored.");
						}
					}

				}
				if (targetFile == null) {
					System.err.println("You need to specify <target>.");
					throw new Exception();
				}
				if (outputFile == null) {
					String name = targetFile.getName();
					int dotPos = name.lastIndexOf('.');
					String baseName = name.substring(0, dotPos);
					String extension = name.substring(dotPos);
					outputFile = new File(baseName + "-"
							+ action.toString().toLowerCase() + "d" + extension);
				}
			} catch (Exception e) {
				System.err.println(USAGE);
				return;
			}

			try {
				if (logFile != null) {
					log = new Logger(logFile);
				} else {
					System.err.print("No log file specified. ");
					System.err.println("Output to standard out.");
					log = new Logger();
				}

			} catch (IOException e) {
				System.err.print("Log file is not writable. ");
				System.err.println("Output to standard out.");
				log = new Logger();
			}
			Workbook target = null;
			Workbook config = null;
			try {
				target = WorkbookFactory.create(targetFile);
				if (configFile != null) {
					System.err.println("Using config file "
							+ configFile.getName());
					config = WorkbookFactory.create(configFile);
				}
			} catch (InvalidFormatException e) {
				System.err
						.println("Target or config file is corrupted. Please check you have specified a valid target file.");
				return;
			} catch (IOException e) {
				System.err
						.println("Target or config file is not writable. Please check you have specified a valid target file.");
				return;
			}
			switch (action) {
			case MERGE:
				try {
					Merger m = new Merger(target, log, config);
					int size = inputFiles.size();
					for (int i = 0; i < size; i++) {
						File f = inputFiles.get(i);
						log.s("[" + (i + 1) + " / " + size + "] " + f.getName());
						if (logFile != null) {
							System.err.println("[" + (i + 1) + " / " + size
									+ "] " + f.getName());

						}
						m.merge(f);
						log.s("======== Done!");
						log.flush();
					}
					Utils.write(target, outputFile);
				} catch (IllegalSpreadSheetException e) {
					// e.printStackTrace();
					System.err
							.println("The target file is illegal. Pleaes check the header format.");
					return;
				} catch (IOException e) {
					// e.printStackTrace();
					System.err.println("Cannot write output file to "
							+ outputFile.getName() + ".");
				}
				break;
			case NORMALIZE:
				try {
					Normalizer n = new Normalizer(config, log);
					n.normalize(target);

					Utils.write(target, outputFile);

				} catch (IOException e) {
					System.err.println("Cannot write output file to "
							+ outputFile.getName() + ".");
				}
				break;
			case ANALYZE:
				
				break;
			case NONE:
			default:
				break;
			}

		} catch (Exception e) {
			e.printStackTrace();

		} finally {

			if (log != null) {
				log.close();
			}
		}
	}
}
