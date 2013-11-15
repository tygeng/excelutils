package tygeng.excelutils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * User: tianyu
 * Date: 11/10/13
 * Time: 12:25 PM
 * To change this template use File | Settings | File Templates.
 */
public class TrackerEntryPoint {
    public static void main(String[] args) {
        File loggerFile = null;
        File targetFile = null;
        List<File> inputFiles = new ArrayList<File>();
        File outputFile = null;
        String updateColumn = null;
        boolean countReason = false;
        try {
            for (int i = 0; i < args.length; i++) {
                if ("-t".equals(args[i])) {
                    i++;
                    if (Utils.isExcelFile(args[i])) {
                        targetFile = new File(args[i]);
                    }
                } else if ("-d".equals(args[i])) {
                    i++;
                    File dir = new File(args[i]);
                    if (dir.isDirectory() && dir.canRead()) {
                        File[] files = dir.listFiles();
                        for (File file : files) {
                            if (Utils.isExcelFile(file)) {
                                inputFiles.add(file);
                            }
                        }
                    } else {
                        System.err.println("Please specify a readable directory after '-d'.");
                    }
                } else if ("-o".equals(args[i])) {
                    i++;

                    outputFile = new File(Utils.setExtension(args[i], "xlsx"));

                } else if ("-l".equals(args[i])) {
                    i++;
                    loggerFile = new File(Utils.setExtension(args[i], "txt"));
                } else if ("-c".equals(args[i])) {
                    i++;
                    updateColumn = args[i].toUpperCase();
                } else if ("-r".equals(args[i])) {
                    countReason = true;
                } else {
                    if (Utils.isExcelFile(args[i])) {
                        inputFiles.add(new File(args[i]));
                    }

                }
            }


        } catch (IndexOutOfBoundsException e) {
            usage();
            return;
        }
        boolean needUsage = false;
        if (targetFile == null) {
            System.err.print("Please specify a valid target file after '-t'.");
            needUsage = true;
        }
        if (updateColumn == null) {
            System.err.println("Please specify a column to update.");
            needUsage = true;
        }
        if (outputFile == null) {
            if (targetFile != null) {
                outputFile = new File(Utils.getOutputName(targetFile.getName(), "tracked"));
            }
        } else if (outputFile.equals(targetFile)) {
            System.err.println("Please specify a different output file than the target file.");
            needUsage = true;
        }
        if (inputFiles.isEmpty()) {
            System.err.println("Please specify some input files.");
            needUsage = true;
        }
        if (needUsage) {
            usage();
            return;
        }
        Logger log = null;
        if (loggerFile == null) {
            log = new Logger();
        } else {
            try {
                log = new Logger(loggerFile);
            } catch (IOException e) {
                System.err.println("Cannot write to log file " + loggerFile.getName());
                return;
            }
        }
        // Done collecting input files.
        Tracker tracker = new Tracker(log);
        for (File file : inputFiles) {
            System.err.println("Reading input file " + file.getName() + "...");

            tracker.accumulate(file, countReason);
            log.flush();
        }
        try {
            System.err.println("Updating column " + updateColumn + " of " + targetFile + ".");
            tracker.update(targetFile, outputFile, updateColumn);
            log.flush();
        } catch (IOException e) {

            System.err.println("Cannot read target file.");
        } catch (InvalidFormatException e) {
            System.err.println("Target file is corrupted.");
        }
        log.close();

    }

    private static void usage() {

        System.err.println("Usage: tracker -t <target> -c <column> [-l <log>] [-d <input dir>] [-o <output>] [<files> ...] [-r]\n");
        System.err.println("<target>     the report file");
        System.err.println("<column>     the column to be upated, for example 'A' or 'BU'. Note, case doesn't matter. ");
        System.err.println("<log>        used to collect errors in a separate file");
        System.err.println("<input dir>  the directory containing all input files that were sent to you");
        System.err.println("<output>     the output file name, default to YYYY-MM-DD-hhmmss-<target>-tracked.xlsx");
        System.err.println("<files>      other input files not in the input directory");
        System.err.println("-r           count valid Late Reason Code");

    }
}
