package tygeng.excelutils;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import tygeng.common.utils.string.StringUtils;

import java.io.*;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;


/**
 * Created with IntelliJ IDEA.
 * User: tianyu
 * Date: 11/9/13
 * Time: 8:10 PM
 * To change this template use File | Settings | File Templates.
 */
public class Tracker {
    private Map<String, CarrierStats> counters;
    private Map<String, CarrierStats> mpcCounters;
    private Logger log;

    public Tracker(Logger log) {
        counters = new HashMap<String, CarrierStats>();
        mpcCounters = new HashMap<String, CarrierStats>();
        this.log = log;
    }

    private void counterPlus(String key) {
        String keyNormed = StringUtils.normalize4Hash(key);
        CarrierStats count = counters.get(StringUtils.normalize4Hash(keyNormed));
        if (count == null) {
            count = new CarrierStats(key);
            counters.put(keyNormed, count);
        }
        count.plus();
    }

    public void accumulate(File workbookFile, boolean countReason) {

        Workbook wb;
        try {
            wb = WorkbookFactory.create(workbookFile);
        } catch (IOException e) {
            log.m("Cannot read input file " + workbookFile.getName() + ".");
            return;
        } catch (InvalidFormatException e) {
            log.m("Input file " + workbookFile + " is corrupted.");
            return;
        }

        int numSheets = wb.getNumberOfSheets();
        log.s(">>> " + workbookFile.getName());
        for (int i = 0; i < numSheets; i++) {
            Sheet sheet = wb.getSheetAt(i);
            Row header = sheet.getRow(0);
            if (header == null) {
                log.m(">>> Sheet '" + sheet.getSheetName() + "' doesn't contain a valid header. Skipped");
                continue;
            }
            int carrerNameIndex = Utils.getFirstOccuranceInRow(header, "Carrier Name");
            int reasonIndex = Utils.getFirstOccuranceInRow(header, "Late Reason Code");

            if (carrerNameIndex == -1) {
                log.m(">>> Sheet '" + sheet.getSheetName() + "' doesn't contain header cell 'Carrier Name'. Skipped");
                continue;
            }
            log.m(">>> Processing sheet '" + sheet.getSheetName() + "...");
            int numRows = sheet.getLastRowNum() + 1;
            for (int k = 1; k < numRows; k++) {
                Row row = sheet.getRow(k);
                if (row != null) {
                    Cell cell = row.getCell(carrerNameIndex);
                    if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING) {
                        if (countReason) {
                            Cell reason = row.getCell(reasonIndex);
                            if (reason != null && reason.getCellType() == Cell.CELL_TYPE_STRING) {
                                String reasonValue = reason.getStringCellValue();
                                if (reasonValue != null) {
                                    String reasonValueNormed = StringUtils.normalize4Hash(reasonValue);
                                    if (reasonValueNormed.matches("\\w\\w")) {
                                        String name = StringUtils.normalize4Hash(cell.getStringCellValue());
                                        counterPlus(name);
                                    }else{
                                        log.m("!!! Row "+(k+1)+" contains illegal reason code.");
                                    }
                                }

                            }
                        } else {
                            String name = StringUtils.normalize4Hash(cell.getStringCellValue());
                            counterPlus(name);
                        }

                    }
                }
            }
        }
    }

    public void update(File input, File output, String columnIndex) throws IOException, InvalidFormatException {
        Workbook wb = WorkbookFactory.create(input);
        int updateIndex = CellReference.convertColStringToIndex(columnIndex);

        Sheet firstSheet = wb.getSheetAt(0);
        updateSheet(firstSheet, updateIndex, true);

        Sheet secondSheet = wb.getSheetAt(1);
        updateSheet(secondSheet, updateIndex, false);

        OutputStream out = new BufferedOutputStream(new FileOutputStream(output));
        wb.write(out);
        out.close();

    }

    public void updateSheet(Sheet sheet, int updateIndex, boolean firstSheet) {
        int ROW_STARTING_INDEX = 5;
        Row header = sheet.getRow(ROW_STARTING_INDEX - 1);
        int carrierNameIndex = Utils.getFirstOccuranceInRow(header, "Carrier Name", "Freight Forward");
        Map<String, Integer> carrierIndices = new HashMap<String, Integer>();
        int numRow = sheet.getLastRowNum() + 1;
        for (int i = ROW_STARTING_INDEX; i < numRow; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell carrierNameCell = row.getCell(carrierNameIndex);
                if (carrierNameCell != null && carrierNameCell.getCellType() == Cell.CELL_TYPE_STRING) {
                    String carrierName = StringUtils.normalize4Hash(carrierNameCell.getStringCellValue());
                    carrierIndices.put(carrierName, i);
                }
            }
        }
        Set<Map.Entry<String, CarrierStats>> carrierCounts;
        if (firstSheet) {

            carrierCounts = this.counters.entrySet();
        } else {
            carrierCounts = this.mpcCounters.entrySet();
        }

        for (Map.Entry<String, CarrierStats> entry : carrierCounts) {
            CarrierStats count = entry.getValue();
            if (count == null || count.getValue() == 0) {
                continue;
            }
            Integer rowIndex = carrierIndices.get(entry.getKey());
            if (rowIndex != null) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.createCell(updateIndex);
                cell.setCellValue(count.getValue());
            } else if (firstSheet) {
                this.mpcCounters.put(entry.getKey(), entry.getValue());
            } else {
                // log the carrier that doesn't exist in the accumualted table.
                log.m("Carrier " + entry.getValue().getOriginalName() + " not in the spreadsheet to be updated.");

            }
        }
        log.s("Done updating sheet " + sheet.getSheetName());

    }

    private static class CarrierStats {
        private int value;
        private String originalName;

        public CarrierStats(String originalName) {
            this.originalName = originalName;
            value = 0;
        }

        public void plus() {
            value++;
        }

        public int getValue() {
            return value;
        }

        public String getOriginalName() {
            return originalName;
        }
    }


}
