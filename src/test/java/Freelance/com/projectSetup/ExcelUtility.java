package Freelance.com.projectSetup;

import config.VARIABLES;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.testng.annotations.DataProvider;

public class ExcelUtility {

    // ─── FIX 3: Class-level lock object for true static synchronization ───────
    private static final Object FILE_LOCK = new Object();

    @DataProvider(name = "excelData")
    public static Object[][] getExcelData() {
        File file = new File(VARIABLES.EXCEL_FILE_PATH);
        Object[][] data = null;

        // ─── FIX 4: Single try-with-resources ensures streams always close ────
        try (FileInputStream excelFile = new FileInputStream(file);
             Workbook workBook = WorkbookFactory.create(excelFile)) {

            Sheet sheet = workBook.getSheet(VARIABLES.SHEET_NAME);

            int totalRows    = sheet.getLastRowNum();
            int totalColumns = sheet.getRow(0).getLastCellNum();

            ArrayList<Object[]> dataList = new ArrayList<>();
            DataFormatter formatter = new DataFormatter();

            for (int i = 1; i <= totalRows; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Object[] rowData = new Object[totalColumns + 1]; // +1 for row index
                rowData[0] = i; // Store 1-based Excel row number

                for (int j = 0; j < totalColumns; j++) {
                    Cell cell = row.getCell(j);
                    rowData[j + 1] = formatter.formatCellValue(cell);
                }
                dataList.add(rowData);
            }

            data = new Object[dataList.size()][];
            dataList.toArray(data);

        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("Error reading excel file: " + e.getMessage());
        }

        return data;
    }

    /**
     * Update test status in Excel file.
     *
     * @param rowIndex    The Excel row number (1-based)
     * @param status      PASS / FAIL / SKIP
     */
    public static void updateTestStatus(int rowIndex, String status) {
        updateTestStatus(rowIndex, status, 55); // Default column BD (index 55)
    }

    /**
     * Update test status with a custom column index.
     *
     * @param rowIndex     The Excel row number (1-based)
     * @param status       PASS / FAIL / SKIP
     * @param statusColumn 0-based column index (e.g. 55 = column BD)
     */
    public static void updateTestStatus(int rowIndex, String status, int statusColumn) {

        File file = new File(VARIABLES.EXCEL_FILE_PATH);

        // ─── FIX 3: Synchronize on a static lock so parallel threads are safe ─
        synchronized (FILE_LOCK) {

            Workbook workBook = null;

            // ─── FIX 4: Use separate try blocks so read and write are clean ───
            try (FileInputStream fis = new FileInputStream(file)) {
                workBook = WorkbookFactory.create(fis);
            } catch (IOException e) {
                System.err.println("❌ Error reading Excel for row " + rowIndex + ": " + e.getMessage());
                e.printStackTrace();
                return;
            }

            try {
                Sheet sheet = workBook.getSheet(VARIABLES.SHEET_NAME);

                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    row = sheet.createRow(rowIndex);
                }

                Cell statusCell = row.getCell(statusColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                statusCell.setCellValue(status);

                // ─── FIX 2: Reuse styles — never call createCellStyle() repeatedly ───
                CellStyle style = getOrCreateStyle(workBook, status);
                statusCell.setCellStyle(style);

                sheet.autoSizeColumn(statusColumn);

                // ─── FIX 1: Write then close workbook inside finally ──────────────
                try (FileOutputStream fos = new FileOutputStream(file)) {
                    workBook.write(fos);
                }

                System.out.println("✅ Excel updated: Row " + rowIndex + " = " + status);

            } catch (Exception e) {
                System.err.println("❌ Error updating Excel status for row " + rowIndex + ": " + e.getMessage());
                e.printStackTrace();
            } finally {
                // ─── FIX 1: Workbook is always closed, even if write() throws ────
                try {
                    workBook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    // ─── FIX 2: Cache styles per workbook instance to avoid the 64k limit ────
    //           Key = "PASS" / "FAIL" / "SKIP"
    private static final Map<String, CellStyle> styleCache = new HashMap<>();

    private static CellStyle getOrCreateStyle(Workbook workBook, String status) {

        // Each time the file is re-opened a new Workbook instance is created,
        // so we key on both the workbook identity and the status string.
        String cacheKey = System.identityHashCode(workBook) + "_" + status.toUpperCase();

        if (styleCache.containsKey(cacheKey)) {
            return styleCache.get(cacheKey);
        }

        CellStyle style = workBook.createCellStyle();
        Font font = workBook.createFont();

        switch (status.toUpperCase()) {
            case "PASS":
                style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                font.setColor(IndexedColors.WHITE.getIndex());
                break;
            case "FAIL":
                style.setFillForegroundColor(IndexedColors.RED.getIndex());
                font.setColor(IndexedColors.WHITE.getIndex());
                break;
            case "SKIP":
                style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                font.setColor(IndexedColors.BLACK.getIndex());
                break;
            default:
                style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
                font.setColor(IndexedColors.BLACK.getIndex());
                break;
        }

        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(font);

        styleCache.put(cacheKey, style);
        return style;
    }
}