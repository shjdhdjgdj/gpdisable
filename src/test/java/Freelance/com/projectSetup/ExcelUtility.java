package Freelance.com.projectSetup;

import config.VARIABLES;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.*;
import org.testng.annotations.DataProvider;

public class ExcelUtility {

    // ─── Single static lock — safe across parallel TestNG threads ─────────────
    private static final Object FILE_LOCK = new Object();

    // Pre-cache colour indices for readability
    private static final short COLOR_GREEN  = IndexedColors.GREEN.getIndex();
    private static final short COLOR_RED    = IndexedColors.RED.getIndex();
    private static final short COLOR_YELLOW = IndexedColors.YELLOW.getIndex();
    private static final short COLOR_WHITE  = IndexedColors.WHITE.getIndex();
    private static final short COLOR_BLACK  = IndexedColors.BLACK.getIndex();

    // ─────────────────────────────────────────────────────────────────────────
    //  DATA PROVIDER
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Reads all data rows from the configured Excel sheet.
     * Row index (1-based) is stored at position 0 of each row array so that
     * updateTestStatus() can write back to the exact source row.
     */
    @DataProvider(name = "excelData")
    public static Object[][] getExcelData() {
        File file = new File(VARIABLES.EXCEL_FILE_PATH);

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workBook = WorkbookFactory.create(fis)) {

            Sheet sheet       = workBook.getSheet(VARIABLES.SHEET_NAME);
            int totalRows     = sheet.getLastRowNum();
            int totalColumns  = sheet.getRow(0).getLastCellNum();
            DataFormatter fmt = new DataFormatter();
            ArrayList<Object[]> dataList = new ArrayList<>();

            for (int i = 1; i <= totalRows; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue; // skip blank rows

                Object[] rowData = new Object[totalColumns + 1]; // +1 for row index at [0]
                rowData[0] = i; // 1-based Excel row number

                for (int j = 0; j < totalColumns; j++) {
                    Cell cell = row.getCell(j);
                    rowData[j + 1] = fmt.formatCellValue(cell);
                }
                dataList.add(rowData);
            }

            Object[][] data = new Object[dataList.size()][];
            dataList.toArray(data);
            return data;

        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("Error reading Excel file: " + e.getMessage());
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  STATUS WRITER  (default column BD = index 55)
    // ─────────────────────────────────────────────────────────────────────────

    /** Convenience overload — writes to column BD (index 55). */
    public static void updateTestStatus(int rowIndex, String status) {
        updateTestStatus(rowIndex, status, 55);
    }

    /**
     * Writes PASS / FAIL / SKIP with colour coding to the given column.
     *
     * FIX: Old code called workBook.createCellStyle() on every invocation.
     * Excel has a hard limit of 64 000 cell styles per workbook; hitting it
     * corrupts the file. This version scans existing styles and reuses a
     * matching one — creating at most 3 new styles ever (one per status).
     *
     * @param rowIndex     1-based Excel row number (as stored at data[0])
     * @param status       "PASS", "FAIL", or "SKIP"
     * @param statusColumn 0-based column index to write into
     */
    public static void updateTestStatus(int rowIndex, String status, int statusColumn) {

        File file = new File(VARIABLES.EXCEL_FILE_PATH);

        synchronized (FILE_LOCK) {

            Workbook workBook = null;

            // Step 1 — read (stream closed before write to avoid file-lock conflicts)
            try (FileInputStream fis = new FileInputStream(file)) {
                workBook = WorkbookFactory.create(fis);
            } catch (IOException e) {
                System.err.println("❌ Cannot read Excel for row " + rowIndex + ": " + e.getMessage());
                e.printStackTrace();
                return;
            }

            // Step 2 — edit in memory
            try {
                Sheet sheet = workBook.getSheet(VARIABLES.SHEET_NAME);

                Row row = sheet.getRow(rowIndex);
                if (row == null) row = sheet.createRow(rowIndex);

                Cell cell = row.getCell(statusColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                cell.setCellValue(status);
                cell.setCellStyle(findOrCreateStyle(workBook, status)); // ✅ reuse styles
                sheet.autoSizeColumn(statusColumn);

                // Step 3 — write (try-with-resources guarantees fos.close())
                try (FileOutputStream fos = new FileOutputStream(file)) {
                    workBook.write(fos);
                    fos.flush();
                }

                System.out.println("✅ Excel updated — Row " + rowIndex + " = " + status);

            } catch (Exception e) {
                System.err.println("❌ Error writing Excel row " + rowIndex + ": " + e.getMessage());
                e.printStackTrace();
            } finally {
                // Step 4 — always close workbook
                try { workBook.close(); } catch (IOException ignored) {}
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  STYLE HELPER
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Scans the workbook's style table and returns a matching style if one
     * already exists. Creates a new style only when necessary (first time
     * each status is seen). This keeps the total new-style count at ≤ 3,
     * well within Excel's 64 k limit.
     */
    private static CellStyle findOrCreateStyle(Workbook workBook, String status) {

        short bgColor   = COLOR_WHITE;
        short fontColor = COLOR_BLACK;

        switch (status.toUpperCase()) {
            case "PASS": bgColor = COLOR_GREEN;  fontColor = COLOR_WHITE; break;
            case "FAIL": bgColor = COLOR_RED;    fontColor = COLOR_WHITE; break;
            case "SKIP": bgColor = COLOR_YELLOW; fontColor = COLOR_BLACK; break;
        }

        final short finalBg   = bgColor;
        final short finalFont = fontColor;

        // Scan existing styles — reuse on match
        int numStyles = workBook.getNumCellStyles();
        for (int i = 0; i < numStyles; i++) {
            CellStyle existing = workBook.getCellStyleAt(i);
            if (existing.getFillPattern() == FillPatternType.SOLID_FOREGROUND
                    && existing.getFillForegroundColor() == finalBg) {
                return existing; // ✅ reuse — do NOT create another
            }
        }

        // First occurrence of this status — create exactly one style for it
        CellStyle style = workBook.createCellStyle();
        Font font       = workBook.createFont();
        style.setFillForegroundColor(finalBg);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        font.setColor(finalFont);
        style.setFont(font);
        return style;
    }
}