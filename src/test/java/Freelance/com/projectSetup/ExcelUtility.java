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

    // ─── Single static lock for all file access across threads ───────────────
    private static final Object FILE_LOCK = new Object();

    private static final short COLOR_GREEN  = IndexedColors.GREEN.getIndex();
    private static final short COLOR_RED    = IndexedColors.RED.getIndex();
    private static final short COLOR_YELLOW = IndexedColors.YELLOW.getIndex();
    private static final short COLOR_WHITE  = IndexedColors.WHITE.getIndex();
    private static final short COLOR_BLACK  = IndexedColors.BLACK.getIndex();

    // ─────────────────────────────────────────────────────────────────────────
    //  DATA PROVIDER
    // ─────────────────────────────────────────────────────────────────────────

    @DataProvider(name = "excelData")
    public static Object[][] getExcelData() {
        File file = new File(VARIABLES.EXCEL_FILE_PATH);

        try (FileInputStream excelFile = new FileInputStream(file);
             Workbook workBook = WorkbookFactory.create(excelFile)) {

            Sheet sheet       = workBook.getSheet(VARIABLES.SHEET_NAME);
            int totalRows     = sheet.getLastRowNum();
            int totalColumns  = sheet.getRow(0).getLastCellNum();
            DataFormatter fmt = new DataFormatter();
            ArrayList<Object[]> dataList = new ArrayList<>();

            for (int i = 1; i <= totalRows; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Object[] rowData = new Object[totalColumns + 1];
                rowData[0] = i; // 1-based row index

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
            throw new RuntimeException("Error reading excel file: " + e.getMessage());
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  STATUS WRITER — default column BD (index 55)
    // ─────────────────────────────────────────────────────────────────────────

    /** Convenience overload — writes to column BD (index 55). */
    public static void updateTestStatus(int rowIndex, String status) {
        updateTestStatus(rowIndex, status, 55);
    }

    /**
     * Writes PASS / FAIL / SKIP with colour to the given column.
     *
     * Why the previous version corrupted and this one does not:
     *
     *  CORRUPTION CAUSE 1 — autoSizeColumn()
     *    autoSizeColumn() forces POI to evaluate every formula/cell in the
     *    column to measure text width. On large .xlsx sheets this leaves
     *    internal shared-strings XML in an inconsistent state, producing
     *    "We found a problem with some content" on next open.
     *    FIX: replaced with sheet.setColumnWidth() — a plain integer, no eval.
     *
     *  CORRUPTION CAUSE 2 — new CellStyle() on every call (old doc-19 version)
     *    The overloaded method in doc-19 called workBook.createCellStyle() every
     *    single invocation. Excel's hard limit is 64 000 styles; exceeding it
     *    silently corrupts the file.
     *    FIX: findOrCreateStyle() scans existing styles and reuses matches.
     *
     *  CORRUPTION CAUSE 3 — stream not closed before write
     *    If FileInputStream is still open when FileOutputStream writes the same
     *    file, POI gets a partial/locked read on the next call.
     *    FIX: read stream is closed by try-with-resources BEFORE write opens.
     *
     *  CORRUPTION CAUSE 4 — Workbook not closed on exception
     *    An unclosed XSSFWorkbook holds temp files; if write() throws, the
     *    output file is truncated.
     *    FIX: workBook.close() is in a finally block — always runs.
     */
    public static void updateTestStatus(int rowIndex, String status, int statusColumn) {

        File file = new File(VARIABLES.EXCEL_FILE_PATH);

        synchronized (FILE_LOCK) {

            Workbook workBook = null;

            // ── Step 1: read ─────────────────────────────────────────────────
            // Stream is closed by try-with-resources before write opens the file.
            try (FileInputStream fis = new FileInputStream(file)) {
                workBook = WorkbookFactory.create(fis);
            } catch (IOException e) {
                System.err.println("❌ Cannot read Excel for row " + rowIndex + ": " + e.getMessage());
                e.printStackTrace();
                return;
            }

            // ── Step 2: edit in memory ────────────────────────────────────────
            try {
                Sheet sheet = workBook.getSheet(VARIABLES.SHEET_NAME);

                Row row = sheet.getRow(rowIndex);
                if (row == null) row = sheet.createRow(rowIndex);

                Cell cell = row.getCell(statusColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                cell.setCellValue(status);
                cell.setCellStyle(findOrCreateStyle(workBook, status));

                // Fixed width instead of autoSizeColumn() — safe, no formula eval
                sheet.setColumnWidth(statusColumn, 3000); // ~10 chars wide

                // ── Step 3: write ─────────────────────────────────────────────
                try (FileOutputStream fos = new FileOutputStream(file)) {
                    workBook.write(fos);
                    fos.flush();
                }

                System.out.println("✅ Excel updated — Row " + rowIndex + " = " + status);

            } catch (Exception e) {
                System.err.println("❌ Error writing Excel row " + rowIndex + ": " + e.getMessage());
                e.printStackTrace();
            } finally {
                // ── Step 4: always close workbook ─────────────────────────────
                try { workBook.close(); } catch (IOException ignored) {}
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  STYLE HELPER
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Scans the workbook's existing style table before creating a new style.
     * Each workbook only ever needs 3 custom styles (PASS / FAIL / SKIP).
     * Without this guard, every updateTestStatus() call adds a new CellStyle,
     * hitting Excel's 64 000-style hard limit and corrupting the file.
     */
    private static CellStyle findOrCreateStyle(Workbook workBook, String status) {

        short bgColor   = COLOR_WHITE;
        short fontColor = COLOR_BLACK;

        switch (status.toUpperCase()) {
            case "PASS": bgColor = COLOR_GREEN;  fontColor = COLOR_WHITE; break;
            case "FAIL": bgColor = COLOR_RED;    fontColor = COLOR_WHITE; break;
            case "SKIP": bgColor = COLOR_YELLOW; fontColor = COLOR_BLACK; break;
        }

        final short bg       = bgColor;
        final short fontClr  = fontColor;

        // Reuse an existing style if it already has the right fill
        int numStyles = workBook.getNumCellStyles();
        for (int i = 0; i < numStyles; i++) {
            CellStyle existing = workBook.getCellStyleAt(i);
            if (existing.getFillPattern() == FillPatternType.SOLID_FOREGROUND
                    && existing.getFillForegroundColor() == bg) {
                return existing; // ✅ reuse — do NOT create a new one
            }
        }

        // First time this status is seen — create exactly one style for it
        CellStyle style = workBook.createCellStyle();
        Font font       = workBook.createFont();
        style.setFillForegroundColor(bg);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        font.setColor(fontClr);
        style.setFont(font);
        return style;
    }
}