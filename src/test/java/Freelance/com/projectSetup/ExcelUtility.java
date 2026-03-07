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

    /** Update test status — default column BD (index 55) */
    public static void updateTestStatus(int rowIndex, String status) {
        updateTestStatus(rowIndex, status, 55);
    }

    /**
     * Update test status with a custom column.
     *
     * @param rowIndex     1-based Excel row number
     * @param status       PASS / FAIL / SKIP
     * @param statusColumn 0-based column index
     */
    public static void updateTestStatus(int rowIndex, String status, int statusColumn) {

        File file = new File(VARIABLES.EXCEL_FILE_PATH);

        synchronized (FILE_LOCK) {

            Workbook workBook = null;

            // ── Step 1: Read file (stream closed before write) ───────────────
            try (FileInputStream fis = new FileInputStream(file)) {
                workBook = WorkbookFactory.create(fis);
            } catch (IOException e) {
                System.err.println("❌ Error reading Excel row " + rowIndex + ": " + e.getMessage());
                e.printStackTrace();
                return;
            }

            // ── Step 2: Edit in memory, then write ───────────────────────────
            try {
                Sheet sheet = workBook.getSheet(VARIABLES.SHEET_NAME);

                Row row = sheet.getRow(rowIndex);
                if (row == null) row = sheet.createRow(rowIndex);

                Cell statusCell = row.getCell(statusColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                statusCell.setCellValue(status);

                // FIX: Reuse existing styles — never create a new one if a matching
                //      style already exists in the workbook style table.
                //      Creating a new CellStyle every call hits Excel's 64k limit fast.
                statusCell.setCellStyle(findOrCreateStyle(workBook, status));

                sheet.autoSizeColumn(statusColumn);

                // ── Step 3: Write — fos always closed by try-with-resources ──
                try (FileOutputStream fos = new FileOutputStream(file)) {
                    workBook.write(fos);
                    fos.flush(); // flush OS buffers fully before releasing lock
                }

                System.out.println("✅ Excel updated: Row " + rowIndex + " = " + status);

            } catch (Exception e) {
                System.err.println("❌ Error updating Excel row " + rowIndex + ": " + e.getMessage());
                e.printStackTrace();
            } finally {
                // ── Step 4: Workbook always closed even if write() throws ────
                try { workBook.close(); } catch (IOException ignored) {}
            }
        }
    }

    /**
     * FIX: Scan the workbook's existing style table before creating a new style.
     * Each workbook only ever needs 3 styles (PASS/FAIL/SKIP).
     * Without this, every updateTestStatus() call creates a new CellStyle,
     * which hits Excel's 64k-style hard limit and corrupts the file.
     */
    private static CellStyle findOrCreateStyle(Workbook workBook, String status) {
        short bgColor   = COLOR_WHITE;
        short fontColor = COLOR_BLACK;

        switch (status.toUpperCase()) {
            case "PASS": bgColor = COLOR_GREEN;  fontColor = COLOR_WHITE; break;
            case "FAIL": bgColor = COLOR_RED;    fontColor = COLOR_WHITE; break;
            case "SKIP": bgColor = COLOR_YELLOW; fontColor = COLOR_BLACK; break;
        }

        // Scan existing styles — reuse if a match is found
        int numStyles = workBook.getNumCellStyles();
        for (int i = 0; i < numStyles; i++) {
            CellStyle existing = workBook.getCellStyleAt(i);
            if (existing.getFillForegroundColor() == bgColor
                    && existing.getFillPattern() == FillPatternType.SOLID_FOREGROUND) {
                return existing; // reuse, do NOT create new
            }
        }

        // First time this status is seen — create exactly one style for it
        CellStyle style = workBook.createCellStyle();
        Font font       = workBook.createFont();
        style.setFillForegroundColor(bgColor);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        font.setColor(fontColor);
        style.setFont(font);
        return style;
    }
}