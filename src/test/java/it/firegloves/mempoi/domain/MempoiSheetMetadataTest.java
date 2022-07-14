package it.firegloves.mempoi.domain;

import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnCols;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnHeaders;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnPivotTable;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnRows;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnSubfooterMetadata;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnTable;
import static org.junit.Assert.assertEquals;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class MempoiSheetMetadataTest {

    public static final int SHEET_INDEX = 0;
    public static final String SHEET_NAME = "My sheet";
    public static final int TOTAL_ROWS = 13;
    public static final int FIRST_ROW = 2;
    public static final int LAST_ROW = 3;
    public static final int TOTAL_COLS = 23;
    public static final int COLS_OFFSET = 5;
    public static final int ROWS_OFFSET = 7;
    public static final int FIRST_COL = 5;
    public static final int LAST_COL = 7;
    public static final int TOTAL_DATA_ROWS = 4;
    public static final int HEADER_ROW_INDEX = 1;
    public static final int SUBFOOTER_ROW_INDEX = 6;
    public static final int FIRST_TABLE_ROW = 8;
    public static final int LAST_TABLE_ROW = 9;
    public static final int FIRST_TABLE_COL = 10;
    public static final int LAST_TABLE_COL = 11;
    public static final int FIRST_PIVOT_TABLE_POS_ROW = 12;
    public static final int FIRST_PIVOT_TABLE_POS_COL = 13;
    public static final int FIRST_PIVOT_TABLE_ROW = 14;
    public static final int LAST_PIVOT_TABLE_ROW = 15;
    public static final int FIRST_PIVOT_TABLE_COL = 16;
    public static final int LAST_PIVOT_TABLE_COL = 17;

    @Test
    public void fullTest() {

        XSSFWorkbook workbook = new XSSFWorkbook();

        MempoiSheetMetadata current = new MempoiSheetMetadata()
                .setSpreadsheetVersion(workbook.getSpreadsheetVersion())
                .setSheetIndex(SHEET_INDEX)
                .setSheetName(SHEET_NAME)
                .setFirstDataColumn(FIRST_COL)
                .setLastDataColumn(LAST_COL)
                .setTotalColumns(TOTAL_COLS)
                .setColsOffset(COLS_OFFSET)
                .setRowsOffset(ROWS_OFFSET)
                .setFirstDataRow(FIRST_ROW)
                .setLastDataRow(LAST_ROW)
                .setTotalRows(TOTAL_ROWS)
                .setTotalDataRows(TOTAL_DATA_ROWS)
                .setHeaderRowIndex(HEADER_ROW_INDEX)
                .setSubfooterRowIndex(SUBFOOTER_ROW_INDEX)
                .setFirstTableRow(FIRST_TABLE_ROW)
                .setLastTableRow(LAST_TABLE_ROW)
                .setFirstTableColumn(FIRST_TABLE_COL)
                .setLastTableColumn(LAST_TABLE_COL)
                .setFirstPivotTablePositionRow(FIRST_PIVOT_TABLE_POS_ROW)
                .setFirstPivotTablePositionColumn(FIRST_PIVOT_TABLE_POS_COL)
                .setFirstPivotTableSourceRow(FIRST_PIVOT_TABLE_ROW)
                .setLastPivotTableSourceRow(LAST_PIVOT_TABLE_ROW)
                .setFirstPivotTableSourceColumn(FIRST_PIVOT_TABLE_COL)
                .setLastPivotTableSourceColumn(LAST_PIVOT_TABLE_COL);

        assertEquals(SHEET_NAME, current.getSheetName());
        assertEquals(SpreadsheetVersion.EXCEL2007, current.getSpreadsheetVersion());
        assertOnHeaders(current, HEADER_ROW_INDEX, "F2:H2");
        assertOnRows(current, TOTAL_ROWS, TOTAL_DATA_ROWS, FIRST_ROW, LAST_ROW, "F3:H4", ROWS_OFFSET);
        assertOnSubfooterMetadata(current, SUBFOOTER_ROW_INDEX, "F7:H7");
        assertOnCols(current, TOTAL_COLS, FIRST_COL, LAST_COL, COLS_OFFSET);
        assertOnTable(current, FIRST_TABLE_ROW, LAST_TABLE_ROW, FIRST_TABLE_COL, LAST_TABLE_COL, "K9:L10");
        assertOnPivotTable(current, FIRST_PIVOT_TABLE_POS_ROW, FIRST_PIVOT_TABLE_POS_COL, FIRST_PIVOT_TABLE_ROW,
                FIRST_PIVOT_TABLE_COL, LAST_PIVOT_TABLE_ROW, LAST_PIVOT_TABLE_COL, "N13", "Q15:R16");
    }
}
