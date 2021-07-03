package it.firegloves.mempoi.domain;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class MempoiSheetMetadataTest {

    public static final int SHEET_INDEX = 0;
    public static final String SHEET_NAME = "My sheet";
    public static final int TOTAL_ROWS = 13;
    public static final int FIRST_ROW = 2;
    public static final int LAST_ROW = 3;
    public static final int TOTAL_COLS = 23;
    public static final int FIRST_COL = 5;
    public static final int LAST_COL = 7;

    @Test
    public void fullTest() {
        MempoiSheetMetadata current = new MempoiSheetMetadata()
                .setSheetIndex(SHEET_INDEX)
                .setSheetName(SHEET_NAME)
                .setFirstDataColumn(FIRST_COL)
                .setLastDataColumn(LAST_COL)
                .setTotalColumns(TOTAL_COLS)
                .setFirstDataRow(FIRST_ROW)
                .setLastDataRow(LAST_ROW)
                .setTotalRows(TOTAL_ROWS);

        assertEquals(SHEET_INDEX, current.getSheetIndex());
        assertEquals(SHEET_NAME, current.getSheetName());
        assertEquals(FIRST_COL, current.getFirstDataColumn());
        assertEquals(LAST_COL, current.getLastDataColumn());
        assertEquals(TOTAL_COLS, current.getTotalColumns());
        assertEquals(FIRST_ROW, current.getFirstDataRow());
        assertEquals(LAST_ROW, current.getLastDataRow());
        assertEquals(TOTAL_ROWS, current.getTotalRows());
    }
}
