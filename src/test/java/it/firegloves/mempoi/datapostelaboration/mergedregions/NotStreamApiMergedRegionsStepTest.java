package it.firegloves.mempoi.datapostelaboration.mergedregions;

import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.datapostelaboration.mempoicolumn.mergedregions.NotStreamApiMergedRegionsStep;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.lang.reflect.Field;
import java.sql.PreparedStatement;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.mockito.Mockito.when;

public class NotStreamApiMergedRegionsStepTest {

    @Mock
    private Row row1, row2, row3, row4;
    @Mock
    private Cell cell1, cell2, cell3, cell4;
    private NotStreamApiMergedRegionsStep<String> step;

    private Workbook workbook;
    private Sheet sheet;
    private CellStyle cellStyle;

    private int colInd = 1;

    private List mergedRegionsLimits;

    @Mock
    private PreparedStatement prepStmt;


    @Before
    public void beforeTest() {
        MockitoAnnotations.initMocks(this);

        this.initCellRow(cell1, row1, 0, "test");
        this.initCellRow(cell2, row2, 1, "fuck yeah");
        this.initCellRow(cell3, row3, 2, "fuck yeah");
        this.initCellRow(cell4, row4, 3, "oh wonderful");

        this.workbook = new XSSFWorkbook();
        this.sheet = this.workbook.createSheet();
        this.cellStyle = this.workbook.createCellStyle();

        this.step = new NotStreamApiMergedRegionsStep<>(this.cellStyle, this.colInd);

        try {
            Field mergedRegionsLimitsField = NotStreamApiMergedRegionsStep.class.getDeclaredField("mergedRegionsLimits");
            mergedRegionsLimitsField.setAccessible(true);
            mergedRegionsLimits = (List) mergedRegionsLimitsField.get(this.step);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    private void initCellRow(Cell cell, Row row, int rowNum, String value) {
        when(row.getRowNum()).thenReturn(rowNum);
        when(cell.getRow()).thenReturn(row);
        when(cell.getStringCellValue()).thenReturn(value);
    }


    @Test
    public void performAndCloseAnalysisTest() {

        this.step.performAnalysis(cell1, cell1.getStringCellValue());
        assertEquals("0 mergedRegionsLimits", 0, this.mergedRegionsLimits.size());

        this.step.performAnalysis(cell2, cell2.getStringCellValue());
        assertEquals("0 mergedRegionsLimits", 0, this.mergedRegionsLimits.size());

        this.step.performAnalysis(cell3, cell3.getStringCellValue());
        assertEquals("0 mergedRegionsLimits", 0, this.mergedRegionsLimits.size());

        this.step.performAnalysis(cell4, cell4.getStringCellValue());
        assertEquals("1 mergedRegionsLimits", 1, this.mergedRegionsLimits.size());

        this.step.closeAnalysis(5);
        assertEquals("1 mergedRegionsLimits", 1, this.mergedRegionsLimits.size());
    }


    @Test
    public void mergeRegionTest() {

        String name = "test name";
        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt).withSheetName(name).build();
        this.sheet = this.workbook.createSheet(mempoiSheet.getSheetName());

        for (int i = 0; i < 10; i++) {
            Row r = this.sheet.createRow(i);
            for (int c = 0; c < 20; c++) {
                r.createCell(c);
            }
        }

        this.step.performAnalysis(cell1, cell1.getStringCellValue());
        this.step.performAnalysis(cell2, cell2.getStringCellValue());
        this.step.performAnalysis(cell3, cell3.getStringCellValue());
        this.step.performAnalysis(cell4, cell4.getStringCellValue());
        this.step.closeAnalysis(5);
        this.step.execute(mempoiSheet, this.workbook);

        List<CellRangeAddress> cellRangeAddresseList = this.sheet.getMergedRegions();

        assertEquals("Merged regions num = 1", 1, this.sheet.getNumMergedRegions());

        this.checkMergedRegion(1, 2, cellRangeAddresseList.get(0));
    }



    private void checkMergedRegion(int firstRow, int lastRow, CellRangeAddress ca) {
        assertEquals("Merged regions first row", firstRow, ca.getFirstRow());
        assertEquals("Merged regions last row", lastRow, ca.getLastRow());
        assertEquals("Merged regions first column", colInd, ca.getFirstColumn());
        assertEquals("Merged regions last column", colInd, ca.getLastColumn());
        assertEquals("Merged regions cell style", cellStyle, this.sheet.getRow(firstRow).getCell(colInd).getCellStyle());
    }


    @Test(expected = MempoiException.class)
    public void executNullSheet() {

        this.step.execute(null, this.workbook);
    }

    @Test(expected = MempoiException.class)
    public void executeNullWorkbook() {

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt).withSheetName("test").build();
        this.step.execute(mempoiSheet, null);
    }
}
