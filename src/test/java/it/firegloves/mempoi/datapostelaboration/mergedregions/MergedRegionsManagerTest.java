package it.firegloves.mempoi.datapostelaboration.mergedregions;

import static org.junit.Assert.assertEquals;
import static org.mockito.Mockito.when;

import it.firegloves.mempoi.datapostelaboration.mempoicolumn.mergedregions.MergedRegionsManager;
import it.firegloves.mempoi.exception.MempoiException;
import java.util.Optional;
import org.apache.commons.lang3.tuple.ImmutablePair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

public class MergedRegionsManagerTest {

    @Mock
    private Row row1, row2, row3, row4;
    @Mock
    private Cell cell1, cell2, cell3, cell4;
    private MergedRegionsManager<String> mergedRegionsManager;

    private Sheet sheet;
    private CellStyle cellStyle;

    @Before
    public void beforeTest() {
        this.mergedRegionsManager = new MergedRegionsManager<>();
        MockitoAnnotations.initMocks(this);

        this.initCellRow(cell1, row1, 0, "test");
        this.initCellRow(cell2, row2, 1, "fuck yeah");
        this.initCellRow(cell3, row3, 2, "fuck yeah");
        this.initCellRow(cell4, row4, 3, "oh wonderful");

        Workbook workbook = new XSSFWorkbook();
        this.sheet = workbook.createSheet();
        this.cellStyle = workbook.createCellStyle();
    }


    private void initCellRow(Cell cell, Row row, int rowNum, String value) {
        when(row.getRowNum()).thenReturn(rowNum);
        when(cell.getRow()).thenReturn(row);
        when(cell.getStringCellValue()).thenReturn(value);
    }


    @Test
    public void performAndCloseAnalysisTest() {

        Optional<ImmutablePair<Integer, Integer>> opt1 = this.mergedRegionsManager.performAnalysis(cell1, cell1.getStringCellValue());
        Optional<ImmutablePair<Integer, Integer>> opt2 = this.mergedRegionsManager.performAnalysis(cell2, cell2.getStringCellValue());
        Optional<ImmutablePair<Integer, Integer>> opt3 = this.mergedRegionsManager.performAnalysis(cell3, cell3.getStringCellValue());
        Optional<ImmutablePair<Integer, Integer>> opt4 = this.mergedRegionsManager.performAnalysis(cell4, cell4.getStringCellValue());
        Optional<ImmutablePair<Integer, Integer>> opt5 = this.mergedRegionsManager.closeAnalysis(5);

        assertEquals("Opt 1 is not present", false, opt1.isPresent());
        assertEquals("Opt 2 is not present", false, opt2.isPresent());
        assertEquals("Opt 3 is not present", false, opt3.isPresent());
        assertEquals("Opt 4 is present", true, opt4.isPresent());
        assertEquals("Opt 4 is present", false, opt5.isPresent());

    }

    @Test(expected = MempoiException.class)
    public void performAnalysisNullCell() {

        this.mergedRegionsManager.performAnalysis(null, cell1.getStringCellValue());
    }

    @Test(expected = MempoiException.class)
    public void performAnalysisNullValue() {

        this.mergedRegionsManager.performAnalysis(cell1, null);
    }

    @Test(expected = MempoiException.class)
    public void performAnalysisNullCellAndValue() {

        this.mergedRegionsManager.performAnalysis(null, null);
    }


    @Test
    public void mergeRegionTest() {

        int firstRow = 1, lastRow = 4, colInd = 1;

        for (int i = 0; i < 10; i++) {
            Row r = this.sheet.createRow(i);
            for (int c=0; c<20; c++) {
                r.createCell(c);
            }
        }

        boolean merged = this.mergedRegionsManager.mergeRegion(this.sheet, this.cellStyle, firstRow, lastRow, colInd);
        CellRangeAddress ca = this.sheet.getMergedRegions().get(0);

        assertEquals("Merged regions true", true, merged);
        assertEquals("Merged regions num = 1", 1, this.sheet.getNumMergedRegions());
        assertEquals("Merged regions first row", firstRow, ca.getFirstRow());
        assertEquals("Merged regions last row", lastRow, ca.getLastRow());
        assertEquals("Merged regions first column", colInd, ca.getFirstColumn());
        assertEquals("Merged regions last column", colInd, ca.getLastColumn());
        assertEquals("Merged regions cell style", cellStyle, this.sheet.getRow(firstRow).getCell(colInd).getCellStyle());


        merged = this.mergedRegionsManager.mergeRegion(this.sheet, this.cellStyle, -5, -2, colInd);

        assertEquals("Merged regions true", false, merged);
        assertEquals("Merged regions num = 1", 1, this.sheet.getNumMergedRegions());


        merged = this.mergedRegionsManager.mergeRegion(this.sheet, this.cellStyle, 0, 0, colInd);

        assertEquals("Merged regions true", false, merged);
        assertEquals("Merged regions num = 1", 1, this.sheet.getNumMergedRegions());


        merged = this.mergedRegionsManager.mergeRegion(this.sheet, this.cellStyle, 3, 3, colInd);

        assertEquals("Merged regions true", false, merged);
        assertEquals("Merged regions num = 1", 1, this.sheet.getNumMergedRegions());
    }


    @Test
    public void mergeRegionLastRowLowerThanFirst() {

        int firstRow = 5, lastRow = 4, colInd = 1;

        boolean merged = this.mergedRegionsManager.mergeRegion(this.sheet, this.cellStyle, firstRow, lastRow, colInd);

        assertEquals("Merged regions false", false, merged);
    }

    @Test(expected = MempoiException.class)
    public void mergeRegionNullSheet() {

        int firstRow = 5, lastRow = 4, colInd = 1;

        this.mergedRegionsManager.mergeRegion(null, this.cellStyle, firstRow, lastRow, colInd);
    }

    @Test
    public void mergeRegionNullCellStyle() {

        int firstRow = 5, lastRow = 4, colInd = 1;

        this.mergedRegionsManager.mergeRegion(this.sheet, null, firstRow, lastRow, colInd);
    }

    @Test
    public void mergeRegionNegativeFirstRow() {

        int firstRow = -5, lastRow = 4, colInd = 1;

        this.mergedRegionsManager.mergeRegion(this.sheet, this.cellStyle, firstRow, lastRow, colInd);
    }

    @Test
    public void mergeRegionNegativeColInd() {

        int firstRow = -5, lastRow = 4, colInd = -1;

        this.mergedRegionsManager.mergeRegion(this.sheet, this.cellStyle, firstRow, lastRow, colInd);
    }


    @Test
    public void closeAnalysis() {

        Optional<ImmutablePair<Integer, Integer>> optPair = this.mergedRegionsManager.closeAnalysis(5);

        assertEquals((int) optPair.get().left, 0);
        assertEquals((int) optPair.get().right, 5);
    }

}
