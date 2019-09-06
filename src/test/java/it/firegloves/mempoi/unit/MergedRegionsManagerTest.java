package it.firegloves.mempoi.unit;

import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.dataelaborationpipeline.mempoicolumn.StreamApiElaborationStep;
import it.firegloves.mempoi.dataelaborationpipeline.mempoicolumn.mergedregions.MergedRegionsManager;
import it.firegloves.mempoi.dataelaborationpipeline.mempoicolumn.mergedregions.StreamApiMergedRegionsStep;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import it.firegloves.mempoi.util.Errors;
import org.apache.commons.lang3.tuple.ImmutablePair;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hamcrest.core.IsInstanceOf;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.MockitoAnnotations;

import java.lang.reflect.Method;
import java.sql.PreparedStatement;
import java.util.List;
import java.util.Optional;

import static org.junit.Assert.*;
import static org.mockito.Mockito.when;

public class MergedRegionsManagerTest {

    @Mock
    private Row row1, row2, row3, row4;
    @Mock
    private Cell cell1, cell2, cell3, cell4;
    private MergedRegionsManager<String> mergedRegionsManager;

    private Workbook workbook;
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

        this.workbook = new XSSFWorkbook();
        this.sheet = this.workbook.createSheet();
        this.cellStyle = this.workbook.createCellStyle();
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
    public void mergeRegionFailTest() {

        int firstRow = 5, lastRow = 4, colInd = 1;

        boolean merged = this.mergedRegionsManager.mergeRegion(this.sheet, this.cellStyle, firstRow, lastRow, colInd);

        assertEquals("Merged regions false", false, merged);
    }

    // TODO add test for sheet null?
}