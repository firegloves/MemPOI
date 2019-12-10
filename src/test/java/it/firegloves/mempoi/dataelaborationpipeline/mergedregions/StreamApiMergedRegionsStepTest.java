package it.firegloves.mempoi.dataelaborationpipeline.mergedregions;

import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.dataelaborationpipeline.mempoicolumn.StreamApiElaborationStep;
import it.firegloves.mempoi.dataelaborationpipeline.mempoicolumn.mergedregions.NotStreamApiMergedRegionsStep;
import it.firegloves.mempoi.dataelaborationpipeline.mempoicolumn.mergedregions.StreamApiMergedRegionsStep;
import it.firegloves.mempoi.domain.MempoiSheet;
import org.apache.commons.lang3.tuple.ImmutablePair;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.sql.PreparedStatement;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.mockito.Mockito.when;

public class StreamApiMergedRegionsStepTest {

    private StreamApiMergedRegionsStep<String> step;

    private SXSSFWorkbook workbook;
    private SXSSFSheet sheet;
    private CellStyle cellStyle;
    private MempoiSheet mempoiSheet;

    private int colInd = 1;
    private Method mergeRegionMethod;

    @Mock
    private PreparedStatement prepStmt;

    private String[] cellValues = {"test", "fuck yeah", "fuck yeah", "oh wonderful"};

    @Before
    public void beforeTest() {
        MockitoAnnotations.initMocks(this);

        this.mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt).withSheetName("test name").build();

        this.workbook = new SXSSFWorkbook();
        this.sheet = this.workbook.createSheet(mempoiSheet.getSheetName());
        this.cellStyle = this.workbook.createCellStyle();

        this.step = new StreamApiMergedRegionsStep<>(this.cellStyle, this.colInd, this.workbook, this.mempoiSheet);

        for (int i = 0; i < 4; i++) {
            Row r = this.sheet.createRow(i);
            for (int c = 0; c < 20; c++) {
                Cell cell = r.createCell(c);
                cell.setCellValue(cellValues[i % 4]);
            }
        }
    }


    @Test
    public void performAndCloseAnalysisTest() {

        Cell c = this.sheet.getRow(0).getCell(0);
        this.step.performAnalysis(c, c.getStringCellValue());
        assertEquals("0 mergedRegionsLimits", 0, this.sheet.getNumMergedRegions());

        c = this.sheet.getRow(1).getCell(0);
        this.step.performAnalysis(c, c.getStringCellValue());
        assertEquals("1 mergedRegionsLimits", 0, this.sheet.getNumMergedRegions());

        c = this.sheet.getRow(2).getCell(0);
        this.step.performAnalysis(c, c.getStringCellValue());
        assertEquals("0 mergedRegionsLimits", 0, this.sheet.getNumMergedRegions());

        c = this.sheet.getRow(3).getCell(0);
        this.step.performAnalysis(c, c.getStringCellValue());
        assertEquals("1 mergedRegionsLimits", 1, this.sheet.getNumMergedRegions());

        this.step.closeAnalysis(5);
        assertEquals("1 mergedRegionsLimits", 1, this.sheet.getNumMergedRegions());
    }


    @Test
    public void mergeRegionTest() {

        try {
            this.mergeRegionMethod = StreamApiMergedRegionsStep.class.getDeclaredMethod("mergeRegion", Pair.class);
            this.mergeRegionMethod.setAccessible(true);

            this.mergeRegionMethod.invoke(this.step, new ImmutablePair<>(0, 1));
            assertEquals("Merged regions num 1", 1, this.sheet.getNumMergedRegions());

            this.mergeRegionMethod.invoke(this.step, new ImmutablePair<>(2, 5));
            assertEquals("Merged regions num 2", 2, this.sheet.getNumMergedRegions());

            this.mergeRegionMethod.invoke(this.step, new ImmutablePair<>(10, 5));
            assertEquals("Merged regions num 2", 2, this.sheet.getNumMergedRegions());

            this.mergeRegionMethod.invoke(this.step, new ImmutablePair<>(-4, -1));
            assertEquals("Merged regions num 2", 2, this.sheet.getNumMergedRegions());

            this.mergeRegionMethod.invoke(this.step, new ImmutablePair<>(0, 0));
            assertEquals("Merged regions num 2", 2, this.sheet.getNumMergedRegions());

            this.mergeRegionMethod.invoke(this.step, new ImmutablePair<>(5, 5));
            assertEquals("Merged regions num 2", 2, this.sheet.getNumMergedRegions());

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }


}