package it.firegloves.mempoi.datapostelaboration.mergedregions;

import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.datapostelaboration.mempoicolumn.mergedregions.StreamApiMergedRegionsStep;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import org.apache.commons.lang3.tuple.ImmutablePair;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.lang.reflect.Method;
import java.sql.PreparedStatement;

import static org.junit.Assert.assertEquals;

public class StreamApiMergedRegionsStepTest {

    private StreamApiMergedRegionsStep<String> step;

    private SXSSFSheet sheet;

    @Mock
    private PreparedStatement prepStmt;

    private final String[] cellValues = {"test", "fuck yeah", "fuck yeah", "oh wonderful"};

    @Before
    public void beforeTest() {
        MockitoAnnotations.initMocks(this);

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt).withSheetName("test name")
                .build();

        SXSSFWorkbook workbook = new SXSSFWorkbook();
        this.sheet = workbook.createSheet(mempoiSheet.getSheetName());
        CellStyle cellStyle = workbook.createCellStyle();

        int colInd = 1;
        this.step = new StreamApiMergedRegionsStep<>(cellStyle, colInd, workbook, mempoiSheet);

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
            Method mergeRegionMethod = StreamApiMergedRegionsStep.class.getDeclaredMethod("mergeRegion", Pair.class);
            mergeRegionMethod.setAccessible(true);

            mergeRegionMethod.invoke(this.step, new ImmutablePair<>(0, 1));
            assertEquals("Merged regions num 1", 1, this.sheet.getNumMergedRegions());

            mergeRegionMethod.invoke(this.step, new ImmutablePair<>(2, 5));
            assertEquals("Merged regions num 2", 2, this.sheet.getNumMergedRegions());

            mergeRegionMethod.invoke(this.step, new ImmutablePair<>(10, 5));
            assertEquals("Merged regions num 2", 2, this.sheet.getNumMergedRegions());

            mergeRegionMethod.invoke(this.step, new ImmutablePair<>(-4, -1));
            assertEquals("Merged regions num 2", 2, this.sheet.getNumMergedRegions());

            mergeRegionMethod.invoke(this.step, new ImmutablePair<>(0, 0));
            assertEquals("Merged regions num 2", 2, this.sheet.getNumMergedRegions());

            mergeRegionMethod.invoke(this.step, new ImmutablePair<>(5, 5));
            assertEquals("Merged regions num 2", 2, this.sheet.getNumMergedRegions());

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


}
