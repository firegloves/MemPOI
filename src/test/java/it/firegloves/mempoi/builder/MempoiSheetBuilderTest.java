package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.config.MempoiConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.domain.footer.StandardMempoiFooter;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.RoseStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.sql.PreparedStatement;

import static org.junit.Assert.*;

public class MempoiSheetBuilderTest {

    @Mock
    private PreparedStatement prepStmt;

    @Before
    public void prepare() {
        MockitoAnnotations.initMocks(this);
    }


    @Test
    public void mempoiSheetBuilderFullPopulated() {

        String sheetName = "test name";
        String footerName = "test footer";
        String[] mergedCols = new String[]{"col1", "col2"};
        StyleTemplate styleTemplate = new RoseStyleTemplate();
        Workbook wb = new XSSFWorkbook();
        NumberSumSubFooter numberSumSubFooter = new NumberSumSubFooter();
        ForestStyleTemplate forestStyleTemplate = new ForestStyleTemplate();

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withStyleTemplate(forestStyleTemplate)
                .withPrepStmt(prepStmt)
                .withSheetName(sheetName)
                .withMempoiSubFooter(numberSumSubFooter)
                .withCommonDataCellStyle(styleTemplate.getCommonDataCellStyle(wb))
                .withDateCellStyle(styleTemplate.getDateCellStyle(wb))
                .withDatetimeCellStyle(styleTemplate.getDatetimeCellStyle(wb))
                .withHeaderCellStyle(styleTemplate.getHeaderCellStyle(wb))
                .withNumberCellStyle(styleTemplate.getNumberCellStyle(wb))
                .withSubFooterCellStyle(styleTemplate.getSubfooterCellStyle(wb))
                .withMempoiFooter(new StandardMempoiFooter(wb, footerName))
                .withMergedRegionColumns(mergedCols)
                .withWorkbook(wb)
                .withMempoiTableBuilder(TestHelper.getTestMempoiTableBuilder(wb))
                .withMempoiPivotTableBuilder(TestHelper.getTestMempoiPivotTableBuilder(wb))
                .build();

        assertEquals("Style template ForestTemplate", forestStyleTemplate, mempoiSheet.getStyleTemplate());
        assertEquals("Prepared Statement", prepStmt, mempoiSheet.getPrepStmt());
        assertEquals("Sheet name", sheetName, mempoiSheet.getSheetName());
        assertEquals("Subfooter", numberSumSubFooter, mempoiSheet.getMempoiSubFooter().get());
        AssertionHelper.validateCellStyle(styleTemplate.getCommonDataCellStyle(wb), mempoiSheet.getCommonDataCellStyle());
        AssertionHelper.validateCellStyle(styleTemplate.getDateCellStyle(wb), mempoiSheet.getDateCellStyle());
        AssertionHelper.validateCellStyle(styleTemplate.getDatetimeCellStyle(wb), mempoiSheet.getDatetimeCellStyle());
        AssertionHelper.validateCellStyle(styleTemplate.getHeaderCellStyle(wb), mempoiSheet.getHeaderCellStyle());
        AssertionHelper.validateCellStyle(styleTemplate.getNumberCellStyle(wb), mempoiSheet.getNumberCellStyle());
        AssertionHelper.validateCellStyle(styleTemplate.getSubfooterCellStyle(wb), mempoiSheet.getSubFooterCellStyle());
        assertEquals("footer text", footerName, mempoiSheet.getMempoiFooter().get().getCenterText());
        assertArrayEquals("merged cols", mergedCols, mempoiSheet.getMergedRegionColumns());
        assertEquals("workbook", wb, mempoiSheet.getWorkbook());

       AssertionHelper.validateMempoiTable(wb, mempoiSheet.getMempoiTable().get());
       AssertionHelper.validateMempoiPivotTable(wb, mempoiSheet.getMempoiPivotTable().get());
    }


    @Test
    public void mempoiSheetBuilderNoOverridenStyle() {

        Workbook wb = new XSSFWorkbook();
        ForestStyleTemplate forestStyleTemplate = new ForestStyleTemplate();

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withStyleTemplate(forestStyleTemplate)
                .withPrepStmt(prepStmt)
                .withWorkbook(wb)
                .build();

        assertEquals("Style template ForestTemplate", forestStyleTemplate, mempoiSheet.getStyleTemplate());
        AssertionHelper.validateCellStyle(forestStyleTemplate.getCommonDataCellStyle(wb), mempoiSheet.getCommonDataCellStyle());
        AssertionHelper.validateCellStyle(forestStyleTemplate.getDateCellStyle(wb), mempoiSheet.getDateCellStyle());
        AssertionHelper.validateCellStyle(forestStyleTemplate.getDatetimeCellStyle(wb), mempoiSheet.getDatetimeCellStyle());
        AssertionHelper.validateCellStyle(forestStyleTemplate.getHeaderCellStyle(wb), mempoiSheet.getHeaderCellStyle());
        AssertionHelper.validateCellStyle(forestStyleTemplate.getNumberCellStyle(wb), mempoiSheet.getNumberCellStyle());
        AssertionHelper.validateCellStyle(forestStyleTemplate.getSubfooterCellStyle(wb), mempoiSheet.getSubFooterCellStyle());
        assertEquals("workbook", wb, mempoiSheet.getWorkbook());
        assertFalse(mempoiSheet.getMempoiTable().isPresent());
    }


    @Test
    public void mempoiSheetBuilderOverridenStyle() {

        Workbook wb = new XSSFWorkbook();
        StyleTemplate styleTemplate = new RoseStyleTemplate();
        ForestStyleTemplate forestStyleTemplate = new ForestStyleTemplate();

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withStyleTemplate(forestStyleTemplate)
                .withCommonDataCellStyle(styleTemplate.getCommonDataCellStyle(wb))
                .withDateCellStyle(styleTemplate.getDateCellStyle(wb))
                .withWorkbook(wb)
                .withPrepStmt(prepStmt)
                .build();

        assertEquals("Style template ForestTemplate", forestStyleTemplate, mempoiSheet.getStyleTemplate());
        AssertionHelper.validateCellStyle(styleTemplate.getCommonDataCellStyle(wb), mempoiSheet.getCommonDataCellStyle());
        AssertionHelper.validateCellStyle(styleTemplate.getDateCellStyle(wb), mempoiSheet.getDateCellStyle());
        AssertionHelper.validateCellStyle(forestStyleTemplate.getDatetimeCellStyle(wb), mempoiSheet.getDatetimeCellStyle());
        AssertionHelper.validateCellStyle(forestStyleTemplate.getHeaderCellStyle(wb), mempoiSheet.getHeaderCellStyle());
        AssertionHelper.validateCellStyle(forestStyleTemplate.getNumberCellStyle(wb), mempoiSheet.getNumberCellStyle());
        AssertionHelper.validateCellStyle(forestStyleTemplate.getSubfooterCellStyle(wb), mempoiSheet.getSubFooterCellStyle());
        assertEquals("workbook", wb, mempoiSheet.getWorkbook());
    }


    @Test(expected = MempoiException.class)
    public void mempoiSheetBuilderWithoutPrepStmt() {

        Workbook wb = new XSSFWorkbook();

        MempoiSheetBuilder.aMempoiSheet()
                .withWorkbook(wb)
                .build();
    }


    @Test
    public void mempoiSheetBuilderForcingGenerationEmptyArray() {

        MempoiConfig.getInstance().setForceGeneration(true);

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMergedRegionColumns(new String[0])
                .build();

        assertNotNull("Force generation empty array - mempoi sheet not null", mempoiSheet);
        assertNull("Force generation empty array - merged regions array null", mempoiSheet.getMergedRegionColumns());
    }

    @Test
    public void mempoiSheetBuilderForcingGenerationNullArray() {

        MempoiConfig.getInstance().setForceGeneration(true);

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMergedRegionColumns(null)
                .build();

        assertNotNull("Force generation null array - mempoi sheet not null", mempoiSheet);
        assertNull("Force generation null array - merged regions array null", mempoiSheet.getMergedRegionColumns());
    }


    @Test(expected = MempoiException.class)
    public void mempoiSheetBuilderNotForcingGenerationEmptyArray() {

        // TODO eliminare quando sarà thread safe
        MempoiConfig.getInstance().setForceGeneration(false);

        MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMergedRegionColumns(new String[0])
                .build();

    }

    @Test(expected = MempoiException.class)
    public void mempoiSheetBuilderNotForcingGenerationNullEmpty() {

        MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMergedRegionColumns(null)
                .build();

    }

    @Test
    public void mempoiSheetBuilderChangedPrepStmt() {

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .build();

        assertEquals(prepStmt, mempoiSheet.getPrepStmt());

        mempoiSheet.setPrepStmt(null);

        assertNull(mempoiSheet.getPrepStmt());
    }

}
