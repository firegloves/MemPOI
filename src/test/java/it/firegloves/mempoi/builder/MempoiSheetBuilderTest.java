package it.firegloves.mempoi.builder;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

import it.firegloves.mempoi.datapostelaboration.mempoicolumn.MempoiColumnElaborationStep;
import it.firegloves.mempoi.datapostelaboration.mempoicolumn.mergedregions.NotStreamApiMergedRegionsStep;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiColumnConfig.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.domain.datatransformation.StringDataTransformationFunction;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.domain.footer.StandardMempoiFooter;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.RoseStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.ForceGenerationUtils;
import it.firegloves.mempoi.testutil.MempoiColumnConfigTestHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.sql.PreparedStatement;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

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
                .withIntegerCellStyle(styleTemplate.getIntegerCellStyle(wb))
                .withFloatingPointCellStyle(styleTemplate.getFloatingPointCellStyle(wb))
                .withSubFooterCellStyle(styleTemplate.getSubfooterCellStyle(wb))
                .withMempoiFooter(new StandardMempoiFooter(wb, footerName))
                .withMergedRegionColumns(mergedCols)
                .withWorkbook(wb)
                .withMempoiTableBuilder(TestHelper.getTestMempoiTableBuilder(wb))
                .withMempoiPivotTableBuilder(TestHelper.getTestMempoiPivotTableBuilder(wb))
                .addMempoiColumnConfig(MempoiColumnConfigTestHelper.getTestMempoiColumnConfig())
                .build();

        assertEquals("Style template ForestTemplate", forestStyleTemplate, mempoiSheet.getStyleTemplate());
        assertEquals("Prepared Statement", prepStmt, mempoiSheet.getPrepStmt());
        assertEquals("Sheet name", sheetName, mempoiSheet.getSheetName());
        assertEquals("Subfooter", numberSumSubFooter, mempoiSheet.getMempoiSubFooter().get());
        AssertionHelper.validateCellStyle(styleTemplate.getCommonDataCellStyle(wb), mempoiSheet.getCommonDataCellStyle());
        AssertionHelper.validateCellStyle(styleTemplate.getDateCellStyle(wb), mempoiSheet.getDateCellStyle());
        AssertionHelper.validateCellStyle(styleTemplate.getDatetimeCellStyle(wb), mempoiSheet.getDatetimeCellStyle());
        AssertionHelper.validateCellStyle(styleTemplate.getHeaderCellStyle(wb), mempoiSheet.getHeaderCellStyle());
        AssertionHelper.validateCellStyle(styleTemplate.getIntegerCellStyle(wb), mempoiSheet.getIntegerCellStyle());
        AssertionHelper.validateCellStyle(styleTemplate.getFloatingPointCellStyle(wb),
                mempoiSheet.getFloatingPointCellStyle());
        AssertionHelper.validateCellStyle(styleTemplate.getSubfooterCellStyle(wb), mempoiSheet.getSubFooterCellStyle());
        assertEquals("footer text", footerName, mempoiSheet.getMempoiFooter().get().getCenterText());
        assertArrayEquals("merged cols", mergedCols, mempoiSheet.getMergedRegionColumns());
        assertEquals("workbook", wb, mempoiSheet.getWorkbook());

        AssertionHelper.validateMempoiTable(wb, mempoiSheet.getMempoiTable().get());
        AssertionHelper.validateMempoiPivotTable(wb, mempoiSheet.getMempoiPivotTable().get());

        assertEquals(1, mempoiSheet.getColumnConfigMap().size());
        AssertionHelper.validateMempoiColumnConfig(MempoiColumnConfigTestHelper.getTestMempoiColumnConfig(),
                mempoiSheet.getColumnConfigMap().get(MempoiColumnConfigTestHelper.COLUMN_NAME));
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
        AssertionHelper.validateCellStyle(forestStyleTemplate.getIntegerCellStyle(wb), mempoiSheet.getIntegerCellStyle());
        AssertionHelper.validateCellStyle(forestStyleTemplate.getFloatingPointCellStyle(wb), mempoiSheet.getFloatingPointCellStyle());
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
        AssertionHelper.validateCellStyle(forestStyleTemplate.getIntegerCellStyle(wb), mempoiSheet.getIntegerCellStyle());
        AssertionHelper.validateCellStyle(forestStyleTemplate.getFloatingPointCellStyle(wb), mempoiSheet.getFloatingPointCellStyle());
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

        ForceGenerationUtils.executeTestWithForceGeneration(() -> {

            MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                    .withPrepStmt(prepStmt)
                    .withMergedRegionColumns(new String[0])
                    .build();

            assertNotNull("Force generation empty array - mempoi sheet not null", mempoiSheet);
            assertNull("Force generation empty array - merged regions array null", mempoiSheet.getMergedRegionColumns());

        });
    }

    @Test
    public void mempoiSheetBuilderForcingGenerationNullArray() {

        ForceGenerationUtils.executeTestWithForceGeneration(() -> {

            MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                    .withPrepStmt(prepStmt)
                    .withMergedRegionColumns(null)
                    .build();

            assertNotNull("Force generation null array - mempoi sheet not null", mempoiSheet);
            assertNull("Force generation null array - merged regions array null", mempoiSheet.getMergedRegionColumns());

        });
    }


    @Test(expected = MempoiException.class)
    public void mempoiSheetBuilderNotForcingGenerationEmptyArray() {

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

    @Test
    public void mempoiSheetBuilder_withMempoiTableAndPivotTable() {

        Workbook wb = new XSSFWorkbook();
        MempoiTable mempoiTable = TestHelper.getTestMempoiTable(wb);
        MempoiPivotTable mempoiPivotTable = TestHelper.getTestMempoiPivotTable(wb);

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMempoiTable(mempoiTable)
                .withMempoiPivotTable(mempoiPivotTable)
                .build();

        assertEquals(mempoiTable, mempoiSheet.getMempoiTable().get());
        assertEquals(mempoiPivotTable, mempoiSheet.getMempoiPivotTable().get());
    }


    @Test
    public void mempoiSheetBuilderWithDataElaborationStepMap() {

        Map<String, List<MempoiColumnElaborationStep>> postElaborationStepMap = new HashMap<>();

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withDataElaborationStepMap(postElaborationStepMap)
                .build();

        assertEquals(postElaborationStepMap, mempoiSheet.getDataElaborationStepMap());
    }


    @Test
    public void mempoiSheetBuilderWithDataElaborationStep() {

        Workbook wb = new XSSFWorkbook();
        NotStreamApiMergedRegionsStep step = TestHelper.getNotStreamApiMergedRegionsStep(wb);

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withDataElaborationStep(TestHelper.MEMPOI_COLUMN_NAME, step)
                .build();

        Map<String, List<MempoiColumnElaborationStep>> actual = mempoiSheet.getDataElaborationStepMap();
        assertEquals(1, actual.keySet().size());
        assertEquals(1, actual.get(TestHelper.MEMPOI_COLUMN_NAME).size());
        assertEquals(step, actual.get(TestHelper.MEMPOI_COLUMN_NAME).get(0));
    }


    @Test(expected = MempoiException.class)
    public void withDataElaborationStepMapNullAndDataElaborationStepWillFail() {

        Workbook wb = new XSSFWorkbook();
        NotStreamApiMergedRegionsStep step = TestHelper.getNotStreamApiMergedRegionsStep(wb);

        MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withDataElaborationStepMap(null)
                .withDataElaborationStep(TestHelper.MEMPOI_COLUMN_NAME, step)
                .build();
    }


    @Test
    public void withDataElaborationStepMapNullAndDataElaborationStepAndForceGenerationShouldWork() {

        ForceGenerationUtils.executeTestWithForceGeneration(() -> {

            Workbook wb = new XSSFWorkbook();
            NotStreamApiMergedRegionsStep step = TestHelper.getNotStreamApiMergedRegionsStep(wb);

            MempoiSheetBuilder.aMempoiSheet()
                    .withPrepStmt(prepStmt)
                    .withDataElaborationStepMap(null)
                    .withDataElaborationStep(TestHelper.MEMPOI_COLUMN_NAME, step)
                    .build();
        });
    }

    @Test
    public void withMultipleMempoiColumnConfigShouldReturnTheEntireMempoiColumnConfigSet() {

        MempoiColumnConfig mempoiColumnConfig1 = MempoiColumnConfigTestHelper.getTestMempoiColumnConfig();

        String colName = "test";
        StringDataTransformationFunction<Date> dataTransformationFunction = new StringDataTransformationFunction<Date>() {
            @Override
            public Date transform(String value) throws MempoiException {
                return new Date();
            }
        };
        MempoiColumnConfig mempoiColumnConfig2 = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                .withColumnName(colName)
                .withDataTransformationFunction(dataTransformationFunction)
                .build();

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMempoiColumnConfigList(Arrays.asList(mempoiColumnConfig1, mempoiColumnConfig2))
                .build();

        AssertionHelper.validateMempoiColumnConfig(mempoiColumnConfig1,
                mempoiSheet.getColumnConfigMap().get(MempoiColumnConfigTestHelper.COLUMN_NAME));
        AssertionHelper.validateMempoiColumnConfig(mempoiColumnConfig2, mempoiSheet.getColumnConfigMap().get(colName));
    }

}
