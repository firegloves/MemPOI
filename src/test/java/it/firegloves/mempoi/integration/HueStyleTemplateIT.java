package it.firegloves.mempoi.integration;

import static it.firegloves.mempoi.testutil.AssertionHelper.assertOnSimpleTextHeaderOrFooterGeneratedFile;
import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiReport;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.styles.template.AquaStyleTemplate;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.HueStyleTemplate;
import it.firegloves.mempoi.styles.template.PanegiriconStyleTemplate;
import it.firegloves.mempoi.styles.template.PurpleStyleTemplate;
import it.firegloves.mempoi.styles.template.RoseStyleTemplate;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StoneStyleTemplate;
import it.firegloves.mempoi.styles.template.SummerStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.ResultSet;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class HueStyleTemplateIT extends IntegrationBaseIT {

    @Test
    public void testWithFileAndStandardTemplate() throws Exception {
        execTest("test_with_file_and_standard_template.xlsx", StandardStyleTemplate.class);
    }


    @Test
    public void testWithFileAndSummerTemplate() throws Exception {
        execTest("test_with_file_and_summer_template.xlsx", SummerStyleTemplate.class);
    }

    @Test
    public void testWithFileAndAquaTemplate() throws Exception {
        execTest("test_with_file_and_aqua_template.xlsx", AquaStyleTemplate.class);
    }

    @Test
    public void testWithFileAndForestTemplate() throws Exception {
        execTest("test_with_file_and_forest_template.xlsx", ForestStyleTemplate.class);
    }

    @Test
    public void testWithFileAndForestTemplateOverriden() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_forest_template_overriden.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .withStyleTemplate(new ForestStyleTemplate())
                .withColsHeaderCellStyle(headerCellStyle)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .withEvaluateCellFormulas(true)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS,
                TestHelper.SUM_CELL_FORMULA, null);
        // TODO add header overriden style check

    }

    @Test
    public void testWithFileAndForestTemplateOverridenWithXSSF() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_forest_template_overriden_xssf.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();

        CellStyle commonDataCellStyle = workbook.createCellStyle();
        commonDataCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
        commonDataCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        commonDataCellStyle.setAlignment(HorizontalAlignment.CENTER);

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .withStyleTemplate(new ForestStyleTemplate())
                .withCommonDataCellStyle(commonDataCellStyle)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .withEvaluateCellFormulas(true)
                .build();

        MempoiReport mempoiReport = memPOI.prepareMempoiReport().get();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), mempoiReport.getFile());

        AssertionHelper.assertOnGeneratedFile(this.createStatement(), mempoiReport.getFile(), TestHelper.COLUMNS, TestHelper.HEADERS,
                TestHelper.SUM_CELL_FORMULA, null);

        File file = new File(mempoiReport.getFile());
        try (InputStream inp = new FileInputStream(file)) {
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);
            for (int i = 1; i < 11; i++) {
                Row row = sheet.getRow(i);
                AssertionHelper.assertOnCellStyle(row.getCell(4).getCellStyle(), commonDataCellStyle);
                AssertionHelper.assertOnCellStyle(row.getCell(5).getCellStyle(), commonDataCellStyle);
                AssertionHelper.assertOnCellStyle(row.getCell(6).getCellStyle(), commonDataCellStyle);
            }
        }
    }


    @Test
    public void testWithFileAndForestTemplateOverridenOnSimpleTextHeaderOnMempoi() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_forest_template_overriden_simple_text_header.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        String headerText = "My simple header";

        HueStyleTemplate styleTemplate = new ForestStyleTemplate();

        MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                .withSimpleHeaderText(headerText)
                .withPrepStmt(prepStmt)
                .build();

        CellStyle simpleTextHeaderCellStyle = workbook.createCellStyle();
        simpleTextHeaderCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
        simpleTextHeaderCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(sheet)
                .withStyleTemplate(styleTemplate)
                .withSimpleTextHeaderCellStyle(simpleTextHeaderCellStyle)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .withEvaluateCellFormulas(true)
                .build();

        CompletableFuture<MempoiReport> fut = memPOI.prepareMempoiReport();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get().getFile());

        AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get().getFile(), TestHelper.COLUMNS,
                TestHelper.HEADERS, "SUM(H3:H12)", styleTemplate, 0, 0, 0, true);

        assertOnSimpleTextHeaderOrFooterGeneratedFile(fut.get().getFile(), headerText, 0, 0,
                TestHelper.COLUMNS.length - 1, 0, simpleTextHeaderCellStyle);
    }

    @Test
    public void testWithFileAndForestTemplateOverridenOnSimpleTextHeaderOnMempoiSheet() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_forest_template_overriden_simple_text_header_on_sheet.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        String headerText = "My simple header";

        HueStyleTemplate styleTemplate = new ForestStyleTemplate();

        CellStyle simpleTextHeaderCellStyle = workbook.createCellStyle();
        simpleTextHeaderCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
        simpleTextHeaderCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                .withSimpleHeaderText(headerText)
                .withSimpleTextHeaderCellStyle(simpleTextHeaderCellStyle)
                .withPrepStmt(prepStmt)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(sheet)
                .withStyleTemplate(styleTemplate)
                .withSimpleTextHeaderCellStyle(simpleTextHeaderCellStyle)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .withEvaluateCellFormulas(true)
                .build();

        CompletableFuture<MempoiReport> fut = memPOI.prepareMempoiReport();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get().getFile());

        AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get().getFile(), TestHelper.COLUMNS,
                TestHelper.HEADERS, "SUM(H3:H12)", styleTemplate, 0, 0, 0, true);

        assertOnSimpleTextHeaderOrFooterGeneratedFile(fut.get().getFile(), headerText, 0, 0,
                TestHelper.COLUMNS.length - 1, 0, simpleTextHeaderCellStyle);
    }


    @Test
    public void testWithFileAndForestTemplateOverridenOnSimpleTextFooterOnMempoi() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_forest_template_overriden_simple_text_footer.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        String footerText = "My simple footer";

        HueStyleTemplate styleTemplate = new ForestStyleTemplate();

        MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                .withSimpleFooterText(footerText)
                .withPrepStmt(prepStmt)
                .build();

        CellStyle simpleTextFooterCellStyle = workbook.createCellStyle();
        simpleTextFooterCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
        simpleTextFooterCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(sheet)
                .withStyleTemplate(styleTemplate)
                .withSimpleTextFooterCellStyle(simpleTextFooterCellStyle)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .withEvaluateCellFormulas(true)
                .build();

        CompletableFuture<MempoiReport> fut = memPOI.prepareMempoiReport();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get().getFile());

        AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get().getFile(), TestHelper.COLUMNS,
                TestHelper.HEADERS, null, styleTemplate, 0, 0, 0, false);

        assertOnSimpleTextHeaderOrFooterGeneratedFile(fut.get().getFile(), footerText, 12, 0,
                TestHelper.COLUMNS.length - 1, 0, simpleTextFooterCellStyle);
    }

    @Test
    public void testWithFileAndForestTemplateOverridenOnSimpleTextFooterOnMempoiSheet() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_forest_template_overriden_simple_text_footer_on_sheet.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        String footerText = "My simple footer";

        HueStyleTemplate styleTemplate = new ForestStyleTemplate();

        CellStyle simpleTextFooterCellStyle = workbook.createCellStyle();
        simpleTextFooterCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
        simpleTextFooterCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                .withSimpleFooterText(footerText)
                .withSimpleTextFooterCellStyle(simpleTextFooterCellStyle)
                .withPrepStmt(prepStmt)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(sheet)
                .withStyleTemplate(styleTemplate)
                .withSimpleTextFooterCellStyle(simpleTextFooterCellStyle)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .withEvaluateCellFormulas(true)
                .build();

        CompletableFuture<MempoiReport> fut = memPOI.prepareMempoiReport();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get().getFile());

        AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get().getFile(), TestHelper.COLUMNS,
                TestHelper.HEADERS, null, styleTemplate, 0, 0, 0, false);

        assertOnSimpleTextHeaderOrFooterGeneratedFile(fut.get().getFile(), footerText, 12, 0,
                TestHelper.COLUMNS.length - 1, 0, simpleTextFooterCellStyle);
    }


    @Test
    public void testWithFileAndForestTemplateOverridenOnSheetstyler() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_forest_template_overriden_on_sheetstyler.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle floatingPointCellStyle = workbook.createCellStyle();
        floatingPointCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        floatingPointCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        MempoiSheet sheet = new MempoiSheet(prepStmt);
        sheet.setFloatingPointCellStyle(floatingPointCellStyle);

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(sheet)
                .withStyleTemplate(new ForestStyleTemplate())
                .withColsHeaderCellStyle(headerCellStyle)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .withEvaluateCellFormulas(true)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS,
                TestHelper.SUM_CELL_FORMULA, null);
        // TODO add header overriden style check

    }


    @Test
    public void testWithFileAndStoneTemplate() throws Exception {
        execTest("test_with_file_and_stone_template.xlsx", StoneStyleTemplate.class);
    }

    @Test
    public void testWithFileAndRoseTemplate() throws Exception {
        execTest("test_with_file_and_rose_template.xlsx", RoseStyleTemplate.class);
    }


    @Test
    public void testWithFileAndPurpleTemplate() throws Exception {
        execTest("test_with_file_and_purple_template.xlsx", PurpleStyleTemplate.class);
    }


    @Test
    public void testWithFileAndPanegiriconTemplate() throws Exception {
        execTest("test_with_file_and_panegiricon_template.xlsx", PanegiriconStyleTemplate.class);
    }

    private void execTest(String fileName, Class<? extends HueStyleTemplate> hueStyleTemplateClass) throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), fileName);
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        String headerText = "My simple header";
        final HueStyleTemplate styleTemplate = hueStyleTemplateClass.getConstructor().newInstance();

        MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                .withSimpleHeaderText(headerText)
                .withSimpleFooterText(TestHelper.SIMPLE_TEXT_FOOTER)
                .withPrepStmt(prepStmt)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(sheet)
                .withStyleTemplate(styleTemplate)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .withEvaluateCellFormulas(true)
                .build();

        final MempoiReport mempoiReport = memPOI.prepareMempoiReport().get();

        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), mempoiReport.getFile());

        AssertionHelper.assertOnGeneratedFile(this.createStatement(), mempoiReport.getFile(), TestHelper.COLUMNS,
                TestHelper.HEADERS, "SUM(H3:H12)", styleTemplate, 0, 0, 0, true);

        assertOnSimpleTextHeaderOrFooterGeneratedFile(mempoiReport.getFile(), headerText, 0, 0,
                TestHelper.COLUMNS.length - 1, 0, styleTemplate.getSimpleTextHeaderCellStyle(workbook));

        assertOnSimpleTextHeaderOrFooterGeneratedFile(mempoiReport.getFile(), TestHelper.SIMPLE_TEXT_FOOTER, 13, 0,
                TestHelper.COLUMNS.length - 1, 1, styleTemplate.getSimpleTextFooterCellStyle(workbook));
    }

    @Test
    public void testWithFileAndMultipleSheetTemplates() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_multiple_sheet_templates.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        CellStyle integerCellStyle = workbook.createCellStyle();
        integerCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        integerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        MempoiSheet catsheet = new MempoiSheet(prepStmt, "Cats");
        catsheet.setStyleTemplate(new ForestStyleTemplate());
        catsheet.setIntegerCellStyle(integerCellStyle);

        MempoiSheet dogsSheet = new MempoiSheet(conn.prepareStatement(
                super.createQuery(TestHelper.COLUMNS_2, TestHelper.HEADERS_2, TestHelper.NO_LIMITS)), "Dogs");
        dogsSheet.setStyleTemplate(new SummerStyleTemplate());

        List<MempoiSheet> sheetList = Arrays.asList(catsheet, dogsSheet);

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .withMempoiSheetList(sheetList)
//                    .withStyleTemplate(new PanegiriconStyleTemplate())     <----- it has no effects because for each sheet a template is specified
                .withMempoiSubFooter(new NumberSumSubFooter())
                .withEvaluateCellFormulas(true)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        // validates first sheet
        AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS,
                TestHelper.SUM_CELL_FORMULA, null);

        try (InputStream inp = new FileInputStream(fut.get())) {
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);

                // validates first sheet's header
            AssertionHelper.assertOnHeaderRow(sheet.getRow(0), TestHelper.HEADERS,
                    new ForestStyleTemplate().getColsHeaderCellStyle(workbook));

            // validate custom numerCellStyle
            CellStyle actual = sheet.getRow(1).getCell(0).getCellStyle();
            assertEquals(integerCellStyle.getFillForegroundColor(), actual.getFillForegroundColor());
            assertEquals(integerCellStyle.getFillPattern(), actual.getFillPattern());
        }

        // validates second sheet
        AssertionHelper.assertOnSecondPrepStmtSheet(conn.prepareStatement(
                        super.createQuery(TestHelper.COLUMNS_2, TestHelper.HEADERS_2, TestHelper.NO_LIMITS)), fut.get(), 1,
                TestHelper.HEADERS_2, true, new SummerStyleTemplate());
    }
}
