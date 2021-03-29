package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.styles.template.AquaStyleTemplate;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.PanegiriconStyleTemplate;
import it.firegloves.mempoi.styles.template.PurpleStyleTemplate;
import it.firegloves.mempoi.styles.template.RoseStyleTemplate;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StoneStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.styles.template.SummerStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;
import org.mockito.Mock;

public class HueStyleTemplateIT extends IntegrationBaseIT {

    @Mock
    StyleTemplate styleTemplate;


    @Test
    public void testWithFileAndStandardTemplate() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_standard_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new StandardStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS, TestHelper.SUM_CELL_FORMULA, new StandardStyleTemplate());

    }


    @Test
    public void testWithFileAndSummerTemplate() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_file_and_summer_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new SummerStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS, TestHelper.SUM_CELL_FORMULA, new SummerStyleTemplate());
    }

    @Test
    public void testWithFileAndAquaTemplate() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_file_and_aqua_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

    MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new AquaStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS, TestHelper.SUM_CELL_FORMULA, new AquaStyleTemplate());

    }

    @Test
    public void testWithFileAndForestTemplate() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_file_and_forest_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS, TestHelper.SUM_CELL_FORMULA, new ForestStyleTemplate());

    }

    @Test
    public void testWithFileAndForestTemplateOverriden() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_file_and_forest_template_overriden.xlsx");
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
                    .withHeaderCellStyle(headerCellStyle)
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS, TestHelper.SUM_CELL_FORMULA, null);
            // TODO add header overriden style check

    }


    @Test
    public void testWithFileAndForestTemplateOverridenOnSheetstyler() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_forest_template_overriden_on_sheetstyler.xlsx");
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
                    .withHeaderCellStyle(headerCellStyle)
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS, TestHelper.SUM_CELL_FORMULA, null);
            // TODO add header overriden style check

    }


    @Test
    public void testWithFileAndStoneTemplate() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_stone_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new StoneStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS, TestHelper.SUM_CELL_FORMULA, new StoneStyleTemplate());
    }

    @Test
    public void testWithFileAndRoseTemplate() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_rose_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .withStyleTemplate(new RoseStyleTemplate())
                .withMempoiSubFooter(new NumberSumSubFooter())
                .withEvaluateCellFormulas(true)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS, TestHelper.SUM_CELL_FORMULA, new RoseStyleTemplate());
    }


    @Test
    public void testWithFileAndPurpleTemplate() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_purple_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new PurpleStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS, TestHelper.SUM_CELL_FORMULA, new PurpleStyleTemplate());
    }


    @Test
    public void testWithFileAndPanegiriconTemplate() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_panegiricon_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new PanegiriconStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS, TestHelper.SUM_CELL_FORMULA, new PanegiriconStyleTemplate());
    }

    @Test
    public void testWithFileAndMultipleSheetTemplates() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_multiple_sheet_templates.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        CellStyle integerCellStyle = workbook.createCellStyle();
        integerCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        integerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        MempoiSheet catsheet = new MempoiSheet(prepStmt, "Cats");
        catsheet.setStyleTemplate(new ForestStyleTemplate());
        catsheet.setIntegerCellStyle(integerCellStyle);


        MempoiSheet dogsSheet = new MempoiSheet(conn.prepareStatement(super.createQuery(TestHelper.COLUMNS_2, TestHelper.HEADERS_2, TestHelper.NO_LIMITS)), "Dogs");
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
            AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS, TestHelper.SUM_CELL_FORMULA, null);

            try (InputStream inp = new FileInputStream(fut.get())) {
                Workbook wb = WorkbookFactory.create(inp);
                Sheet sheet = wb.getSheetAt(0);

                // validates first sheet's header
                AssertionHelper.assertOnHeaderRow(sheet.getRow(0), TestHelper.HEADERS, new ForestStyleTemplate().getHeaderCellStyle(workbook));

                // validate custom numerCellStyle
                CellStyle actual = sheet.getRow(1).getCell(0).getCellStyle();
                assertEquals(integerCellStyle.getFillForegroundColor(), actual.getFillForegroundColor());
                assertEquals(integerCellStyle.getFillPattern(), actual.getFillPattern());
            }

            // validates second sheet
            AssertionHelper.assertOnSecondPrepStmtSheet(conn.prepareStatement(super.createQuery(TestHelper.COLUMNS_2, TestHelper.HEADERS_2, TestHelper.NO_LIMITS)), fut.get(), 1, TestHelper.HEADERS_2, true, new SummerStyleTemplate());
    }
}
