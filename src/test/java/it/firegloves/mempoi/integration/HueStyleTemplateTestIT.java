package it.firegloves.mempoi.integration;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.styles.template.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;
import org.mockito.Mock;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.SQLException;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.CompletableFuture;

import static org.junit.Assert.assertEquals;

public class HueStyleTemplateTestIT extends FunctionalBaseTestIT {

    @Mock
    StyleTemplate styleTemplate;


    @Test
    public void testWithFileAndStandardTemplate() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_standard_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
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

            super.validateGeneratedFile(this.createStatement(), fut.get(), COLUMNS, HEADERS, SUM_CELL_FORMULA, new StandardStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void testWithFileAndSummerTemplate() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_file_and_summer_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
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

            super.validateGeneratedFile(this.createStatement(), fut.get(), COLUMNS, HEADERS, SUM_CELL_FORMULA, new SummerStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testWithFileAndAquaTemplate() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_file_and_aqua_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
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

            super.validateGeneratedFile(this.createStatement(), fut.get(), COLUMNS, HEADERS, SUM_CELL_FORMULA, new AquaStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testWithFileAndForestTemplate() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_file_and_forest_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
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

            super.validateGeneratedFile(this.createStatement(), fut.get(), COLUMNS, HEADERS, SUM_CELL_FORMULA, new ForestStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testWithFileAndForestTemplateOverriden() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_file_and_forest_template_overriden.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
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

            super.validateGeneratedFile(this.createStatement(), fut.get(), COLUMNS, HEADERS, SUM_CELL_FORMULA, null);
            // TODO add header overriden style check

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void testWithFileAndForestTemplateOverridenOnSheetstyler() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_forest_template_overriden_on_sheetstyler.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle numberCellStyle = workbook.createCellStyle();
        numberCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        numberCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        MempoiSheet sheet = new MempoiSheet(prepStmt);
        sheet.setNumberCellStyle(numberCellStyle);

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
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

            super.validateGeneratedFile(this.createStatement(), fut.get(), COLUMNS, HEADERS, SUM_CELL_FORMULA, null);
            // TODO add header overriden style check

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void testWithFileAndStoneTemplate() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_stone_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
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

            super.validateGeneratedFile(this.createStatement(), fut.get(), COLUMNS, HEADERS, SUM_CELL_FORMULA, new StoneStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testWithFileAndRoseTemplate() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_rose_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
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

            super.validateGeneratedFile(this.createStatement(), fut.get(), COLUMNS, HEADERS, SUM_CELL_FORMULA, new RoseStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void testWithFileAndPurpleTemplate() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_purple_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
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

            super.validateGeneratedFile(this.createStatement(), fut.get(), COLUMNS, HEADERS, SUM_CELL_FORMULA, new PurpleStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void testWithFileAndPanegiriconTemplate() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_panegiricon_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
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

            super.validateGeneratedFile(this.createStatement(), fut.get(), COLUMNS, HEADERS, SUM_CELL_FORMULA, new PanegiriconStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testWithFileAndMultipleSheetTemplates() throws SQLException {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_multiple_sheet_templates.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        CellStyle numberCellStyle = workbook.createCellStyle();
        numberCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        numberCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        MempoiSheet catsheet = new MempoiSheet(prepStmt, "Cats");
        catsheet.setStyleTemplate(new ForestStyleTemplate());
        catsheet.setNumberCellStyle(numberCellStyle);


        MempoiSheet dogsSheet = new MempoiSheet(conn.prepareStatement(super.createQuery(COLUMNS_2, HEADERS_2, NO_LIMITS)), "Dogs");
        dogsSheet.setStyleTemplate(new SummerStyleTemplate());

        List<MempoiSheet> sheetList = Arrays.asList(catsheet, dogsSheet);

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
//                    .withDebug(true)
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
            super.validateGeneratedFile(this.createStatement(), fut.get(), COLUMNS, HEADERS, SUM_CELL_FORMULA, null);

            try (InputStream inp = new FileInputStream(fut.get())) {
                Workbook wb = WorkbookFactory.create(inp);
                Sheet sheet = wb.getSheetAt(0);

                // validates first sheet's header
                super.validateHeaderRow(sheet.getRow(0), HEADERS, new ForestStyleTemplate().getHeaderCellStyle(workbook));

                // validate custom numerCellStyle
                CellStyle actual = sheet.getRow(1).getCell(7).getCellStyle();
                assertEquals(numberCellStyle.getFillForegroundColor(), actual.getFillForegroundColor());
                assertEquals(numberCellStyle.getFillPattern(), actual.getFillPattern());
            }

            // validates second sheet
            super.validateSecondPrepStmtSheet(conn.prepareStatement(super.createQuery(COLUMNS_2, HEADERS_2, NO_LIMITS)), fut.get(), 1, HEADERS_2, true, new SummerStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
