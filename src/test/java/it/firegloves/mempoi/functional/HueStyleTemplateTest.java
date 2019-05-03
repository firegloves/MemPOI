package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.styles.template.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.util.concurrent.CompletableFuture;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.hamcrest.MatcherAssert.assertThat;

public class HueStyleTemplateTest extends FunctionalBaseTest {

    @Test
    public void test_with_file_and_standard_template() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_standard_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setStyleTemplate(new StandardStyleTemplate())
                    .setMempoiSubFooter(new NumberSumSubFooter())
                    .setEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void test_with_file_and_summer_template() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_file_and_summer_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setStyleTemplate(new SummerStyleTemplate())
                    .setMempoiSubFooter(new NumberSumSubFooter())
                    .setEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void test_with_file_and_aqua_template() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_file_and_aqua_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setStyleTemplate(new AquaStyleTemplate())
                    .setMempoiSubFooter(new NumberSumSubFooter())
                    .setEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void test_with_file_and_forest_template() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_file_and_forest_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setMempoiSubFooter(new NumberSumSubFooter())
                    .setStyleTemplate(new ForestStyleTemplate())
                    .setEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void test_with_file_and_forest_template_overriden() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_file_and_forest_template_overriden.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setStyleTemplate(new ForestStyleTemplate())
                    .setHeaderCellStyle(headerCellStyle)
                    .setMempoiSubFooter(new NumberSumSubFooter())
                    .setEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void test_with_file_and_stone_template() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_stone_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setStyleTemplate(new StoneStyleTemplate())
                    .setMempoiSubFooter(new NumberSumSubFooter())
                    .setEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void test_with_file_and_rose_template() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_rose_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setStyleTemplate(new RoseStyleTemplate())
                    .setMempoiSubFooter(new NumberSumSubFooter())
                    .setEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void test_with_file_and_purple_template() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_purple_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setStyleTemplate(new PurpleStyleTemplate())
                    .setMempoiSubFooter(new NumberSumSubFooter())
                    .setEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void test_with_file_and_panegiricon_template() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_panegiricon_template.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setStyleTemplate(new PanegiriconStyleTemplate())
                    .setMempoiSubFooter(new NumberSumSubFooter())
                    .setEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
