package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiStylerBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.*;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.SummerStyleTemplate;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.util.concurrent.CompletableFuture;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.hamcrest.MatcherAssert.assertThat;

public class FooterTest extends FunctionalBaseTest {



    @Test
    public void test_with_file_and_no_footer_no_subfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_no_footer_no_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        MempoiStyler reportStyler = new MempoiStylerBuilder(workbook)
                .setStyleTemplate(new SummerStyleTemplate())
                .build();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setReportMempoiStyler(reportStyler)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void test_with_file_and_number_sum_subfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_sum_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        MempoiStyler reportStyler = new MempoiStylerBuilder(workbook)
                .setStyleTemplate(new SummerStyleTemplate())
                .build();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setMempoiSubFooter(new NumberSumSubFooter())
                    .setEvaluateCellFormulas(true)
                    .setReportMempoiStyler(reportStyler)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void test_with_file_and_number_max_subfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_max_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        MempoiStyler reportStyler = new MempoiStylerBuilder(workbook)
                .setStyleTemplate(new SummerStyleTemplate())
                .build();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setMempoiSubFooter(new NumberMaxSubFooter())
                    .setEvaluateCellFormulas(true)
                    .setReportMempoiStyler(reportStyler)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void test_with_file_and_number_min_subfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_min_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        MempoiStyler reportStyler = new MempoiStylerBuilder(workbook)
                .setStyleTemplate(new SummerStyleTemplate())
                .build();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setMempoiSubFooter(new NumberMinSubFooter())
                    .setEvaluateCellFormulas(true)
                    .setReportMempoiStyler(reportStyler)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void test_with_file_and_number_average_subfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_average_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        MempoiStyler reportStyler = new MempoiStylerBuilder(workbook)
                .setStyleTemplate(new SummerStyleTemplate())
                .build();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setMempoiSubFooter(new NumberAverageSubFooter())
                    .setReportMempoiStyler(reportStyler)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void test_with_file_and_number_custom_min_subfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_custom_min_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        MempoiStyler reportStyler = new MempoiStylerBuilder(workbook)
                .setStyleTemplate(new SummerStyleTemplate())
                .build();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setMempoiSubFooter(new NumberFormulaSubFooter() {
                        @Override
                        protected String getFormula(String colLetter, int firstDataRowIndex, int lastDataRowIndex) {
                            return "MIN(" + colLetter + firstDataRowIndex + ":" + colLetter + lastDataRowIndex + ")";
                        }
                    })
                    .setEvaluateCellFormulas(true)
                    .setReportMempoiStyler(reportStyler)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void test_with_file_and_multiple_sheet_and_number_sum_subfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_multiple_sheet_and_number_sum_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        MempoiStyler reportStyler = new MempoiStylerBuilder(workbook)
                .setStyleTemplate(new SummerStyleTemplate())
                .build();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt, "Dogs sheet"))
                    .addMempoiSheet(new MempoiSheet(conn.prepareStatement("SELECT id, creation_date, dateTime, timeStamp AS STAMPONE, name, valid, usefulChar, decimalOne, bitTwo, doublone, floattone, interao, mediano, attempato, interuccio FROM mempoi.export_test LIMIT 0, 10"), "Cats sheet"))
                    .setMempoiSubFooter(new NumberSumSubFooter())
                    .setReportMempoiStyler(reportStyler)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void test_with_file_and_standard_footer() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_standard_footer.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        MempoiStyler reportStyler = new MempoiStylerBuilder(workbook)
                .setStyleTemplate(new SummerStyleTemplate())
                .build();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setMempoiFooter(new StandardMempoiFooter(workbook, "My Poor Company"))
                    .setReportMempoiStyler(reportStyler)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void test_with_file_and_multiple_sheet_and_standard_footer() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_multiple_sheet_and_standard_footer.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        MempoiStyler reportStyler = new MempoiStylerBuilder(workbook)
                .setStyleTemplate(new SummerStyleTemplate())
                .build();

        try {

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt, "Dogs sheet"))
                    .addMempoiSheet(new MempoiSheet(conn.prepareStatement("SELECT id, creation_date, dateTime, timeStamp AS STAMPONE, name, valid, usefulChar, decimalOne, bitTwo, doublone, floattone, interao, mediano, attempato, interuccio FROM mempoi.export_test LIMIT 0, 10"), "Cats sheet"))
                    .setMempoiSubFooter(new NumberSumSubFooter())
                    .setMempoiFooter(new StandardMempoiFooter(workbook, "My Poor Company"))
                    .setReportMempoiStyler(reportStyler)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
