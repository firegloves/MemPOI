package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.*;
import it.firegloves.mempoi.styles.template.SummerStyleTemplate;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.util.concurrent.CompletableFuture;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.junit.Assert.assertEquals;

public class FooterTest extends FunctionalBaseTest {



    @Test
    public void testWithFileAndNoFooterNoSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_no_footer_no_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new SummerStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testWithFileAndNumberSumSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_sum_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withEvaluateCellFormulas(true)
                    .withStyleTemplate(new SummerStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testWithFileAndNumberMaxSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_max_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withMempoiSubFooter(new NumberMaxSubFooter())
                    .withEvaluateCellFormulas(true)
                    .withStyleTemplate(new SummerStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void testWithFileAndNumberMinSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_min_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withMempoiSubFooter(new NumberMinSubFooter())
                    .withEvaluateCellFormulas(true)
                    .withStyleTemplate(new SummerStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void testWithFileAndNumberAverageSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_average_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withMempoiSubFooter(new NumberAverageSubFooter())
                    .withStyleTemplate(new SummerStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void testWithFileAndNumberCustomMinSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_custom_min_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withMempoiSubFooter(new NumberFormulaSubFooter() {
                        @Override
                        protected String getFormula(String colLetter, int firstDataRowIndex, int lastDataRowIndex) {
                            return "MIN(" + colLetter + firstDataRowIndex + ":" + colLetter + lastDataRowIndex + ")";
                        }
                    })
                    .withEvaluateCellFormulas(true)
                    .withStyleTemplate(new SummerStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void testWithFileAndMultipleSheetAndNumberSumSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_multiple_sheet_and_number_sum_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt, "Dogs sheet"))
                    .addMempoiSheet(new MempoiSheet(conn.prepareStatement("SELECT id, creation_date, dateTime, timeStamp AS STAMPONE, name, valid, usefulChar, decimalOne, bitTwo, doublone, floattone, interao, mediano, attempato, interuccio FROM mempoi.export_test LIMIT 0, 10"), "Cats sheet"))
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withStyleTemplate(new SummerStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void testWithFileAndStandardFooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_standard_footer.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withMempoiFooter(new StandardMempoiFooter(workbook, "My Poor Company"))
                    .withStyleTemplate(new SummerStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void testWithFileAndMultipleSheetAndStandardFooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_multiple_sheet_and_standard_footer.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt, "Dogs sheet"))
                    .addMempoiSheet(new MempoiSheet(conn.prepareStatement("SELECT id, creation_date, dateTime, timeStamp AS STAMPONE, name, valid, usefulChar, decimalOne, bitTwo, doublone, floattone, interao, mediano, attempato, interuccio FROM mempoi.export_test LIMIT 0, 10"), "Cats sheet"))
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withMempoiFooter(new StandardMempoiFooter(workbook, "My Poor Company"))
                    .withStyleTemplate(new SummerStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}