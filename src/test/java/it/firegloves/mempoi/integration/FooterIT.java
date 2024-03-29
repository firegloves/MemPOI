package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.NumberAverageSubFooter;
import it.firegloves.mempoi.domain.footer.NumberFormulaSubFooter;
import it.firegloves.mempoi.domain.footer.NumberMaxSubFooter;
import it.firegloves.mempoi.domain.footer.NumberMinSubFooter;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.domain.footer.StandardMempoiFooter;
import it.firegloves.mempoi.styles.template.SummerStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.util.concurrent.CompletableFuture;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

public class FooterIT extends IntegrationBaseIT {


    @Test
    public void testWithFileAndNoFooterNoSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_no_footer_no_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new SummerStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }

    @Test
    public void testWithFileAndNumberSumSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_sum_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
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
            AssertionHelper.failAssertion(e);
        }
    }

    @Test
    public void testWithFileAndNumberMaxSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_max_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
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
            AssertionHelper.failAssertion(e);
        }
    }


    @Test
    public void testWithFileAndNumberMinSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_min_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
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
            AssertionHelper.failAssertion(e);
        }
    }


    @Test
    public void testWithFileAndNumberAverageSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_average_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
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
            AssertionHelper.failAssertion(e);
        }
    }


    @Test
    public void testWithFileAndNumberCustomMinSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_number_custom_min_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
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
            AssertionHelper.failAssertion(e);
        }
    }


    @Test
    public void testWithFileAndMultipleSheetAndNumberSumSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_multiple_sheet_and_number_sum_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt, "Dogs sheet"))
                    .addMempoiSheet(new MempoiSheet(conn.prepareStatement("SELECT id, creation_date, dateTime, timeStamp AS STAMPONE, name, valid, usefulChar, decimalOne, bitTwo, doublone, floattone, interao, mediano, attempato, interuccio FROM " + TestHelper.TABLE_EXPORT_TEST + " LIMIT 0, 10"), "Cats sheet"))
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withStyleTemplate(new SummerStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }


    @Test
    public void testWithFileAndStandardFooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_standard_footer.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
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
            AssertionHelper.failAssertion(e);
        }
    }


    @Test
    public void testWithFileAndMultipleSheetAndStandardFooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_multiple_sheet_and_standard_footer.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt, "Dogs sheet"))
                    .addMempoiSheet(new MempoiSheet(conn.prepareStatement("SELECT id, creation_date, dateTime, timeStamp AS STAMPONE, name, valid, usefulChar, decimalOne, bitTwo, doublone, floattone, interao, mediano, attempato, interuccio FROM " + TestHelper.TABLE_EXPORT_TEST + " LIMIT 0, 10"), "Cats sheet"))
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withMempoiFooter(new StandardMempoiFooter(workbook, "My Poor Company"))
                    .withStyleTemplate(new SummerStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }


    @Test
    public void testWithFileAndMultipleSheetWithMultipleSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_multiple_sheet_and_multiple_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MempoiSheet sheet1 = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Dogs sheet")
                    .withPrepStmt(prepStmt)
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .build();

            MempoiSheet sheet2 = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Cats sheet")
                    .withPrepStmt(super.createStatement())
                    .withMempoiSubFooter(new NumberAverageSubFooter())
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withEvaluateCellFormulas(true)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(sheet1)
                    .addMempoiSheet(sheet2)
                    .withMempoiFooter(new StandardMempoiFooter(workbook, "My Poor Company"))
                    .withStyleTemplate(new SummerStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }
}
