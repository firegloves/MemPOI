/**
 * this class is meant to be run manually to analyze performances
 */

package it.firegloves.mempoi.integration;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.domain.footer.StandardMempoiFooter;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.StoneStyleTemplate;
import it.firegloves.mempoi.styles.template.SummerStyleTemplate;
import it.firegloves.mempoi.testutil.ConnectionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Ignore;
import org.junit.Test;

import java.io.File;
import java.util.concurrent.CompletableFuture;

import static org.junit.Assert.assertEquals;

@Ignore
public class SpeedIT extends IntegrationBaseIT {

    // in order to run tests you need to first run DBPopulator's main method
    // then adjust db connection string accordingly with your parameters

    @Before
    public void init() {
        try {
            this.conn = ConnectionHelper.getConnection();
            this.prepStmt = this.conn.prepareStatement("SELECT id, creation_date AS DATA_BELLISSIMA, dateTime, timeStamp, name, valid, usefulChar, decimalOne, bitTwo, doublone, floattone, interao, mediano, attempato, interuccio " +
                    "FROM " + TestHelper.TABLE_SPEED_TEST);


            if (! this.outReportFolder.exists() && ! this.outReportFolder.mkdirs()) {
                throw new MempoiException("Error in creating out report file folder: " + this.outReportFolder.getAbsolutePath() + ". Maybe permissions problem?");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void speedTestSXSSFWorkbook_1() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "speed_test_1.xlsx");

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void speedTestSXSSFWorkbook_2() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "speed_test_2.xlsx");

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
//                    .withAdjustColumnWidth(true)                     // adjusting col size on large data set will consistently slow down generation process
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void speedTestSXSSFWorkbook_3() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "speed_test_3.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withWorkbook(workbook)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void speedTestSXSSFWorkbook_4() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "speed_test_4.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withWorkbook(workbook)
                    .withStyleTemplate(new SummerStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void speedTestSXSSFWorkbook_5() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "speed_test_5.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withWorkbook(workbook)
                    .withStyleTemplate(new StoneStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withMempoiFooter(new StandardMempoiFooter(workbook, "MemPOI attack!"))
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void speedTestHSSFWorkbook_1() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_HSSFWorkbook_1.xlsx");

        try {
            this.conn = ConnectionHelper.getConnection();
            this.prepStmt = this.conn.prepareStatement("SELECT id, creation_date AS DATA_BELLISSIMA, dateTime, timeStamp, name, valid, usefulChar, decimalOne, bitTwo, doublone, floattone, interao, mediano, attempato, interuccio " +
                    "FROM "+ TestHelper.TABLE_SPEED_TEST + " LIMIT 0, 65500");
        } catch (Exception e) {
            e.printStackTrace();
        }

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withWorkbook(new HSSFWorkbook())
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void speedTestXSSFWorkbook_1() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_XSSFWorkbook_1.xlsx");

        try {
            this.conn = ConnectionHelper.getConnection();
            this.prepStmt = this.conn.prepareStatement("SELECT id, creation_date AS DATA_BELLISSIMA, dateTime, timeStamp, name, valid, usefulChar, decimalOne, bitTwo, doublone, floattone, interao, mediano, attempato, interuccio " +
                    "FROM " + TestHelper.TABLE_SPEED_TEST + " LIMIT 0, 65500");
        } catch (Exception e) {
            e.printStackTrace();
        }

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withWorkbook(new XSSFWorkbook())
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}

