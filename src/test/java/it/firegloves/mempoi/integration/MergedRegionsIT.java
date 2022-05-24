package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.RoseStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.ForceGenerationUtils;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.CompletionException;
import java.util.concurrent.ExecutionException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;


public class MergedRegionsIT extends IntegrationBaseMergedRegionsIT {

    /***********************************************************************
     *                               HSSF
     **********************************************************************/

    @Test
    public void testWithFileAndMergedRegionsHSSF() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_merged_regions_HSSF.xlsx");
        int limit = 450;

        try {
            prepStmt = this
                    .createStatement(null, limit);    // TODO create tests that exceed HSSF limits and try to manage it

            // sheet
            MempoiSheet sheet = this.createMempoiSheetMergedRegions(prepStmt);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new HSSFWorkbook())
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate());

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test
    public void testWithFileAndMergedRegionsHSSFValidateData() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_HSSF_validate_data.xlsx");
        int limit = 100;

        try {
            prepStmt = this.createStatement(null, limit);

            // sheet
            MempoiSheet sheet = this.createMempoiSheetMergedRegions(prepStmt);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new HSSFWorkbook())
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            int mergedRegionsNum = this.getMergedRegionsNumber(limit, false);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate());
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum);

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test
    public void testWithFileAndMergedRegionsHSSFValidateDataMultiColumn() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_HSSF_validate_data_mulitcolumn.xlsx");
        int limit = 400;

        try {
            prepStmt = this.createStatement(null, limit);

            // sheet
            MempoiSheet sheet = this.createMempoiSheetMergedRegionsMultiColumn(prepStmt);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new HSSFWorkbook())
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            // sum of 2 columns regions
            int mergedRegionsNum = this.getMergedRegionsNumber(limit, true);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate());
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum);

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test(expected = CompletionException.class)
    public void testWithFileAndMergedRegionsHSSFToManyRows() throws SQLException {

        prepStmt = this.createStatement(null, 100_000);

        // sheet
        MempoiSheet sheet = this.createMempoiSheetMergedRegions(prepStmt);

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withStyleTemplate(new ForestStyleTemplate())
                .withWorkbook(new HSSFWorkbook())
                .addMempoiSheet(sheet)
                .build();

        memPOI.prepareMempoiReportToByteArray().join();
    }


    @Test
    public void testWithFileAndMergedRegionsHSSFValidateDataMultisheet() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_HSSF_validate_data_multisheet.xlsx");
        int limit1 = 200;
        int limit2 = 100;

        try {
            prepStmt = this.createStatement(null, limit1);
            PreparedStatement prepStmt2 = this.createStatement(null, limit2);

            // sheet 1
            MempoiSheet sheet1 = this.createMempoiSheetMergedRegions(prepStmt);

            // sheet 2
            MempoiSheet sheet2 = this.createMempoiSheet2(prepStmt2);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new HSSFWorkbook())
                    .addMempoiSheet(sheet1)
                    .addMempoiSheet(sheet2)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            int mergedRegionsNum1 = this.getMergedRegionsNumber(limit1, false);
            int mergedRegionsNum2 = this.getMergedRegionsNumber(limit2, false);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit1), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate(), 0);
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum1, 0);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit2), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new RoseStyleTemplate(), 1);
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum2, 1);


        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test
    public void testWithFileAndMergedRegionsHSSFValidateDataMultisheetMultiColumn() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_HSSF_validate_data_multisheet_mulitcolumn.xlsx");
        int limit1 = 400;
        int limit2 = 400;

        try {
            prepStmt = this.createStatement(null, limit1);
            PreparedStatement prepStmt2 = this.createStatement(null, limit2);

            // sheet 1
            MempoiSheet sheet1 = this.createMempoiSheetMergedRegionsMultiColumn(prepStmt);

            // sheet 2
            MempoiSheet sheet2 = this.createMempoiSheet2MultiColumn(prepStmt2);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new HSSFWorkbook())
                    .addMempoiSheet(sheet1)
                    .addMempoiSheet(sheet2)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            int mergedRegionsNum1 = this.getMergedRegionsNumber(limit1, true);
            int mergedRegionsNum2 = this.getMergedRegionsNumber(limit2, true);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit1), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate(), 0);
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum1, 0);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit2), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new RoseStyleTemplate(), 1);
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum2, 1);


        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    /***********************************************************************
     *                               XSSF
     **********************************************************************/


    @Test
    public void testWithFileAndMergedRegionsXSSF() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_merged_regions_XSSF.xlsx");
        int limit = 10_000;

        try {
            prepStmt = this.createStatement(null, limit);

            // sheet
            MempoiSheet sheet = this.createMempoiSheetMergedRegions(prepStmt);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new XSSFWorkbook())
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            int mergedRegionsNum = this.getMergedRegionsNumber(limit, false);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate());
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum);

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test
    public void testWithFileAndMergedRegionsXSSFValidateDataMultisheet() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_XSSF_validate_data_multisheet.xlsx");
        int limit1 = 200;
        int limit2 = 100;

        try {
            prepStmt = this.createStatement(null, limit1);
            PreparedStatement prepStmt2 = this.createStatement(null, limit2);

            // sheet 1
            MempoiSheet sheet1 = this.createMempoiSheetMergedRegions(prepStmt);

            // sheet 2
            MempoiSheet sheet2 = this.createMempoiSheet2(prepStmt2);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new XSSFWorkbook())
                    .addMempoiSheet(sheet1)
                    .addMempoiSheet(sheet2)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            int mergedRegionsNum1 = this.getMergedRegionsNumber(limit1, false);
            int mergedRegionsNum2 = this.getMergedRegionsNumber(limit2, false);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit1), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate(), 0);
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum1, 0);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit2), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new RoseStyleTemplate(), 1);
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum2, 1);


        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test
    public void testWithFileAndMergedRegionsXSSFMultiColumn() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_XSSF_multicolumn.xlsx");
        int limit = 10_000;

        try {
            prepStmt = this.createStatement(null, limit);

            // sheet
            MempoiSheet sheet = this.createMempoiSheetMergedRegionsMultiColumn(prepStmt);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new XSSFWorkbook())
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            int mergedRegionsNum = this.getMergedRegionsNumber(limit, true);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate());
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum);

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test
    public void testWithFileAndMergedRegionsXSSFValidateDataMultisheetMultiColumn() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_XSSF_validate_data_multisheet_multicolumn.xlsx");
        int limit1 = 400;
        int limit2 = 800;

        try {
            prepStmt = this.createStatement(null, limit1);
            PreparedStatement prepStmt2 = this.createStatement(null, limit2);

            // sheet 1
            MempoiSheet sheet1 = this.createMempoiSheetMergedRegionsMultiColumn(prepStmt);

            // sheet 2
            MempoiSheet sheet2 = this.createMempoiSheet2MultiColumn(prepStmt2);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new XSSFWorkbook())
                    .addMempoiSheet(sheet1)
                    .addMempoiSheet(sheet2)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            int mergedRegionsNum1 = this.getMergedRegionsNumber(limit1, true);
            int mergedRegionsNum2 = this.getMergedRegionsNumber(limit2, true);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit1), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate(), 0);
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum1, 0);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit2), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new RoseStyleTemplate(), 1);
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum2, 1);


        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    /***********************************************************************
     *                               SXSSF
     **********************************************************************/


    @Test
    public void testWithFileAndMergedRegionsSXSSFAndFixedRowAccessWindowSize() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_SXSSF_fixed_row_access_windows_size.xlsx");
        int limit = 10_000;

        try {
            prepStmt = this
                    .createStatement(null, limit);    // TODO create tests that exceed HSSF limits and try to manage it

            // sheet
            MempoiSheet sheet = this.createMempoiSheetMergedRegions(prepStmt);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new SXSSFWorkbook(200))
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            int mergedRegionsNum = this.getMergedRegionsNumber(limit, false);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate());
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum);

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test
    public void testWithFileAndMergedRegionsSXSSFAndFixedRowAccessWindowSizeMultisheet() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_SXSSF_fixed_row_access_windows_size_multisheet.xlsx");
        int limit1 = 500;
        int limit2 = 200;

        try {
            prepStmt = this
                    .createStatement(null, limit1);    // TODO create tests that exceed HSSF limits and try to manage it
            PreparedStatement prepStmt2 = this.createStatement(null, limit2);

            // sheet 1
            MempoiSheet sheet1 = this.createMempoiSheetMergedRegions(prepStmt);

            // sheet 2
            MempoiSheet sheet2 = this.createMempoiSheet2(prepStmt2);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new SXSSFWorkbook(200))
                    .addMempoiSheet(sheet1)
                    .addMempoiSheet(sheet2)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            int mergedRegionsNum1 = this.getMergedRegionsNumber(limit1, false);
            int mergedRegionsNum2 = this.getMergedRegionsNumber(limit2, false);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit1), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate());
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum1);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit2), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new RoseStyleTemplate(), 1);
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum2, 1);

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test
    public void testWithFileAndMergedRegionsSXSSFAndFixedRowAccessWindowSizeMultiColumn() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_SXSSF_fixed_row_access_windows_size_multicolumn.xlsx");
        int limit = 10_000;

        try {
            prepStmt = this
                    .createStatement(null, limit);    // TODO create tests that exceed HSSF limits and try to manage it

            // sheet
            MempoiSheet sheet = this.createMempoiSheetMergedRegionsMultiColumn(prepStmt);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new SXSSFWorkbook(200))
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            int mergedRegionsNum = this.getMergedRegionsNumber(limit, true);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate());
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum);

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test
    public void testWithFileAndMergedRegionsSXSSFAndVariableRowAccessWindowSize() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_SXSSF_variable_row_access_windows_size.xlsx");
        int limit = 10_000;

        try {
            prepStmt = this
                    .createStatement(null, limit);    // TODO create tests that exceed HSSF limits and try to manage it

            // sheet
            MempoiSheet sheet = this.createMempoiSheetMergedRegions(prepStmt);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new SXSSFWorkbook(-1))
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            int mergedRegionsNum = this.getMergedRegionsNumber(limit, false);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate());
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum);

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test(expected = ExecutionException.class)
    public void testWithFileAndMergedRegionsSXSSFAndVariableRowAccessWindowSizeMultiColumn()
            throws SQLException, ExecutionException, InterruptedException {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_SXSSF_variable_row_access_windows_size_multicolumn.xlsx");
        int limit = 200;

        prepStmt = this.createStatement(null, limit);

        // sheet
        MempoiSheet sheet = this.createMempoiSheetMergedRegionsMultiColumn(prepStmt);

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .withStyleTemplate(new ForestStyleTemplate())
                .withWorkbook(new SXSSFWorkbook(-1))
                .addMempoiSheet(sheet)
                .build();

        memPOI.prepareMempoiReportToFile().get();
    }


    @Test
    public void testWithFileAndMergedRegionsSXSSFAndVariableRowAccessWindowSizeMultisheet() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_SXSSF_variable_row_access_windows_size_multisheet.xlsx");
        int limit1 = 300;
        int limit2 = 100;

        try {
            prepStmt = this.createStatement(null, limit1);
            PreparedStatement prepStmt2 = this.createStatement(null, limit2);

            // sheet 1
            MempoiSheet sheet1 = this.createMempoiSheetMergedRegions(prepStmt);

            // sheet 2
            MempoiSheet sheet2 = this.createMempoiSheet2(prepStmt2);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new SXSSFWorkbook(-1))
                    .addMempoiSheet(sheet1)
                    .addMempoiSheet(sheet2)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            int mergedRegionsNum1 = this.getMergedRegionsNumber(limit1, false);
            int mergedRegionsNum2 = this.getMergedRegionsNumber(limit2, false);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit1), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new ForestStyleTemplate());
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum1);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit2), fut.get(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new RoseStyleTemplate(), 1);
            AssertionHelper.assertOnMergedRegions(fut.get(), mergedRegionsNum2, 1);

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test(expected = ExecutionException.class)
    public void testWithFileAndMergedRegionsSXSSFAndLimitedFixedRowAccessWindowSize()
            throws SQLException, ExecutionException, InterruptedException {

        prepStmt = this
                .createStatement(null, 50_000);    // TODO create tests that exceed HSSF limits and try to manage it

        // sheet
        MempoiSheet sheet = this.createMempoiSheetMergedRegions(prepStmt);

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withStyleTemplate(new ForestStyleTemplate())
                .withWorkbook(new SXSSFWorkbook(50))
                .addMempoiSheet(sheet)
                .build();

        memPOI.prepareMempoiReportToByteArray().get();
    }

    // SXSSF multicolumn is not supported -> multisheet and multicolumn also


    /***********************************************************************
     *                               GENERICS
     **********************************************************************/

    @Test(expected = MempoiException.class)
    public void testWithFileAndMergedRegionsHSSFNullMergedRegionsFail() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_HSSF_force_generation.xlsx");
        int limit = 450;

        prepStmt = this.createStatement(null, limit);

        MempoiSheet sheet = this.createMempoiSheetMergedRegions(prepStmt, null);

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .withStyleTemplate(new ForestStyleTemplate())
                .withWorkbook(new HSSFWorkbook())
                .addMempoiSheet(sheet)
                .build();

        memPOI.prepareMempoiReportToFile().get();
    }

    @Test(expected = MempoiException.class)
    public void testWithFileAndMergedRegionsHSSFEmptyMergedRegionsFail() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "test_with_file_and_merged_regions_HSSF_force_generation.xlsx");
        int limit = 450;

        prepStmt = this.createStatement(null, limit);

        MempoiSheet sheet = this.createMempoiSheetMergedRegions(prepStmt, new String[0]);

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .withStyleTemplate(new ForestStyleTemplate())
                .withWorkbook(new HSSFWorkbook())
                .addMempoiSheet(sheet)
                .build();

        memPOI.prepareMempoiReportToFile().get();
    }

    @Test
    public void testWithFileAndMergedRegionsHSSFNullMergedRegionsForceGeneration() {

        ForceGenerationUtils.executeTestWithForceGeneration(() -> {

            File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                    "test_with_file_and_merged_regions_HSSF_force_generation.xlsx");
            int limit = 450;

            try {
                prepStmt = this.createStatement(null, limit);

                MemPOI memPOI = MempoiBuilder.aMemPOI()
                        .withFile(fileDest)
                        .withStyleTemplate(new ForestStyleTemplate())
                        .withWorkbook(new HSSFWorkbook())
                        .addMempoiSheet(this.createMempoiSheetMergedRegions(prepStmt, null))
                        .build();

                CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
                assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());
                AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit), fut.get(), TestHelper.COLUMNS,
                        TestHelper.HEADERS, null, new ForestStyleTemplate());

            } catch (Exception e) {
                throw new MempoiException(e);
            }
        });
    }


    @Test
    public void testWithFileAndMergedRegionsHSSFEmptyMergedRegionsForceGeneration() {

        ForceGenerationUtils.executeTestWithForceGeneration(() -> {

            File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                    "test_with_file_and_merged_regions_HSSF_force_generation.xlsx");
            int limit = 450;

            try {
                prepStmt = this.createStatement(null, limit);

                MempoiSheet sheet = this.createMempoiSheetMergedRegions(prepStmt, new String[0]);

                MemPOI memPOI = MempoiBuilder.aMemPOI()
                        .withFile(fileDest)
                        .withStyleTemplate(new ForestStyleTemplate())
                        .withWorkbook(new HSSFWorkbook())
                        .addMempoiSheet(sheet)
                        .build();

                CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
                assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());
                AssertionHelper.assertOnGeneratedFile(this.createStatement(null, limit), fut.get(), TestHelper.COLUMNS,
                        TestHelper.HEADERS, null, new ForestStyleTemplate());

            } catch (Exception e) {
                fail();
            }
        });
    }

    /***********************************************************************
     *                               OTHERS
     **********************************************************************/

    /**
     * convenience method
     *
     * @param prepStmt
     * @return
     */
    private MempoiSheet createMempoiSheetMergedRegions(PreparedStatement prepStmt) {
        return this.createMempoiSheetMergedRegions(prepStmt, new String[]{"name"});
    }

    /**
     * convenience method
     *
     * @param prepStmt
     * @return
     */
    private MempoiSheet createMempoiSheetMergedRegionsMultiColumn(PreparedStatement prepStmt) {
        return this.createMempoiSheetMergedRegions(prepStmt, new String[]{"name", "usefulChar"});
    }

    /**
     * @param prepStmt
     * @return
     */
    private MempoiSheet createMempoiSheetMergedRegions(PreparedStatement prepStmt, String[] mergedColumns) {

        return MempoiSheetBuilder.aMempoiSheet()
                .withSheetName("Merged regions name column")
                .withPrepStmt(prepStmt)
                .withMergedRegionColumns(mergedColumns)
                .build();
    }


    /**
     * convenience method
     *
     * @param prepStmt
     * @return
     */
    private MempoiSheet createMempoiSheet2(PreparedStatement prepStmt) {
        return this.createMempoiSheet2(prepStmt, new String[]{"name"});
    }

    /**
     * convenience method
     *
     * @param prepStmt
     * @return
     */
    private MempoiSheet createMempoiSheet2MultiColumn(PreparedStatement prepStmt) {
        return this.createMempoiSheet2(prepStmt, new String[]{"name", "usefulChar"});
    }


    /**
     * @param prepStmt
     * @return
     */
    private MempoiSheet createMempoiSheet2(PreparedStatement prepStmt, String[] mergedColumns) {

        return MempoiSheetBuilder.aMempoiSheet()
                .withSheetName("Merged regions name column 2")
                .withPrepStmt(prepStmt)
                .withMergedRegionColumns(mergedColumns)
                .withStyleTemplate(new RoseStyleTemplate())
                .build();

    }


    /**
     * calculates and returns the
     *
     * @param limit
     * @param multiColumn
     * @return
     */
    private int getMergedRegionsNumber(int limit, boolean multiColumn) {
        return (int) Math.ceil((double) limit / 100) + (multiColumn ? (int) Math.ceil((double) limit / 80) : 0);
    }
}
