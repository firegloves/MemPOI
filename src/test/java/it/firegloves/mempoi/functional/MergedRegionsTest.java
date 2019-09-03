package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.CompletionException;
import java.util.concurrent.ExecutionException;

import static org.junit.Assert.assertEquals;


public class MergedRegionsTest extends FunctionalBaseMergedRegionsTest {

    /***********************************************************************
     *                               HSSF
     **********************************************************************/

    @Test
    public void testWithFileAndMergedRegionsHSSF() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_merged_regions_HSSF.xlsx");
        int limit = 450;

        try {
            prepStmt = this.createStatement(null, limit);    // TODO create tests that exceed HSSF limits and try to manage it

            // dogs sheet
            MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Merged regions name column")
                    .withPrepStmt(prepStmt)
                    .withMergedRegionColumns(new String[]{"name"})
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
//                    .withDebug(true)
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new HSSFWorkbook())
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            super.validateGeneratedFile(this.createStatement(null, limit), fut.get(), COLUMNS, HEADERS, null, new ForestStyleTemplate());

            // TODO add merged regions test code

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }
    }


    @Test
    public void testWithFileAndMergedRegionsHSSFValidateData() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_merged_regions_HSSF_validate_daa.xlsx");
        int limit = 500;

        try {
            prepStmt = this.createStatement(null, limit);    // TODO create tests that exceed HSSF limits and try to manage it

            // dogs sheet
            MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Merged regions name column")
                    .withPrepStmt(prepStmt)
                    .withMergedRegionColumns(new String[]{"name"})
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
//                    .withDebug(true)
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new HSSFWorkbook())
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            super.validateGeneratedFile(this.createStatement(null, limit), fut.get(), COLUMNS, HEADERS, null, new ForestStyleTemplate());
            super.validateMergedRegions(fut.get(), limit);

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }
    }


    @Test(expected = CompletionException.class)
    public void testWithFileAndMergedRegionsHSSF_toManyRows() throws SQLException {

        prepStmt = this.createStatement(null, 100_000);

        // dogs sheet
        MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName("Merged regions name column")
                .withPrepStmt(prepStmt)
                .withMergedRegionColumns(new String[]{"name"})
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
//                    .withDebug(true)
                .withStyleTemplate(new ForestStyleTemplate())
                .withWorkbook(new HSSFWorkbook())
                .addMempoiSheet(sheet)
                .build();

        memPOI.prepareMempoiReportToByteArray().join();
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

            // dogs sheet
            MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Merged regions name column")
                    .withPrepStmt(prepStmt)
                    .withMergedRegionColumns(new String[]{"name"})
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
//                    .withDebug(true)
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new XSSFWorkbook())
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            super.validateGeneratedFile(this.createStatement(null, limit), fut.get(), COLUMNS, HEADERS, null, new ForestStyleTemplate());
            super.validateMergedRegions(fut.get(), limit);

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }
    }


    /***********************************************************************
     *                               SXSSF
     **********************************************************************/


    @Test
    public void testWithFileAndMergedRegionsSXSSFAndFixedRowAccessWindowSize() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_merged_regions_SXSSF_fixed_row_access_windows_size.xlsx");
        int limit = 200_000;

        try {
            prepStmt = this.createStatement(null, limit);    // TODO create tests that exceed HSSF limits and try to manage it

            // dogs sheet
            MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Merged regions name column")
                    .withPrepStmt(prepStmt)
                    .withMergedRegionColumns(new String[]{"name"})
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
//                    .withDebug(true)
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new SXSSFWorkbook(200))
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            super.validateGeneratedFile(this.createStatement(null, limit), fut.get(), COLUMNS, HEADERS, null, new ForestStyleTemplate());
            super.validateMergedRegions(fut.get(), limit);

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }
    }


    @Test
    public void testWithFileAndMergedRegionsSXSSFAndVariableRowAccessWindowSize() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_merged_regions_SXSSF_variable_row_access_windows_size.xlsx");
        int limit = 200_000;

        try {
            prepStmt = this.createStatement(null, limit);    // TODO create tests that exceed HSSF limits and try to manage it

            // dogs sheet
            MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Merged regions name column")
                    .withPrepStmt(prepStmt)
                    .withMergedRegionColumns(new String[]{"name"})
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
//                    .withDebug(true)
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new SXSSFWorkbook(-1))
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            super.validateGeneratedFile(this.createStatement(null, limit), fut.get(), COLUMNS, HEADERS, null, new ForestStyleTemplate());
            super.validateMergedRegions(fut.get(), limit);

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }
    }


    @Test(expected = ExecutionException.class)
    public void testWithFileAndMergedRegionsSXSSFAndLimitedFixedRowAccessWindowSize() throws SQLException, ExecutionException, InterruptedException {

        prepStmt = this.createStatement(null, 200_000);    // TODO create tests that exceed HSSF limits and try to manage it

        // dogs sheet
        MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName("Merged regions name column")
                .withPrepStmt(prepStmt)
                .withMergedRegionColumns(new String[]{"name"})
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
//                    .withDebug(true)
                .withStyleTemplate(new ForestStyleTemplate())
                .withWorkbook(new SXSSFWorkbook(50))
                .addMempoiSheet(sheet)
                .build();

        memPOI.prepareMempoiReportToByteArray().get();
    }


    // TODO add multisheet tests
}