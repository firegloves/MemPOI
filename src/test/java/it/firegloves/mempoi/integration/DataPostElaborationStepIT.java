package it.firegloves.mempoi.integration;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.datapostelaboration.mempoicolumn.MempoiColumnElaborationStep;
import it.firegloves.mempoi.datapostelaboration.mempoicolumn.mergedregions.NotStreamApiMergedRegionsStep;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.PreparedStatement;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.CompletableFuture;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;


public class DataPostElaborationStepIT extends IntegrationBaseMergedRegionsIT {

    /***********************************************************************
     *                               HSSF
     **********************************************************************/

    @Test
    public void testWithFileAndGenericDataPostElaborationHSSF() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_generic_data_post_elaboration_HSSF.xlsx");
        int limit = 450;

        try {
            prepStmt = this.createStatement(null, limit);
            Workbook wb = new HSSFWorkbook();

            // sheet
            MempoiSheet sheet = this.createMempoiSheet(prepStmt, wb).build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(wb)
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper.validateGeneratedFile(this.createStatement(null, limit), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS, null, new ForestStyleTemplate());

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test
    public void testWithFileAndMultipleGenericDataPostElaborationHSSF() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_multiple_generic_data_post_elaboration_HSSF.xlsx");
        int limit = 450;

        try {
            prepStmt = this.createStatement(null, limit);
            Workbook wb = new HSSFWorkbook();

            // sheet
            MempoiSheet sheet = this.createMempoiSheetMultipleStep(prepStmt, wb).build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(wb)
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            this.validateMultipleStepFile(fut.get());

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }

    @Test
    public void testWithMultipleGenericAndMergedRegionsHSSF() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_multiple_generic_and_merged_HSSF.xlsx");
        int limit = 450;

        try {
            prepStmt = this.createStatement(null, limit);
            Workbook wb = new HSSFWorkbook();

            StyleTemplate template = new ForestStyleTemplate();

            // sheet
            MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Multiple steps")
                    .withPrepStmt(prepStmt)
                    .withDataElaborationStep("name", new DummyDataPostElaborationStep(4))
                    .withDataElaborationStep("name", new NotStreamApiMergedRegionsStep<>(template.getCommonDataCellStyle(wb), 4))
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(wb)
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            this.validateMultipleStepFile(fut.get());

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    /***********************************************************************
     *                               OTHERS
     **********************************************************************/

    /**
     * dummy MempoiColumnElaborationStep for testing purpose
     */
    private class DummyDataPostElaborationStep implements MempoiColumnElaborationStep<String> {

        int colIndex;

        public DummyDataPostElaborationStep(int colIndex) {
            this.colIndex = colIndex;
        }

        @Override
        public void performAnalysis(Cell cell, String value) {

        }

        @Override
        public void closeAnalysis(int currRowNum) {

        }

        @Override
        public void execute(MempoiSheet mempoiSheet, Workbook workbook) {

            Sheet sheet = workbook.getSheet(mempoiSheet.getSheetName());
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();

            for (int i = 1; i < sheet.getLastRowNum(); i++) {

                Cell cell = sheet.getRow(i).getCell(colIndex);
                if (this.isMerged(mergedRegions, cell).isPresent() && cell.getStringCellValue().equals("hi dear")) {
                    cell.setCellValue("changed region!");
                } else {
                    cell.setCellValue("changed cell!");
                }

            }
        }

        private Optional<CellRangeAddress> isMerged(List<CellRangeAddress> mergedRegions, Cell cell) {
            return mergedRegions.stream()
                    .filter(cellAddresses -> cellAddresses.isInRange(cell))
                    .findFirst();
        }
    }


    /**
     * @param prepStmt
     * @return
     */
    private MempoiSheetBuilder createMempoiSheet(PreparedStatement prepStmt, Workbook wb) {

        StyleTemplate template = new ForestStyleTemplate();

        return MempoiSheetBuilder.aMempoiSheet()
                .withSheetName("Merged regions name column")
                .withPrepStmt(prepStmt)
                .withDataElaborationStep("name", new NotStreamApiMergedRegionsStep<>(template.getCommonDataCellStyle(wb), 4))
                .withDataElaborationStep("usefulChar", new NotStreamApiMergedRegionsStep<>(template.getCommonDataCellStyle(wb), 6));
    }


    /**
     * @param prepStmt
     * @return
     */
    private MempoiSheetBuilder createMempoiSheetMultipleStep(PreparedStatement prepStmt, Workbook wb) {

        return this.createMempoiSheet(prepStmt, wb)
                .withDataElaborationStep("name", new DummyDataPostElaborationStep(4));
    }


    /**
     * validates columns of multistep generated file
     *
     * @param fileToValidate
     */
    private void validateMultipleStepFile(String fileToValidate) throws Exception {

        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {

            Workbook wb = WorkbookFactory.create(inp);

            Sheet sheet = wb.getSheetAt(0);

            for (int i = 1; i < sheet.getLastRowNum(); i++) {
                String cellValue = sheet.getRow(i).getCell(4).getStringCellValue();
                assertTrue(cellValue.equals("hello dog") || cellValue.equals("changed region!") || cellValue.equals("changed cell!"));
            }
        }
    }
}