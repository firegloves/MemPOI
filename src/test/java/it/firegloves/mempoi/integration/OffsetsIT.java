package it.firegloves.mempoi.integration;

import static it.firegloves.mempoi.testutil.AssertionHelper.assertOnSimpleTextHeader;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnCols;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnHeaders;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnRows;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnSubfooter;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.fail;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiReport;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiSheetMetadata;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.sql.SQLException;
import java.util.concurrent.CompletableFuture;
import org.apache.poi.ss.SpreadsheetVersion;
import org.junit.Test;

public class OffsetsIT extends IntegrationBaseIT {

    private final String fileFolder = "/offset";

    private final int colOffset = 3;
    private final int rowOffset = 6;

    @Test
    public void shouldApplyTheColumnOffsetIfSupplied() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath() + fileFolder, "offset_col.xlsx");

        final MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withColumnsOffset(colOffset)
                .withSimpleHeaderText(TestHelper.SIMPLE_TEXT_HEADER)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(mempoiSheet)
                .build();

        try {
            final CompletableFuture<MempoiReport> fut = memPOI.prepareMempoiReport();
            final String file = fut.get().getFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), file);
            MempoiReport report = fut.get();

            AssertionHelper.assertOnGeneratedFile(this.createStatement(), report.getFile(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new StandardStyleTemplate(), 0, 0,
                    colOffset, true);

            assertOnSimpleTextHeader(file, 0, 0, colOffset, 10, TestHelper.SIMPLE_TEXT_HEADER);

            MempoiSheetMetadata mempoiSheetMetadata = report.getMempoiSheetMetadata(0);
            assertNull(mempoiSheetMetadata.getSheetName());
            assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
            assertOnHeaders(mempoiSheetMetadata, 1, "D2:K2");
            assertOnRows(mempoiSheetMetadata, 13, 10, 2, 11, "D3:K12");
            assertOnSubfooter(mempoiSheetMetadata, 12, "D13:K13");
            assertOnCols(mempoiSheetMetadata, 8, 3, 10, colOffset);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void shouldApplyTheRowOffsetIfSupplied() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath() + fileFolder, "offset_row.xlsx");

        final MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withRowsOffset(rowOffset)
                .withSimpleHeaderText(TestHelper.SIMPLE_TEXT_HEADER)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(mempoiSheet)
                .build();

        try {
            final CompletableFuture<MempoiReport> fut = memPOI.prepareMempoiReport();
            final String file = fut.get().getFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), file);

            assertOnRowOffsetSheet(fut.get(), 0, false);

            assertOnSimpleTextHeader(file, 0, rowOffset, 0, 7, TestHelper.SIMPLE_TEXT_HEADER);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void shouldApplyColumnAndRowOffsetsIfSupplied() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath() + fileFolder, "offset_row_and_col.xlsx");

        final MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withSimpleHeaderText(TestHelper.SIMPLE_TEXT_HEADER)
                .withRowsOffset(rowOffset)
                .withColumnsOffset(colOffset)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(mempoiSheet)
                .build();

        try {
            final CompletableFuture<MempoiReport> fut = memPOI.prepareMempoiReport();
            final String file = fut.get().getFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), file);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(), file, TestHelper.COLUMNS, TestHelper.HEADERS,
                    null, new StandardStyleTemplate(), 0, rowOffset+1, colOffset, false);

            MempoiSheetMetadata mempoiSheetMetadata = fut.get().getMempoiSheetMetadata(0);
            assertNull(mempoiSheetMetadata.getSheetName());
            assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
            assertOnHeaders(mempoiSheetMetadata, rowOffset + 1, "D8:K8");
            assertOnRows(mempoiSheetMetadata, 12, 10, 8, 17, "D9:K18", rowOffset);
            assertOnSubfooter(mempoiSheetMetadata, 18, "D20:K20");
            assertOnCols(mempoiSheetMetadata, 8, 3, 10, colOffset);

            assertOnSimpleTextHeader(file, 0, rowOffset, colOffset, 10, TestHelper.SIMPLE_TEXT_HEADER);

            // no assert on table because its area is manually assigned by the user
            // no assert on pivot table because its area is manually assigned by the user

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void shouldApplyColumnAndRowOffsetsIfSuppliedWithMultipleSheet() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath() + fileFolder, "offset_row_and_col_multiple_sheet.xlsx");

        final MempoiSheet mempoiSheet1 = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withSimpleHeaderText(TestHelper.SIMPLE_TEXT_HEADER)
                .withColumnsOffset(colOffset)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .build();

        final MempoiSheet mempoiSheet2 = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(super.createStatement())
                .withRowsOffset(rowOffset)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(mempoiSheet1)
                .addMempoiSheet(mempoiSheet2)
                .build();

        try {
            final CompletableFuture<MempoiReport> fut = memPOI.prepareMempoiReport();
            final String file = fut.get().getFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), file);

            assertOnColOffsetSheet(fut.get(), 0, 13);

            AssertionHelper.assertOnGeneratedFile(this.createStatement(), fut.get().getFile(), TestHelper.COLUMNS,
                    TestHelper.HEADERS, null, new StandardStyleTemplate(), 1, rowOffset,
                    0, false);

            MempoiSheetMetadata mempoiSheetMetadata = fut.get().getMempoiSheetMetadata(1);
            assertNull(mempoiSheetMetadata.getSheetName());
            assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
            assertOnHeaders(mempoiSheetMetadata, 6, "A7:H7");
            assertOnRows(mempoiSheetMetadata, 11, 10, 7, 16, "A8:H17", rowOffset);
            assertOnCols(mempoiSheetMetadata, 8, 0, 7);

            assertOnSimpleTextHeader(file, 0, 0, colOffset, 10, TestHelper.SIMPLE_TEXT_HEADER);

        } catch (Exception e) {
            e.printStackTrace();
            fail();
        }
    }


    private void assertOnColOffsetSheet(MempoiReport report, int sheetIndex, int totalRows) throws SQLException {

        AssertionHelper.assertOnGeneratedFile(this.createStatement(), report.getFile(), TestHelper.COLUMNS,
                TestHelper.HEADERS, null, new StandardStyleTemplate(), sheetIndex, 0,
                colOffset, true);

        MempoiSheetMetadata mempoiSheetMetadata = report.getMempoiSheetMetadata(sheetIndex);
        assertNull(mempoiSheetMetadata.getSheetName());
        assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
        assertOnHeaders(mempoiSheetMetadata, 1, "D2:K2");
        assertOnRows(mempoiSheetMetadata, totalRows, 10, 2, 11, "D3:K12");
        assertOnSubfooter(mempoiSheetMetadata, 12, "D13:K13");
        assertOnCols(mempoiSheetMetadata, 8, 3, 10, colOffset);

        // no assert on table because its area is manually assigned by the user
        // no assert on pivot table because its area is manually assigned by the user
    }

    private void assertOnRowOffsetSheet(MempoiReport report, int sheetIndex, boolean subfooterPresent) throws SQLException {
        AssertionHelper.assertOnGeneratedFile(this.createStatement(), report.getFile(), TestHelper.COLUMNS, TestHelper.HEADERS,
                null, new StandardStyleTemplate(), sheetIndex, rowOffset, 0, true);

        MempoiSheetMetadata mempoiSheetMetadata = report.getMempoiSheetMetadata(sheetIndex);
        assertNull(mempoiSheetMetadata.getSheetName());
        assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
        assertOnHeaders(mempoiSheetMetadata, 7, "A8:H8");
        assertOnRows(mempoiSheetMetadata, 12, 10, 8, 17, "A9:H18", rowOffset);
        assertOnSubfooter(mempoiSheetMetadata, subfooterPresent ? 18 : null, "A20:H20");
        assertOnCols(mempoiSheetMetadata, 8, 0, 7);

        // no assert on table because its area is manually assigned by the user
        // no assert on pivot table because its area is manually assigned by the user
    }
}
