package it.firegloves.mempoi.integration;

import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnCols;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnHeaders;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnRows;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnSubfooter;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

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

    private final int colOffset = 3;
    private final int rowOffset = 6;

    @Test
    public void shouldApplyTheColumnOffsetIfSupplied() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "offset_col.xlsx");

        final MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withColumnsOffset(colOffset)
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

            assertOnColOffsetSheet(fut.get(), 0);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void shouldApplyTheRowOffsetIfSupplied() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "offset_row.xlsx");

        final MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withRowsOffset(rowOffset)
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

            assertOnRowOffsetSheet(fut.get(), 0);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void shouldApplyColumnAndRowOffsetsIfSupplied() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "offset_row_and_col.xlsx");

        final MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
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
                    null, new StandardStyleTemplate(), 0, rowOffset, colOffset);

            MempoiSheetMetadata mempoiSheetMetadata = fut.get().getMempoiSheetMetadata(0);
            assertNull(mempoiSheetMetadata.getSheetName());
            assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
            assertOnHeaders(mempoiSheetMetadata, rowOffset, "D7:K7");
            assertOnRows(mempoiSheetMetadata, 11, 10, 7, 16, "D8:K17", rowOffset);
            assertOnSubfooter(mempoiSheetMetadata, 17, "D19:K19");
            assertOnCols(mempoiSheetMetadata, 8, 3, 10, colOffset);

            // no assert on table because its area is manually assigned by the user
            // no assert on pivot table because its area is manually assigned by the user

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void shouldApplyColumnAndRowOffsetsIfSuppliedWithMultipleSheet() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "offset_row_and_col_multiple_sheet.xlsx");

        final MempoiSheet mempoiSheet1 = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withColumnsOffset(colOffset)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .build();

        final MempoiSheet mempoiSheet2 = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
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

            assertOnColOffsetSheet(fut.get(), 0);
            assertOnRowOffsetSheet(fut.get(), 1);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    private void assertOnColOffsetSheet(MempoiReport report, int sheetIndex) throws SQLException {

        AssertionHelper.assertOnGeneratedFile(this.createStatement(), report.getFile(), TestHelper.COLUMNS,
                TestHelper.HEADERS,
                null, new StandardStyleTemplate(), sheetIndex, 0, colOffset);

        MempoiSheetMetadata mempoiSheetMetadata = report.getMempoiSheetMetadata(sheetIndex);
        assertNull(mempoiSheetMetadata.getSheetName());
        assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
        assertOnHeaders(mempoiSheetMetadata, 0, "D1:K1");
        assertOnRows(mempoiSheetMetadata, 11, 10, 1, 10, "D2:K11");
        assertOnSubfooter(mempoiSheetMetadata, 11, "D12:K12");
        assertOnCols(mempoiSheetMetadata, 8, 3, 10, colOffset);

        // no assert on table because its area is manually assigned by the user
        // no assert on pivot table because its area is manually assigned by the user
    }

    private void assertOnRowOffsetSheet(MempoiReport report, int sheetIndex) throws SQLException {
        AssertionHelper.assertOnGeneratedFile(this.createStatement(), report.getFile(), TestHelper.COLUMNS, TestHelper.HEADERS,
                null, new StandardStyleTemplate(), sheetIndex, rowOffset, 0);

        MempoiSheetMetadata mempoiSheetMetadata = report.getMempoiSheetMetadata(sheetIndex);
        assertNull(mempoiSheetMetadata.getSheetName());
        assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
        assertOnHeaders(mempoiSheetMetadata, 6, "A7:H7");
        assertOnRows(mempoiSheetMetadata, 11, 10, 7, 16, "A8:H17", rowOffset);
        assertOnSubfooter(mempoiSheetMetadata, 17, "A19:H19");
        assertOnCols(mempoiSheetMetadata, 8, 3, 10);

        // no assert on table because its area is manually assigned by the user
        // no assert on pivot table because its area is manually assigned by the user
    }
}
