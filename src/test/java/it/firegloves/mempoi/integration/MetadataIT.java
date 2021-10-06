package it.firegloves.mempoi.integration;

import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertNullPivotTable;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertNullTable;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnCols;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnHeaders;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnPivotTable;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnRows;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnSubfooter;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnTable;
import static it.firegloves.mempoi.testutil.MetadataAssertionHelper.assertOnUsedWorkbookConfig;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiPivotTableBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.builder.MempoiTableBuilder;
import it.firegloves.mempoi.domain.MempoiReport;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiSheetMetadata;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.util.Arrays;
import java.util.Collections;
import java.util.concurrent.ExecutionException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class MetadataIT extends IntegrationBaseIT {

    private final String fileFolder = "/metadata";
    private final String dogsQuery = "SELECT pet_name AS DOG_NAME, pet_race AS DOG_RACE FROM pets WHERE pet_type = 'dog'";
    private final String sheetName = "Dogs sheet";

    @Test
    public void shouldReturnCorrectMetadataExportedOnFile() throws ExecutionException, InterruptedException {

        File fileDest = new File(this.outReportFolder.getAbsolutePath() + fileFolder, "test_with_file.xls");
        final MempoiReport mempoiReport = MempoiBuilder.aMemPOI()
                .withWorkbook(new HSSFWorkbook())
                .withFile(fileDest)
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .build()
                .prepareMempoiReport()
                .get();

        assertEquals(this.outReportFolder.getAbsolutePath() + fileFolder + "/test_with_file.xls", mempoiReport.getFile());
        assertNull(mempoiReport.getBytes());
    }

    @Test
    public void shouldReturnCorrectMetadataWithABaseReport() throws Exception {

        // dogs sheet
        MempoiSheet dogsSheet = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(sheetName)
                .withPrepStmt(conn.prepareStatement(dogsQuery))
                .build();

        final MempoiReport mempoiReport = MempoiBuilder.aMemPOI()
                .withAdjustColumnWidth(true)
                .addMempoiSheet(dogsSheet)
                .build()
                .prepareMempoiReport()
                .get();

        assertNull(mempoiReport.getFile());
        assertNotNull(mempoiReport.getBytes());

        MempoiSheetMetadata mempoiSheetMetadata = mempoiReport.getMempoiSheetMetadata(0);
        assertEquals(sheetName, mempoiSheetMetadata.getSheetName());
        assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
        assertOnHeaders(mempoiSheetMetadata, 0, "A1:B1");
        assertOnRows(mempoiSheetMetadata, 4, 3, 1, 3, "A2:B4");
        assertOnSubfooter(mempoiSheetMetadata, null, null);
        assertOnCols(mempoiSheetMetadata, 2, 0, 1);
        assertNullTable(mempoiSheetMetadata);
        assertNullPivotTable(mempoiSheetMetadata);
        assertOnUsedWorkbookConfig(mempoiReport.getUsedWorkbookConfig(),
                mempoiReport.getUsedWorkbookConfig().getWorkbook(),
                true, false, false,
                Collections.singletonList(dogsSheet), false);
    }

    @Test
    public void shouldReturnCorrectMetadataWithMultipleSheets() throws Exception {

        // dogs sheet
        MempoiSheet dogsSheet = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(sheetName)
                .withPrepStmt(conn.prepareStatement(dogsQuery))
                .build();

        MempoiSheet anotherSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .build();

        final MempoiReport mempoiReport = MempoiBuilder.aMemPOI()
                .withAdjustColumnWidth(true)
                .addMempoiSheet(dogsSheet)
                .addMempoiSheet(anotherSheet)
                .build()
                .prepareMempoiReport()
                .get();

        assertNull(mempoiReport.getFile());
        assertNotNull(mempoiReport.getBytes());

        MempoiSheetMetadata dogMempoiSheetMetadata = mempoiReport.getMempoiSheetMetadata(0);
        assertEquals(sheetName, dogMempoiSheetMetadata.getSheetName());
        assertEquals(SpreadsheetVersion.EXCEL2007, dogMempoiSheetMetadata.getSpreadsheetVersion());
        assertOnHeaders(dogMempoiSheetMetadata, 0, "A1:B1");
        assertOnRows(dogMempoiSheetMetadata, 4, 3, 1, 3, "A2:B4");
        assertOnSubfooter(dogMempoiSheetMetadata, null, null);
        assertOnCols(dogMempoiSheetMetadata, 2, 0, 1);
        assertNullTable(dogMempoiSheetMetadata);
        assertNullPivotTable(dogMempoiSheetMetadata);

        MempoiSheetMetadata mempoiSheetMetadata = mempoiReport.getMempoiSheetMetadata(1);
        assertNull(mempoiSheetMetadata.getSheetName());
        assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
        assertOnHeaders(mempoiSheetMetadata, 0, "A1:H1");
        assertOnRows(mempoiSheetMetadata, 11, 10, 1, 10, "A2:H11");
        assertOnSubfooter(mempoiSheetMetadata, 11, "A12:H12");
        assertOnCols(mempoiSheetMetadata, 8, 0, 7);
        assertNullTable(mempoiSheetMetadata);
        assertNullPivotTable(mempoiSheetMetadata);

        assertOnUsedWorkbookConfig(mempoiReport.getUsedWorkbookConfig(),
                mempoiReport.getUsedWorkbookConfig().getWorkbook(),
                true, false, false, Arrays.asList(dogsSheet, anotherSheet),
                false);
    }

    @Test
    public void shouldReturnCorrectMetadataWithFooterAndSubfooter() throws Exception {

        final MempoiSheet mempoiSheet = new MempoiSheet(prepStmt);

        final MempoiReport mempoiReport = MempoiBuilder.aMemPOI()
                .withAdjustColumnWidth(false)
                .addMempoiSheet(mempoiSheet)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .withEvaluateCellFormulas(true)
                .build()
                .prepareMempoiReport()
                .get();

        assertNull(mempoiReport.getFile());

        MempoiSheetMetadata mempoiSheetMetadata = mempoiReport.getMempoiSheetMetadata(0);
        assertNull(mempoiSheetMetadata.getSheetName());
        assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
        assertOnHeaders(mempoiSheetMetadata, 0, "A1:H1");
        assertOnRows(mempoiSheetMetadata, 11, 10, 1, 10, "A2:H11");
        assertOnSubfooter(mempoiSheetMetadata, 11, "A12:H12");
        assertOnCols(mempoiSheetMetadata, 8, 0, 7);
        assertNullTable(mempoiSheetMetadata);
        assertNullPivotTable(mempoiSheetMetadata);

        assertOnUsedWorkbookConfig(mempoiReport.getUsedWorkbookConfig(),
                mempoiReport.getUsedWorkbookConfig().getWorkbook(),
                false, true, true,
                Collections.singletonList(mempoiSheet), false);
    }

    @Test
    public void shouldReturnCorrectMetadataWithTable() throws Exception {

        XSSFWorkbook workbook = new XSSFWorkbook();

        MempoiTableBuilder mempoiTableBuilder = TestHelper.getTestMempoiTableBuilder(workbook)
                .withAreaReferenceSource("A1:H11");

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMempoiTableBuilder(mempoiTableBuilder)
                .build();

        final MempoiReport mempoiReport = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(mempoiSheet)
                .build()
                .prepareMempoiReport()
                .get();

        assertNull(mempoiReport.getFile());

        MempoiSheetMetadata mempoiSheetMetadata = mempoiReport.getMempoiSheetMetadata(0);
        assertNull(mempoiSheetMetadata.getSheetName());
        assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
        assertOnHeaders(mempoiSheetMetadata, 0, "A1:H1");
        assertOnRows(mempoiSheetMetadata, 11, 10, 1, 10, "A2:H11");
        assertOnCols(mempoiSheetMetadata, 8, 0, 7);
        assertOnSubfooter(mempoiSheetMetadata, null, null);
        assertOnTable(mempoiSheetMetadata, 0, 10, 0, 7, "A1:H11");
        assertNullPivotTable(mempoiSheetMetadata);

        assertOnUsedWorkbookConfig(mempoiReport.getUsedWorkbookConfig(),
                mempoiReport.getUsedWorkbookConfig().getWorkbook(),
                true, false, false,
                Collections.singletonList(mempoiSheet), false);
    }

    @Test
    public void shouldReturnCorrectMetadataWithPivotTable() throws Exception {

        XSSFWorkbook workbook = new XSSFWorkbook();

        final MempoiPivotTableBuilder mempoiPivotTableBuilder = TestHelper
                .getTestMempoiPivotTableBuilder(workbook, null)
                .withAreaReferenceSource("A1:H11")
                .withPosition(new CellReference("J3"));

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMempoiPivotTableBuilder(mempoiPivotTableBuilder)
                .build();

        final MempoiReport mempoiReport = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(mempoiSheet)
                .build()
                .prepareMempoiReport()
                .get();

        assertNull(mempoiReport.getFile());

        MempoiSheetMetadata mempoiSheetMetadata = mempoiReport.getMempoiSheetMetadata(0);
        assertNull(mempoiSheetMetadata.getSheetName());
        assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
        assertOnHeaders(mempoiSheetMetadata, 0, "A1:H1");
        assertOnRows(mempoiSheetMetadata, 11, 10, 1, 10, "A2:H11");
        assertOnCols(mempoiSheetMetadata, 8, 0, 7);
        assertOnSubfooter(mempoiSheetMetadata, null, null);
        assertNullTable(mempoiSheetMetadata);
        assertOnPivotTable(mempoiSheetMetadata, 2, 9, 0, 0, 10, 7, "J3", "A1:H11");

        assertOnUsedWorkbookConfig(mempoiReport.getUsedWorkbookConfig(),
                mempoiReport.getUsedWorkbookConfig().getWorkbook(),
                true, false, false,
                Collections.singletonList(mempoiSheet), false);
    }

    @Test
    public void shouldReturnCorrectMetadataFullFeatures() throws Exception {

        XSSFWorkbook workbook = new XSSFWorkbook();

        MempoiTableBuilder mempoiTableBuilder = TestHelper.getTestMempoiTableBuilder(workbook)
                .withAreaReferenceSource("A1:H11");

        final MempoiPivotTableBuilder mempoiPivotTableBuilder = TestHelper
                .getTestMempoiPivotTableBuilder(workbook, null)
                .withAreaReferenceSource("A1:H11")
                .withPosition(new CellReference("J3"));

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .withMempoiTableBuilder(mempoiTableBuilder)
                .withMempoiPivotTableBuilder(mempoiPivotTableBuilder)
                .build();

        final MempoiReport mempoiReport = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(mempoiSheet)
                .build()
                .prepareMempoiReport()
                .get();

        assertNull(mempoiReport.getFile());

        MempoiSheetMetadata mempoiSheetMetadata = mempoiReport.getMempoiSheetMetadata(0);
        assertNull(mempoiSheetMetadata.getSheetName());
        assertEquals(SpreadsheetVersion.EXCEL2007, mempoiSheetMetadata.getSpreadsheetVersion());
        assertOnHeaders(mempoiSheetMetadata, 0, "A1:H1");
        assertOnRows(mempoiSheetMetadata, 11, 10, 1, 10, "A2:H11");
        assertOnCols(mempoiSheetMetadata, 8, 0, 7);
        assertOnSubfooter(mempoiSheetMetadata, 11, "A12:H12");
        assertOnTable(mempoiSheetMetadata, 0, 10, 0, 7, "A1:H11");
        assertOnPivotTable(mempoiSheetMetadata, 2, 9, 0, 0, 10, 7, "J3", "A1:H11");

        assertOnUsedWorkbookConfig(mempoiReport.getUsedWorkbookConfig(),
                mempoiReport.getUsedWorkbookConfig().getWorkbook(),
                true, false, false,
                Collections.singletonList(mempoiSheet), false);
    }
}
