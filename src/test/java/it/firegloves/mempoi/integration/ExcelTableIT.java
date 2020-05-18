package it.firegloves.mempoi.integration;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.builder.MempoiTableBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import it.firegloves.mempoi.util.Errors;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Constructor;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Arrays;
import java.util.concurrent.CompletableFuture;

import static org.junit.Assert.assertEquals;

public class ExcelTableIT extends IntegrationBaseIT {

    @Rule
    public ExpectedException exceptionRule = ExpectedException.none();


    @Test
    public void addingExcelTable() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_table.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();

        MempoiTableBuilder mempoiTableBuilder = TestHelper.getTestMempoiTableBuilder(workbook)
                .withAreaReference(TestHelper.AREA_REFERENCE_TABLE_DB_DATA);

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMempoiTableBuilder(mempoiTableBuilder)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .addMempoiSheet(mempoiSheet)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        // validates first sheet
        AssertionHelper.validateGeneratedFilePivotTable(this.createStatement(), fut.get(), TestHelper.MEMPOI_COLUMN_NAMES, TestHelper.MEMPOI_COLUMN_NAMES, new StandardStyleTemplate(), 0);

        XSSFSheet sheet = ((XSSFWorkbook)(TestHelper.loadWorkbookFromDisk(fileDest.getAbsolutePath()))).getSheetAt(0);
        AssertionHelper.validateTable(sheet);
    }


    @Test
    public void addingExcelTableToNonXSSFWorkbook_willFail() {

        // TODO check if all iterations are executed

        Arrays.asList(SXSSFWorkbook.class, HSSFWorkbook.class)
                .forEach(wbTypeClass -> {

                    exceptionRule.expect(MempoiException.class);
                    exceptionRule.expectMessage(Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF);

                    File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_table.xlsx");

                    Constructor<? extends Workbook> constructor;
                    Workbook workbook;
                    try {
                        constructor = wbTypeClass.getConstructor();
                        workbook = constructor.newInstance();
                    } catch (Exception e) {
                        throw new RuntimeException();
                    }

                    MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                            .withPrepStmt(prepStmt)
                            .withMempoiTableBuilder(TestHelper.getTestMempoiTableBuilder(workbook))
                            .build();

                    MempoiBuilder.aMemPOI()
                            .withWorkbook(workbook)
                            .withFile(fileDest)
                            .addMempoiSheet(mempoiSheet)
                            .build();
                });
    }


    @Test
    public void manualTest() throws IOException {

        try (Workbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = (XSSFSheet) wb.createSheet();

            // Set which area the table should be placed in
            AreaReference reference = wb.getCreationHelper().createAreaReference(
                    new CellReference(0, 0), new CellReference(2, 2));

            // Create
            XSSFTable table = sheet.createTable(reference); //creates a table having 3 columns as of area reference
            // but all of those have id 1, so we need repairing
//            table.getCTTable().getTableColumns().getTableColumnArray(1).setId(2);
//            table.getCTTable().getTableColumns().getTableColumnArray(2).setId(3);

            table.setName("Test");
            table.setDisplayName("Test_Table");

            // For now, create the initial style in a low-level way
            table.getCTTable().addNewTableStyleInfo();
            table.getCTTable().getTableStyleInfo().setName("TableStyleMedium2");

            // Style the table
            XSSFTableStyleInfo style = (XSSFTableStyleInfo) table.getStyle();
            style.setName("TableStyleMedium2");
            style.setShowColumnStripes(false);
            style.setShowRowStripes(true);
            style.setFirstColumn(false);
            style.setLastColumn(false);
            style.setShowRowStripes(true);
            style.setShowColumnStripes(true);

            // Set the values for the table
            XSSFRow row;
            XSSFCell cell;
            for (int i = 0; i < 50; i++) {
                // Create row
                row = sheet.createRow(i);
                for (int j = 0; j < 3; j++) {
                    // Create cell
                    cell = row.createCell(j);
                    if (i == 0) {
                        cell.setCellValue("Column" + (j + 1));
                    } else {
                        cell.setCellValue((i + 1.0) * (j + 1.0));
                    }
                }
            }

            // Save
            try (FileOutputStream fileOut = new FileOutputStream(new File(this.outReportFolder.getAbsolutePath(),"table-manuale.xlsx"))) {
                wb.write(fileOut);
            }
        }
    }


    @Override
    public PreparedStatement createStatement() throws SQLException {
        return this.conn.prepareStatement(this.createQuery(TestHelper.TABLE_PIVOT_TABLE, TestHelper.MEMPOI_COLUMN_NAMES, TestHelper.MEMPOI_COLUMN_NAMES, 100));
    }
}