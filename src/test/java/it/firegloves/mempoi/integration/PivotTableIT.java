package it.firegloves.mempoi.integration;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiPivotTableBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;

public class PivotTableIT extends IntegrationBaseIT {

    @Rule
    public ExpectedException exceptionRule = ExpectedException.none();

    private XSSFWorkbook wb = new XSSFWorkbook();

    // TODO  add check if table area are ovelapping?

    @Before
    public void setup() throws SQLException {
        this.prepStmt = this.createStatement();
    }

    @Test
    public void addPivotTableWithPositionAndAreaReference() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_pivot_table_and_position_and_area_reference.xlsx");

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(TestHelper.SHEET_NAME)
                .withPrepStmt(prepStmt)
                .withMempoiPivotTableBuilder(TestHelper.getTestMempoiPivotTableBuilderForIT(wb, null))
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(wb)
                .withFile(fileDest)
                .withStyleTemplate(new StandardStyleTemplate())
                .addMempoiSheet(mempoiSheet)
                .build();

        String fileName = memPOI.prepareMempoiReportToFile().get();

        AssertionHelper.validateGeneratedFilePivotTable(this.createStatement(), fileName, TestHelper.MEMPOI_COLUMN_NAMES, TestHelper.MEMPOI_COLUMN_NAMES, new StandardStyleTemplate(), 0);

        Workbook loadedWb = TestHelper.openFile(fileName);
        XSSFPivotTable pivotTable = ((XSSFSheet) loadedWb.getSheet(TestHelper.SHEET_NAME)).getPivotTables().get(0);
        AssertionHelper.assertPivotTable(pivotTable, mempoiSheet.getMempoiPivotTable().get(), mempoiSheet.getColumnList());
    }

    @Test
    public void addPivotTableWithAreaReferenceSourceOnDifferentSheet() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_pivot_table_and_area_referencce_on_different_sheet.xlsx");

        MempoiSheet mempoiSheet1 = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(TestHelper.SHEET_NAME)
                .withPrepStmt(prepStmt)
                .build();

        MempoiPivotTableBuilder mempoiPivotTableBuilder = TestHelper.getTestMempoiPivotTableBuilderForIT(wb, mempoiSheet1);

        MempoiSheet mempoiSheet2 = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(TestHelper.SHEET_NAME_2)
                .withPrepStmt(this.createStatement())
                .withMempoiPivotTableBuilder(mempoiPivotTableBuilder)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(wb)
                .withFile(fileDest)
                .withStyleTemplate(new StandardStyleTemplate())
                .addMempoiSheet(mempoiSheet1)
                .addMempoiSheet(mempoiSheet2)
                .build();

        String fileName = memPOI.prepareMempoiReportToFile().get();

        AssertionHelper.validateGeneratedFilePivotTable(this.createStatement(), fileName, TestHelper.MEMPOI_COLUMN_NAMES, TestHelper.MEMPOI_COLUMN_NAMES, new StandardStyleTemplate(), 0);

        Workbook loadedWb = TestHelper.openFile(fileName);
        XSSFPivotTable pivotTable = ((XSSFSheet) loadedWb.getSheet(TestHelper.SHEET_NAME_2)).getPivotTables().get(0);
        AssertionHelper.assertPivotTable(pivotTable, mempoiSheet2.getMempoiPivotTable().get(), mempoiSheet2.getColumnList());
    }



    @Test
    public void addPivotTableWithPositionAndTableSource() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_pivot_table_and_position_and_table.xlsx");

        MempoiTable mempoiTable = TestHelper.getTestMempoiTableBuilder(wb)
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE_TABLE_DB_DATA)
                .build();

        MempoiPivotTableBuilder mempoiPivotTableBuilder = TestHelper.getTestMempoiPivotTableBuilderForIT(wb, null)
                .withAreaReferenceSource(null)
                .withMempoiTableSource(mempoiTable);

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(TestHelper.SHEET_NAME)
                .withPrepStmt(prepStmt)
                .withMempoiTable(mempoiTable)
                .withMempoiPivotTableBuilder(mempoiPivotTableBuilder)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(wb)
                .withFile(fileDest)
                .withStyleTemplate(new StandardStyleTemplate())
                .addMempoiSheet(mempoiSheet)
                .build();

        String fileName = memPOI.prepareMempoiReportToFile().get();

        AssertionHelper.validateGeneratedFilePivotTable(this.createStatement(), fileName, TestHelper.MEMPOI_COLUMN_NAMES, TestHelper.MEMPOI_COLUMN_NAMES, new StandardStyleTemplate(), 0);

        Workbook loadedWb = TestHelper.openFile(fileName);
        XSSFPivotTable pivotTable = ((XSSFSheet) loadedWb.getSheet(TestHelper.SHEET_NAME)).getPivotTables().get(0);
        AssertionHelper.assertPivotTable(pivotTable, mempoiSheet.getMempoiPivotTable().get(), mempoiSheet.getColumnList());
    }


    @Test
    public void addPivotTableWithPositionAndTableSourceOnDifferentSheet() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_pivot_table_and_position_and_table_on_different_sheet.xlsx");

        MempoiTable mempoiTable = TestHelper.getTestMempoiTableBuilder(wb)
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE_TABLE_DB_DATA)
                .build();

        MempoiSheet mempoiSheet1 = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(TestHelper.SHEET_NAME)
                .withPrepStmt(prepStmt)
                .withMempoiTable(mempoiTable)
                .build();

        MempoiPivotTableBuilder mempoiPivotTableBuilder = TestHelper.getTestMempoiPivotTableBuilderForIT(wb, null)
                .withAreaReferenceSource(null)
                .withMempoiTableSource(mempoiTable);

        MempoiSheet mempoiSheet2 = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(TestHelper.SHEET_NAME_2)
                .withPrepStmt(this.createStatement())
                .withMempoiPivotTableBuilder(mempoiPivotTableBuilder)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(wb)
                .withFile(fileDest)
                .withStyleTemplate(new StandardStyleTemplate())
                .addMempoiSheet(mempoiSheet1)
                .addMempoiSheet(mempoiSheet2)
                .build();

        String fileName = memPOI.prepareMempoiReportToFile().get();

        AssertionHelper.validateGeneratedFilePivotTable(this.createStatement(), fileName, TestHelper.MEMPOI_COLUMN_NAMES, TestHelper.MEMPOI_COLUMN_NAMES, new StandardStyleTemplate(), 0);

        Workbook loadedWb = TestHelper.openFile(fileName);
        XSSFPivotTable pivotTable = ((XSSFSheet) loadedWb.getSheet(TestHelper.SHEET_NAME_2)).getPivotTables().get(0);
        AssertionHelper.assertPivotTable(pivotTable, mempoiSheet2.getMempoiPivotTable().get(), mempoiSheet2.getColumnList());
    }


    @Test(expected = MempoiException.class)
    public void addPivotTableWithTableAsSourceContainedInNotStillGeneratedSheetWillThrowMempoiException() throws Throwable {

        MempoiTable mempoiTable = TestHelper.getTestMempoiTableBuilder(wb)
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE_TABLE_DB_DATA)
                .build();

        MempoiSheet mempoiSheet1 = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(TestHelper.SHEET_NAME)
                .withPrepStmt(prepStmt)
                .withMempoiTable(mempoiTable)
                .build();

        MempoiPivotTableBuilder mempoiPivotTableBuilder = TestHelper.getTestMempoiPivotTableBuilderForIT(wb, null)
                .withAreaReferenceSource(null)
                .withMempoiTableSource(mempoiTable);

        MempoiSheet mempoiSheet2 = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(TestHelper.SHEET_NAME_2)
                .withPrepStmt(this.createStatement())
                .withMempoiPivotTableBuilder(mempoiPivotTableBuilder)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(wb)
                .withFile(new File("temp"))
                .withStyleTemplate(new StandardStyleTemplate())
                .addMempoiSheet(mempoiSheet2)
                .addMempoiSheet(mempoiSheet1)
                .build();
        try {
            memPOI.prepareMempoiReportToFile().get();
        } catch (Exception e) {
            throw e.getCause();
        }
    }


    @Override
    public PreparedStatement createStatement() throws SQLException {
        return this.conn.prepareStatement(this.createQuery(TestHelper.TABLE_PIVOT_TABLE, TestHelper.MEMPOI_COLUMN_NAMES, TestHelper.MEMPOI_COLUMN_NAMES, -1));
    }
}
