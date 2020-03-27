package it.firegloves.mempoi.integration;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.sql.PreparedStatement;
import java.sql.SQLException;

public class PivotTableIT extends IntegrationBaseIT {

    private XSSFWorkbook wb = new XSSFWorkbook();

    @Before
    public void setup() throws SQLException {
        this.prepStmt = this.createStatement();
    }

    @Test
    public void addPivotTable() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_pivot_table.xlsx");

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(TestHelper.SHEET_NAME)
                .withPrepStmt(prepStmt)
                .withMempoiPivotTableBuilder(TestHelper.getTestMempoiPivotTableBuilder(wb, null))
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


    @Override
    public PreparedStatement createStatement() throws SQLException {
        return this.conn.prepareStatement(this.createQuery(TestHelper.TABLE_PIVOT_TABLE, TestHelper.MEMPOI_COLUMN_NAMES, TestHelper.MEMPOI_COLUMN_NAMES, -1));
    }
}