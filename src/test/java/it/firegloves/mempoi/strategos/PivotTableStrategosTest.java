package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.builder.MempoiPivotTableBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.builder.MempoiTableBuilder;
import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.PrivateAccessHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import it.firegloves.mempoi.util.Errors;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.sql.PreparedStatement;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.IntStream;

import static org.junit.Assert.*;

public class PivotTableStrategosTest {

    private XSSFWorkbook wb;
    private XSSFSheet sheet;
    private PivotTableStrategos pivotTableStrategos;

    @Mock
    private PreparedStatement prepStmt;

    @Rule
    public ExpectedException exceptionRule = ExpectedException.none();


    @Before
    public void prepare() {
        MockitoAnnotations.initMocks(this);

        this.wb = new XSSFWorkbook();
        this.sheet = wb.createSheet();
        this.initSheet(this.sheet);
        this.pivotTableStrategos = new PivotTableStrategos(new WorkbookConfig().setWorkbook(wb));
    }

    /**
     * prepares the sheet to manage the pivot table
     */
    private void initSheet(Sheet sheet) {

        for (int i = 0; i < 10; i++) {
            Row row = sheet.createRow(i);

            for (int k = 0; k < 10; k++) {
                row.createCell(k);
            }
        }
    }


    /******************************************************************************************************************
     * createPivotTable
     *****************************************************************************************************************/

    // TODO add more asserts on the generated table: check that the table is created respecting the received instructions
    @Test
    public void createPivotTableWithAreaReferece() throws Exception {

        MempoiSheet mempoiSheet = new MempoiSheet(prepStmt)
                .setSheet(sheet);

        MempoiPivotTable mempoiPivotTable = MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                .withPosition(TestHelper.POSITION)
                .build();

        Method createPivotTableMethod = PrivateAccessHelper.getAccessibleMethod(this.pivotTableStrategos, "createPivotTable", MempoiSheet.class, MempoiPivotTable.class);
        XSSFPivotTable pivotTable = (XSSFPivotTable) createPivotTableMethod.invoke(this.pivotTableStrategos, mempoiSheet, mempoiPivotTable);

        assertNotNull(pivotTable);
        assertEquals(sheet, pivotTable.getParentSheet());
        assertEquals(sheet, pivotTable.getDataSheet());
    }

    @Test
    public void createPivotTableWithTable() throws Exception {


        MempoiSheet mempoiSheet = new MempoiSheet(prepStmt)
                .setSheet(sheet);

        XSSFTable table = sheet.createTable(new AreaReference(TestHelper.AREA_REFERENCE, this.wb.getSpreadsheetVersion()));
        MempoiTable mempoiTable = MempoiTableBuilder.aMempoiTable()
                .withWorkbook(wb)
                .withAreaReference(TestHelper.AREA_REFERENCE)
                .build()
                .setTable(table);

        MempoiPivotTable mempoiPivotTable = MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .withMempoiTableSource(mempoiTable)
                .withPosition(TestHelper.POSITION)
                .build();

        Method createPivotTableMethod = PrivateAccessHelper.getAccessibleMethod(this.pivotTableStrategos, "createPivotTable", MempoiSheet.class, MempoiPivotTable.class);
        XSSFPivotTable pivotTable = (XSSFPivotTable) createPivotTableMethod.invoke(this.pivotTableStrategos, mempoiSheet, mempoiPivotTable);

        assertNotNull(pivotTable);
        assertEquals(sheet, pivotTable.getParentSheet());
        assertEquals(sheet, pivotTable.getDataSheet());
    }


    @Test
    public void createPivotTableWithTableOnDifferentSheet() throws Exception {

        XSSFSheet secondSheet = this.wb.createSheet("SecondSheet");
        this.initSheet(secondSheet);
        MempoiSheet mempoiSecondSheet = new MempoiSheet(prepStmt)
                .setSheet(secondSheet);

        MempoiSheet mempoiSheet = new MempoiSheet(prepStmt)
                .setSheet(sheet);

        XSSFTable table = sheet.createTable(new AreaReference(TestHelper.AREA_REFERENCE, this.wb.getSpreadsheetVersion()));
        MempoiTable mempoiTable = MempoiTableBuilder.aMempoiTable()
                .withWorkbook(wb)
                .withAreaReference(TestHelper.AREA_REFERENCE)
                .build()
                .setTable(table);

        MempoiPivotTable mempoiPivotTable = MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
//                .withMempoiSheetSource(mempoiSecondSheet)
                .withMempoiTableSource(mempoiTable)
                .withPosition(TestHelper.POSITION)
                .build();

        Method createPivotTableMethod = PrivateAccessHelper.getAccessibleMethod(this.pivotTableStrategos, "createPivotTable", MempoiSheet.class, MempoiPivotTable.class);
        XSSFPivotTable pivotTable = (XSSFPivotTable) createPivotTableMethod.invoke(this.pivotTableStrategos, mempoiSheet, mempoiPivotTable);

        assertNotNull(pivotTable);
        assertEquals(sheet, pivotTable.getParentSheet());
        assertEquals(table.getXSSFSheet(), pivotTable.getDataSheet());
    }


    /******************************************************************************************************************
     *                          addColumnsToPivotTable
     *****************************************************************************************************************/

    // UNTESTABLE BECAUSE IT'S PRIVATE AND IT RECEIVES A FUNCTION -> FUNCTION IS NOT CASTABLE TO OBJECT

//    @Test
//    public void addColumnsToPivotTable() throws Exception {
//
//        XSSFPivotTable pivotTable = sheet.createPivotTable(new AreaReference(TestHelper.AREA_REFERENCE, wb.getSpreadsheetVersion()), TestHelper.POSITION);
//        List<MempoiColumn> mempoiColumnList = TestHelper.getMempoiColumnList(wb);
//
//        Method addColumnsToPivotTableMethod = PrivateAccessHelper.getAccessibleMethod(this.pivotTableStrategos, "addColumnsToPivotTable", List.class, List.class, Consumer.class);
//        addColumnsToPivotTableMethod.invoke(this.pivotTableStrategos, Arrays.asList(TestHelper.MEMPOI_COLUMN_AGE), mempoiColumnList, (Integer i) -> pivotTable.addRowLabel(i));
//    }


    /******************************************************************************************************************
     *                          addPivotTable
     *****************************************************************************************************************/

    @Test
    public void addPivotTable() throws Exception {

        MempoiSheet mempoiSheet = TestHelper.getMempoiSheetWithMempoiColumns(wb, prepStmt).setSheet(sheet);

        MempoiPivotTable mempoiPivotTable = TestHelper.getTestMempoiPivotTable(wb, mempoiSheet);

        Method addPivotTableMethod = PrivateAccessHelper.getAccessibleMethod(this.pivotTableStrategos, "addPivotTable", MempoiSheet.class, MempoiPivotTable.class);
        addPivotTableMethod.invoke(this.pivotTableStrategos, mempoiSheet, mempoiPivotTable);

        XSSFPivotTable pivotTable = sheet.getPivotTables().get(0);
        AssertionHelper.assertPivotTable(pivotTable);
    }

    @Test
    public void addPivotTableWithNullValues() throws Exception {

        MempoiSheet mempoiSheet = TestHelper.getMempoiSheetWithMempoiColumns(wb, prepStmt).setSheet(sheet);

        MempoiPivotTable mempoiPivotTable = TestHelper.getTestMempoiPivotTableBuilder(wb)
                .withMempoiSheetSource(mempoiSheet)
                .withRowLabelColumns(null)
                .withColumnLabelColumns(null)
                .withReportFilterColumns(null)
                .build();

        Method addPivotTableMethod = PrivateAccessHelper.getAccessibleMethod(this.pivotTableStrategos, "addPivotTable", MempoiSheet.class, MempoiPivotTable.class);
        addPivotTableMethod.invoke(this.pivotTableStrategos, mempoiSheet, mempoiPivotTable);

        XSSFPivotTable pivotTable = sheet.getPivotTables().get(0);
        AssertionHelper.assertPivotTable(pivotTable, mempoiPivotTable, TestHelper.getMempoiColumnList(wb));
    }


    @Test(expected = Exception.class)
    public void addPivotTablewithNullMempoiSheet_willFail() throws Exception {

        MempoiSheet mempoiSheet = TestHelper.getMempoiSheetWithMempoiColumns(wb, prepStmt).setSheet(sheet);

        MempoiPivotTable mempoiPivotTable = TestHelper.getTestMempoiPivotTable(wb, mempoiSheet);

        Method addPivotTableMethod = PrivateAccessHelper.getAccessibleMethod(this.pivotTableStrategos, "addPivotTable", MempoiSheet.class, MempoiPivotTable.class);
        addPivotTableMethod.invoke(this.pivotTableStrategos, null, mempoiPivotTable);
    }

    @Test(expected = Exception.class)
    public void addPivotTablewithNullMempoiPivotTable_willFail() throws Exception {

        MempoiSheet mempoiSheet = TestHelper.getMempoiSheetWithMempoiColumns(wb, prepStmt).setSheet(sheet);

        Method addPivotTableMethod = PrivateAccessHelper.getAccessibleMethod(this.pivotTableStrategos, "addPivotTable", MempoiSheet.class, MempoiPivotTable.class);
        addPivotTableMethod.invoke(this.pivotTableStrategos, mempoiSheet, null);
    }

    @Test(expected = Exception.class)
    public void addPivotTablewitAllNullParams_willFail() throws Exception {

        Method addPivotTableMethod = PrivateAccessHelper.getAccessibleMethod(this.pivotTableStrategos, "addPivotTable", MempoiSheet.class, MempoiPivotTable.class);
        addPivotTableMethod.invoke(this.pivotTableStrategos, null, null);
    }


    /******************************************************************************************************************
     *                          managePivotTable
     *****************************************************************************************************************/

    @Test
    public void managePivotTableTest() throws Exception {

        MempoiPivotTable mempoiPivotTable = TestHelper.getTestMempoiPivotTableBuilder(wb)
                .withMempoiSheetSource(null)
                .build();

        MempoiSheet mempoiSheet = TestHelper.getMempoiSheetBuilder(wb, prepStmt)
                .withMempoiPivotTable(mempoiPivotTable)
                .build()
                .setSheet(sheet);

        Field columnList = PrivateAccessHelper.getPrivateField(mempoiSheet, "columnList");
        columnList.set(mempoiSheet, TestHelper.getMempoiColumnList(wb));

        this.pivotTableStrategos.manageMempoiPivotTable(mempoiSheet);

        AssertionHelper.assertPivotTableIntoSheet(mempoiSheet);
    }


    @Test(expected = MempoiException.class)
    public void managePivotTableTest_withoutSheet_willFail() {

        MempoiPivotTable mempoiPivotTable = TestHelper.getTestMempoiPivotTableBuilder(wb)
                .withMempoiSheetSource(null)
                .build();

        MempoiSheet mempoiSheet = TestHelper.getMempoiSheetBuilder(wb, prepStmt)
                .withMempoiPivotTable(mempoiPivotTable)
                .build();

        this.pivotTableStrategos.manageMempoiPivotTable(mempoiSheet);
    }

    @Test
    public void managePivotTableTest_withoutMempoiPivotTable() {

        MempoiSheet mempoiSheet = TestHelper.getMempoiSheetBuilder(wb, prepStmt)
                .build()
                .setSheet(sheet);;

        this.pivotTableStrategos.manageMempoiPivotTable(mempoiSheet);

        assertEquals(0, ((XSSFSheet)mempoiSheet.getSheet()).getPivotTables().size());
    }
}
