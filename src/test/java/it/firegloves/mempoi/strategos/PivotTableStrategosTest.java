package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.builder.MempoiPivotTableBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.builder.MempoiTableBuilder;
import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import it.firegloves.mempoi.exception.MempoiException;
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
import java.lang.reflect.Method;
import java.sql.PreparedStatement;
import java.util.Arrays;
import java.util.List;
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

        for (int i=0; i<10; i++) {
            Row row = sheet.createRow(i);

            for (int k=0; k<10; k++) {
                row.createCell(k);
            }
        }
    }


    /******************************************************************************************************************
     * createPivotTable
     *****************************************************************************************************************/


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
                .withMempoiSheetSource(mempoiSecondSheet)
                .withMempoiTableSource(mempoiTable)
                .withPosition(TestHelper.POSITION)
                .build();

        Method createPivotTableMethod = PrivateAccessHelper.getAccessibleMethod(this.pivotTableStrategos, "createPivotTable", MempoiSheet.class, MempoiPivotTable.class);
        XSSFPivotTable pivotTable = (XSSFPivotTable) createPivotTableMethod.invoke(this.pivotTableStrategos, mempoiSheet, mempoiPivotTable);

        assertNotNull(pivotTable);
    }


//    /******************************************************************************************************************
//     *                          addTable
//     *****************************************************************************************************************/
//
//
//    @Test
//    public void addsAnExcelTable() throws Exception {
//
//        XSSFSheet sheet = wb.createSheet();
//        MempoiSheet mempoiSheet = new MempoiSheet(prepStmt)
//                .setSheet(sheet);
//
//        MempoiTable mempoiTable = TestHelper.getTestMempoiTable(wb);
//
//        Method addTableMethod = PrivateAccessHelper.getAccessibleMethod(tableStrategos, "addTable", MempoiSheet.class, MempoiTable.class);
//        addTableMethod.invoke(tableStrategos, mempoiSheet, mempoiTable);
//
//        XSSFTable table = sheet.getTables().get(0);
//
//        assertEquals(1, sheet.getTables().size());
//        this.validateTableColumns(table);
//        assertTrue(table.getCTTable().isSetAutoFilter());
//        assertEquals(TestHelper.TABLE_NAME, table.getCTTable().getName());
//        assertEquals(TestHelper.DISPLAY_TABLE_NAME, table.getCTTable().getDisplayName());
//    }
//
//
//    /******************************************************************************************************************
//     *                          manageMempoiTable
//     *****************************************************************************************************************/
//
//    @Test
//    public void manageMempoiTable() {
//
//        XSSFSheet sheet = wb.createSheet();
//
//        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
//                .withPrepStmt(prepStmt)
//                .withMempoiTableBuilder(TestHelper.getTestMempoiTableBuilder(wb))
//                .build()
//                .setSheet(sheet);
//
//        tableStrategos.manageMempoiTable(mempoiSheet);
//
//        XSSFTable table = sheet.getTables().get(0);
//
//        assertEquals(1, sheet.getTables().size());
//        this.validateTableColumns(table);
//        assertTrue(table.getCTTable().isSetAutoFilter());
//        assertEquals(TestHelper.TABLE_NAME, table.getCTTable().getName());
//        assertEquals(TestHelper.DISPLAY_TABLE_NAME, table.getCTTable().getDisplayName());
//    }
//
//
//    @Test
//    public void manageMempoiTable_withSheetNotOfTypeXSSFSheet_shouldFail() throws Exception {
//
//        Arrays.asList(SXSSFWorkbook.class, HSSFWorkbook.class)
//                .forEach(wbTypeClass -> {
//
//                    exceptionRule.expect(MempoiException.class);
//                    exceptionRule.expectMessage(Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF);
//
//                    Constructor<? extends Workbook> constructor;
//                    Workbook workbook;
//                    try {
//                        constructor = wbTypeClass.getConstructor();
//                        workbook = constructor.newInstance();
//                    } catch (Exception e) {
//                        throw new RuntimeException();
//                    }
//
//                    Sheet sheet = workbook.createSheet();
//                    MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
//                            .withPrepStmt(prepStmt)
//                            .withMempoiTableBuilder(TestHelper.getTestMempoiTableBuilder(wb))
//                            .build();
//
//                    tableStrategos.manageMempoiTable(mempoiSheet);
//                });
//    }
//
//
//
//    /******************************************************************************************************************
//     *                          generic validations
//     *****************************************************************************************************************/
//
//    private void validateTableColumns(XSSFTable table) {
//
//        List<CTTableColumn> tableColumnList = table.getCTTable().getTableColumns().getTableColumnList();
//
//        assertEquals(5, tableColumnList.size());
//        IntStream.range(0, tableColumnList.size()).forEachOrdered(i -> assertEquals(i, tableColumnList.get(i).getId()));
//    }
}
