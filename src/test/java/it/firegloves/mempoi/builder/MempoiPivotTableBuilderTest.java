package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.ForceGenerationUtils;
import it.firegloves.mempoi.testutil.TestHelper;
import it.firegloves.mempoi.util.Errors;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.lang.reflect.Constructor;
import java.util.Arrays;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

public final class MempoiPivotTableBuilderTest {

    // TODO add assertions
    // TODO review this tests

    private final Workbook wb = new XSSFWorkbook();

    @Test
    public void fullyPopulated() {

        AssertionHelper.validateMempoiPivotTable(wb, TestHelper.getTestMempoiPivotTable(wb));
    }


    @Test
    public void withInvalidReferenceArea_throwsMempoiException() {

        Arrays.asList(TestHelper.FAILING_AREA_REFERENCES)
                .forEach(areaRef -> {

                    try {
                        MempoiPivotTableBuilder.aMempoiPivotTable()
                                .withWorkbook(wb)
                                .withPosition(TestHelper.POSITION)
                                .withAreaReferenceSource(areaRef)
                                .build();
                    } catch (MempoiException e) {
                        assertEquals(Errors.ERR_AREA_REFERENCE_NOT_VALID, e.getMessage());
                    }
                });
    }

    @Test(expected = MempoiException.class)
    public void withReferenceAreaAndTable_throwsMempoiException() {

        MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .withPosition(TestHelper.POSITION)
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                .withMempoiTableSource(new MempoiTable())
                .build();
    }

    @Test
    public void withReferenceAreaAndTableAndForceGenerationShouldWork() {

        ForceGenerationUtils.executeTestWithForceGeneration(() -> {

            MempoiPivotTableBuilder.aMempoiPivotTable()
                    .withWorkbook(wb)
                    .withPosition(TestHelper.POSITION)
                    .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                    .withMempoiTableSource(new MempoiTable())
                    .build();
        });
    }


    @Test(expected = MempoiException.class)
    public void withSourceSheetAndSourceTableShouldThrowMempoiException() {

        MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .withPosition(TestHelper.POSITION)
                .withMempoiSheetSource(new MempoiSheet(null))
                .withMempoiTableSource(new MempoiTable())
                .build();
    }

    @Test
    public void withSourceSheetAndSourceTableAndForceGenerationShouldWork() {

        ForceGenerationUtils.executeTestWithForceGeneration(() -> {

            MempoiPivotTableBuilder.aMempoiPivotTable()
                    .withWorkbook(wb)
                    .withPosition(TestHelper.POSITION)
                    .withMempoiSheetSource(new MempoiSheet(null))
                    .withMempoiTableSource(new MempoiTable())
                    .build();
        });
    }


    @Test
    public void withoutReferenceAreaButTable() {

        MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .withPosition(TestHelper.POSITION)
                .withMempoiTableSource(new MempoiTable())
                .build();
    }

    @Test
    public void withoutTableButReferenceArea() {

        MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .withPosition(TestHelper.POSITION)
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                .build();
    }


    @Test(expected = MempoiException.class)
    public void withoutReferenceAreaAndTable_throwsMempoiException() {

        MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .build();
    }


    @Test(expected = MempoiException.class)
    public void withoutPosition_throwsMempoiException() {

        MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                .build();
    }


    @Test
    public void withWorkbookNotOfTypeXSSFWorkbook_throwsMempoiException() {

        Arrays.asList(SXSSFWorkbook.class, HSSFWorkbook.class)
                .forEach(wbTypeClass -> {

                    Constructor<? extends Workbook> constructor;
                    Workbook workbook;
                    try {
                        constructor = wbTypeClass.getConstructor();
                        workbook = constructor.newInstance();
                    } catch (Exception e) {
                        throw new RuntimeException();
                    }

                    try {
                        MempoiPivotTableBuilder.aMempoiPivotTable()
                                .withWorkbook(workbook)
                                .withPosition(TestHelper.POSITION)
                                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                                .build();
                    } catch (MempoiException e) {
                        assertEquals(Errors.ERR_PIVOT_TABLE_SUPPORTS_ONLY_XSSF, e.getMessage());
                    }
                });

    }


    @Test(expected = MempoiException.class)
    public void withNullWorkbook_throwsMempoiException() {

        MempoiPivotTableBuilder.aMempoiPivotTable()
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                .build();
    }


    @Test
    public void minPopulated() {

        MempoiPivotTable mempoiPivotTable = MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .withPosition(TestHelper.POSITION)
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                .build();

        assertEquals(wb, mempoiPivotTable.getWorkbook());
        assertNotNull(mempoiPivotTable.getSource().getAreaReference());
        assertNotNull(mempoiPivotTable.getPosition());
    }

    @Test
    public void withColumnLabelColumnsPopulatedManually() {

        MempoiPivotTable mempoiPivotTable = MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .withPosition(TestHelper.POSITION)
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                .addColumnLabelColumns(DataConsolidateFunction.SUM, TestHelper.SUM_COLS_LABEL_COLUMNS_2)
                .addColumnLabelColumns(DataConsolidateFunction.AVERAGE, TestHelper.AVERAGE_COLS_LABEL_COLUMNS_2)
                .build();

        assertEquals(wb, mempoiPivotTable.getWorkbook());
        assertNotNull(mempoiPivotTable.getSource().getAreaReference());
        assertNotNull(mempoiPivotTable.getPosition());

        AssertionHelper.validateColumnLabelColumns(TestHelper.SUM_COLS_LABEL_COLUMNS_2, mempoiPivotTable.getColumnLabelColumns().get(DataConsolidateFunction.SUM));
        AssertionHelper.validateColumnLabelColumns(TestHelper.AVERAGE_COLS_LABEL_COLUMNS_2, mempoiPivotTable.getColumnLabelColumns().get(DataConsolidateFunction.AVERAGE));

    }
}
