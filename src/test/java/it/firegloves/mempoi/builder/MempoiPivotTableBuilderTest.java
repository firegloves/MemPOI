package it.firegloves.mempoi.builder;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.ForceGenerationUtils;
import it.firegloves.mempoi.testutil.TestHelper;
import it.firegloves.mempoi.util.Errors;
import java.lang.reflect.Constructor;
import java.util.Arrays;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;

public final class MempoiPivotTableBuilderTest {

    // TODO add assertions
    // TODO review this tests

    @Rule
    public ExpectedException exceptionRule = ExpectedException.none();

    private final Workbook wb = new XSSFWorkbook();

    @Test
    public void fullyPopulated() {

        AssertionHelper.validateMempoiPivotTable(wb, TestHelper.getTestMempoiPivotTable(wb));
    }


    @Test
    public void withInvalidReferenceAreaThrowsMempoiException() {

        Arrays.asList(TestHelper.FAILING_AREA_REFERENCES)
                .forEach(areaRef -> {

                    exceptionRule.expect(MempoiException.class);
                    exceptionRule.expectMessage(Errors.ERR_AREA_REFERENCE_NOT_VALID);

                    MempoiPivotTableBuilder.aMempoiPivotTable()
                            .withWorkbook(wb)
                            .withPosition(TestHelper.POSITION)
                            .withAreaReferenceSource(areaRef)
                            .build();
                });
    }

    @Test(expected = MempoiException.class)
    public void withReferenceAreaAndTableThrowsMempoiException() {

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
    public void withoutReferenceAreaAndTableThrowsMempoiException() {

        MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .build();
    }


    @Test(expected = MempoiException.class)
    public void withoutPositionThrowsMempoiException() {

        MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                .build();
    }


    @Test
    public void withWorkbookNotOfTypeXSSFWorkbookThrowsMempoiException() {

        Arrays.asList(SXSSFWorkbook.class, HSSFWorkbook.class)
                .forEach(wbTypeClass -> {

                    exceptionRule.expect(MempoiException.class);
                    exceptionRule.expectMessage(Errors.ERR_PIVOT_TABLE_SUPPORTS_ONLY_XSSF);

                    Constructor<? extends Workbook> constructor;
                    Workbook workbook;
                    try {
                        constructor = wbTypeClass.getConstructor();
                        workbook = constructor.newInstance();
                    } catch (Exception e) {
                        throw new MempoiException("wrong error");
                    }

                    MempoiPivotTableBuilder.aMempoiPivotTable()
                                .withWorkbook(workbook)
                                .withPosition(TestHelper.POSITION)
                                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                                .build();
                });

    }


    @Test(expected = MempoiException.class)
    public void withNullWorkbookThrowsMempoiException() {

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
