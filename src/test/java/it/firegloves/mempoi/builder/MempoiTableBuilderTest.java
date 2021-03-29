package it.firegloves.mempoi.builder;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.ForceGenerationUtils;
import it.firegloves.mempoi.testutil.TestHelper;
import it.firegloves.mempoi.util.Errors;
import java.lang.reflect.Constructor;
import java.util.Arrays;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;
import org.mockito.MockitoAnnotations;

public class MempoiTableBuilderTest {

    private final Workbook wb = new XSSFWorkbook();

    @Rule
    public ExpectedException exceptionRule = ExpectedException.none();

    @Before
    public void prepare() {
        MockitoAnnotations.initMocks(this);
    }

    @Test
    public void fullyPopulated() {

        MempoiTable mempoiTable = TestHelper.getTestMempoiTable(wb);

        assertEquals(TestHelper.TABLE_NAME, mempoiTable.getTableName());
        assertEquals(TestHelper.DISPLAY_TABLE_NAME, mempoiTable.getDisplayTableName());
        assertEquals(TestHelper.AREA_REFERENCE, mempoiTable.getAreaReferenceSource());
        assertEquals(wb, mempoiTable.getWorkbook());
    }

    @Test
    public void withInvalidReferenceAreaThrowsMempoiException() {

        Arrays.asList(TestHelper.FAILING_AREA_REFERENCES)
                .forEach(areaRef -> {

                    exceptionRule.expect(MempoiException.class);
                    exceptionRule.expectMessage(Errors.ERR_AREA_REFERENCE_NOT_VALID);

                    MempoiTableBuilder.aMempoiTable()
                            .withWorkbook(wb)
                            .withAreaReferenceSource(areaRef)
                            .build();
                });
    }


    @Test
    public void withWorkbookNotOfTypeXSSFWorkbookThrowsMempoiException() {

        Arrays.asList(SXSSFWorkbook.class, HSSFWorkbook.class)
                .forEach(wbTypeClass -> {

                    exceptionRule.expect(MempoiException.class);
                    exceptionRule.expectMessage(Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF);

                    Constructor<? extends Workbook> constructor;
                    Workbook workbook;
                    try {
                        constructor = wbTypeClass.getConstructor();
                        workbook = constructor.newInstance();
                    } catch (Exception e) {
                        throw new MempoiException("wrong error");
                    }

                    MempoiTableBuilder.aMempoiTable()
                                .withWorkbook(workbook)
                                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                                .build();
                });

    }


    @Test(expected = MempoiException.class)
    public void withNullWorkbookThrowsMempoiException() {

        MempoiTableBuilder.aMempoiTable()
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                .build();
    }

    @Test(expected = MempoiException.class)
    public void withNullAreaReferenceAndNoAllSheetDataThrowsMempoiException() {

        MempoiTableBuilder.aMempoiTable()
                .withWorkbook(wb)
                .build();
    }

    @Test
    public void withNullAreaReferenceAndNoAllSheetDataAndForceGenerationShouldUseAllSheetData() {

        ForceGenerationUtils.executeTestWithForceGeneration(() -> {

            MempoiTable mempoiTable = MempoiTableBuilder.aMempoiTable()
                    .withWorkbook(wb)
                    .build();

            assertTrue(mempoiTable.isAllSheetData());
            assertNull(mempoiTable.getAreaReferenceSource());
        });
    }

    @Test
    public void minPopulated() {

        MempoiTable mempoiTable = MempoiTableBuilder.aMempoiTable()
                .withWorkbook(wb)
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                .build();

        assertNotNull(mempoiTable.getTableName());
        assertNotNull(mempoiTable.getDisplayTableName());
        assertEquals(TestHelper.AREA_REFERENCE, mempoiTable.getAreaReferenceSource());
        assertEquals(wb, mempoiTable.getWorkbook());
    }

    @Test
    public void addingExcelTableWithWitheSpacesInDisplayNameWillFail() {

        exceptionRule.expect(MempoiException.class);
        exceptionRule.expectMessage(Errors.ERR_TABLE_DISPLAY_NAME);

        TestHelper.getTestMempoiTableBuilder(wb)
                .withDisplayTableName(TestHelper.DISPLAY_TABLE_NAME.replaceAll("_", " "))
                .build();
    }


    @Test(expected = Test.None.class)
    public void addingExcelTableWithWitheSpacesInDisplayNameAndForceGenerationWillSuccedWithUnderScores() {

        ForceGenerationUtils.executeTestWithForceGeneration(() -> {

            MempoiTable mempoiTable = TestHelper.getTestMempoiTableBuilder(wb)
                    .withDisplayTableName(TestHelper.DISPLAY_TABLE_NAME.replaceAll("_", " "))
                    .build();

            assertEquals(TestHelper.DISPLAY_TABLE_NAME.replaceAll(" ", "_"), mempoiTable.getDisplayTableName());
        });
    }


    @Test(expected = MempoiException.class)
    public void testTableAreaReferenceValidation() {

        MempoiTableBuilder.aMempoiTable()
                .withWorkbook(wb)
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                .withAllSheetData(true)
                .build();
    }

    @Test(expected = Test.None.class /* no exception expected */)
    public void testTableAreaReferenceValidationWithForceGeneration() {

        ForceGenerationUtils.executeTestWithForceGeneration(() -> {

            MempoiTable mempoiTable = MempoiTableBuilder.aMempoiTable()
                    .withWorkbook(wb)
                    .withAreaReferenceSource(TestHelper.AREA_REFERENCE)
                    .withAllSheetData(true)
                    .build();

            assertNotNull(mempoiTable.getTableName());
            assertNotNull(mempoiTable.getDisplayTableName());
            assertNull(mempoiTable.getAreaReferenceSource());
            assertEquals(wb, mempoiTable.getWorkbook());
        });
    }
}
