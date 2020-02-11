package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.TestHelper;
import it.firegloves.mempoi.util.Errors;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;
import org.mockito.MockitoAnnotations;

import java.lang.reflect.Constructor;
import java.util.Arrays;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

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
        assertEquals(TestHelper.AREA_REFERENCE, mempoiTable.getAreaReference());
        assertEquals(wb, mempoiTable.getWorkbook());
    }

    @Test
    public void withInvalidReferenceArea_throwsMempoiException() {

        Arrays.asList("A1:5B", "1A:B5", "A:B4", "A1:B", "A1B5", "A1:B5:C6", "", ":").stream()
                .forEach(areaRef -> {

                    exceptionRule.expect(MempoiException.class);
                    exceptionRule.expectMessage(Errors.ERR_AREA_REFERENCE_NOT_VALID);

                    MempoiTableBuilder.aMempoiTable()
                            .withWorkbook(wb)
                            .withAreaReference(areaRef)
                            .build();
                });
    }


    @Test
    public void withWorkbookNotOfTypeXSSFWorkbook_throwsMempoiException() {

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
                        throw new RuntimeException();
                    }

                    MempoiTableBuilder.aMempoiTable()
                            .withWorkbook(workbook)
                            .withAreaReference(TestHelper.AREA_REFERENCE)
                            .build();
                });

    }


    @Test(expected = MempoiException.class)
    public void withNullWorkbook_throwsMempoiException() {

        MempoiTableBuilder.aMempoiTable()
                .withAreaReference(TestHelper.AREA_REFERENCE)
                .build();
    }

    @Test(expected = MempoiException.class)
    public void withNullAreaReference_throwsMempoiException() {

        MempoiTableBuilder.aMempoiTable()
                .withWorkbook(wb)
                .build();
    }

    @Test
    public void minPopulated() {

        MempoiTable mempoiTable = MempoiTableBuilder.aMempoiTable()
                .withWorkbook(wb)
                .withAreaReference(TestHelper.AREA_REFERENCE)
                .build();

        assertNotNull(mempoiTable.getTableName());
        assertNotNull(mempoiTable.getDisplayTableName());
        assertEquals(TestHelper.AREA_REFERENCE, mempoiTable.getAreaReference());
        assertEquals(wb, mempoiTable.getWorkbook());
    }
}
