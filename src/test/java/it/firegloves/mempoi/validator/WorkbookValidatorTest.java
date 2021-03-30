/**
 * created by firegloves
 */

package it.firegloves.mempoi.validator;

import it.firegloves.mempoi.exception.MempoiException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;

public class WorkbookValidatorTest {

    @Rule
    public ExpectedException exceptionRule = ExpectedException.none();
    private WorkbookValidator areaReferenceValidator = new WorkbookValidator();
    private String error = "Error";

    @Test(expected = Test.None.class)
    public void validateWorkbookTypeAndThrowWillSuccess() {

        Workbook[] workbooks = {new XSSFWorkbook(), new HSSFWorkbook(), new SXSSFWorkbook()};
        Arrays.asList(workbooks)
                .forEach(wb -> this.areaReferenceValidator.validateWorkbookTypeOrThrow(wb, wb.getClass(), error));
    }

    @Test(expected = Test.None.class)
    public void validateWorkbookTypeAndThrowDifferentClassWillFail() {

        Map<Workbook, List<Class<? extends Workbook>>> map = new HashMap<>();
        map.put(new XSSFWorkbook(), Arrays.asList(HSSFWorkbook.class, SXSSFWorkbook.class));
        map.put(new HSSFWorkbook(), Arrays.asList(XSSFWorkbook.class, SXSSFWorkbook.class));
        map.put(new SXSSFWorkbook(), Arrays.asList(HSSFWorkbook.class, XSSFWorkbook.class));

        map.keySet().forEach(wb -> {

            map.get(wb).forEach(wbClazz -> {

                exceptionRule.expect(MempoiException.class);
                exceptionRule.expectMessage(error);

                this.areaReferenceValidator.validateWorkbookTypeOrThrow(wb, wbClazz, error);
            });
        });
    }


    @Test(expected = MempoiException.class)
    public void validateWorkbookTypeAndThrowWithNullWorkbookWillFail() {

        this.areaReferenceValidator.validateWorkbookTypeOrThrow(null, HSSFWorkbook.class, error);
    }

    @Test(expected = MempoiException.class)
    public void validateWorkbookTypeAndThrowWithNullClassWillFail() {

        this.areaReferenceValidator.validateWorkbookTypeOrThrow(new HSSFWorkbook(), null, error);
    }

    @Test(expected = MempoiException.class)
    public void validateWorkbookTypeAndThrowWithNullErrorWillFail() {

        this.areaReferenceValidator.validateWorkbookTypeOrThrow(new HSSFWorkbook(), XSSFWorkbook.class, null);
    }
}
