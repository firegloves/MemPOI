/**
 * created by firegloves
 */

package it.firegloves.mempoi.validator;

import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;

public class WorkbookValidator {

    /**
     * check that the received workbook is of the received type
     *
     * @param workbook the workbook to validate
     * @param wbClazz the expected type
     * @param error the error message to put into the exception if thrown
     * @throws MempoiException if validation fails
     */
    public void validateWorkbookTypeOrThrow(Workbook workbook, Class<? extends Workbook> wbClazz, String error) throws MempoiException {

        if (null == workbook) {
            throw new MempoiException(Errors.ERR_WORKBOOK_NULL);
        }

        if (! workbook.getClass().equals(wbClazz)) {
            throw new MempoiException(! StringUtils.isEmpty(error) ? error : Errors.ERR_WORKBOOK_CLASS_NOT_CORRESPONDING);
        }
    }
}
