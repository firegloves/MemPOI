/**
 * created by firegloves
 */

package it.firegloves.mempoi.validator;

import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import java.util.regex.Pattern;
import org.apache.commons.lang3.StringUtils;

public class AreaReferenceValidator {

    private final Pattern areaReferencePattern = Pattern.compile("^[a-zA-Z]+\\d+\\:[a-zA-Z]+\\d+$");

    /**
     * check that areaReference respect the needed format like A1:E5
     *
     * @param areaReference the area reference to validate
     * @return true if area reference is valid, false otherwise
     */
    public boolean validateAreaReference(String areaReference) {

        return ! StringUtils.isEmpty(areaReference) && areaReferencePattern.matcher(areaReference).find();
    }

    /**
     * check that areaReference respect the needed format like A1:E5
     *
     * @param areaReference the area reference to validate
     * @throws MempoiException if validation fails
     */
    public void validateAreaReferenceAndThrow(String areaReference) throws MempoiException {

        if (! this.validateAreaReference(areaReference)) {
            throw new MempoiException(Errors.ERR_AREA_REFERENCE_NOT_VALID);
        }
    }


//    /**
//     * return tru
//     * @param areaRefOne
//     * @param areaRefTwo
//     * @return
//     */
//    public boolean areAreaOverlapping(AreaReference areaRefOne, AreaReference areaRefTwo) {
//
//        String[] areaOneFirstCellRefParts = areaRefOne.getFirstCell().getCellRefParts();
//        String[] areaOneLastCellRefParts = areaRefOne.getLastCell().getCellRefParts();
//        String[] areaTwoFirstCellRefParts = areaRefTwo.getFirstCell().getCellRefParts();
//        String[] areaTwoLastCellRefParts = areaRefTwo.getLastCell().getCellRefParts();
//
//
//    }
}
