/**
 * created by firegloves
 */

package it.firegloves.mempoi.validator;

import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import org.junit.Test;

import java.util.Arrays;

import static org.junit.Assert.*;

public class AreaReferenceValidatorTest {

    private final String[] successfulAreaReferences = { "A1:B5", "C1:C10", "C1:F1", "F10:A1" };
    private final String[] failingAreaReferences = { "A1:5B", "1A:B5", "A:B4", "A1:B", "A1B5", "A1:B5:C6", "", ":", "C1", "A-1:B5" };

    private AreaReferenceValidator areaReferenceValidator = new AreaReferenceValidator();

    @Test
    public void validateAreaReferenceTest_willSuccess() {

        Arrays.stream(successfulAreaReferences).forEach(areaRef -> assertTrue(areaRef, this.areaReferenceValidator.validateAreaReference(areaRef)));
    }

    @Test
    public void validateAreaReferenceTest_willFail() {

        Arrays.stream(failingAreaReferences).forEach(areaRef -> assertFalse(areaRef, this.areaReferenceValidator.validateAreaReference(areaRef)));
    }

    @Test
    public void validateAreaReferenceTest_withNullAreaReference_willFail() {

        assertFalse(this.areaReferenceValidator.validateAreaReference(null));
    }

    @Test
    public void validateAreaReferenceTest_withEmptyAreaReference_willFail() {

        assertFalse(this.areaReferenceValidator.validateAreaReference(""));
    }


    @Test
    public void validateAreaReferenceTestAndThrow_willSuccess() {

        Arrays.stream(successfulAreaReferences).forEach(this.areaReferenceValidator::validateAreaReferenceAndThrow);
    }

    @Test
    public void validateAreaReferenceTestAndThrow_willFail() {

        Arrays.stream(failingAreaReferences).forEach(areaRef -> {

            try {
                this.areaReferenceValidator.validateAreaReferenceAndThrow(areaRef);
            } catch (MempoiException e) {
                assertEquals(Errors.ERR_AREA_REFERENCE_NOT_VALID, e.getMessage());
            }
        });
    }

    @Test(expected = MempoiException.class)
    public void validateAreaReferenceTestAndThrow_withNullAreaReference_willFail() {

        this.areaReferenceValidator.validateAreaReferenceAndThrow(null);
    }

    @Test(expected = MempoiException.class)
    public void validateAreaReferenceTestAndThrow_withEmptyAreaReference_willFail() {

        this.areaReferenceValidator.validateAreaReferenceAndThrow("");
    }
}
