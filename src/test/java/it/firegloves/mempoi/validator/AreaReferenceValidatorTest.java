/**
 * created by firegloves
 */

package it.firegloves.mempoi.validator;

import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.TestHelper;
import it.firegloves.mempoi.util.Errors;
import org.junit.Test;

import java.util.Arrays;

import static org.junit.Assert.*;

public class AreaReferenceValidatorTest {

    private AreaReferenceValidator areaReferenceValidator = new AreaReferenceValidator();

    @Test
    public void validateAreaReferenceTest_willSuccess() {

        Arrays.stream(TestHelper.SUCCESSFUL_AREA_REFERENCES).forEach(areaRef -> assertTrue(areaRef, this.areaReferenceValidator.validateAreaReference(areaRef)));
    }

    @Test
    public void validateAreaReferenceTest_willFail() {

        Arrays.stream(TestHelper.FAILING_AREA_REFERENCES).forEach(areaRef -> assertFalse(areaRef, this.areaReferenceValidator.validateAreaReference(areaRef)));
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

        Arrays.stream(TestHelper.SUCCESSFUL_AREA_REFERENCES).forEach(this.areaReferenceValidator::validateAreaReferenceAndThrow);
    }

    @Test
    public void validateAreaReferenceTestAndThrow_willFail() {

        Arrays.stream(TestHelper.FAILING_AREA_REFERENCES).forEach(areaRef -> {

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
