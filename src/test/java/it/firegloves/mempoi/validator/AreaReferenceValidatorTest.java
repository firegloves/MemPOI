/**
 * created by firegloves
 */

package it.firegloves.mempoi.validator;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.TestHelper;
import it.firegloves.mempoi.util.Errors;
import java.util.Arrays;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;

public class AreaReferenceValidatorTest {

    @Rule
    public ExpectedException exceptionRule = ExpectedException.none();

    private AreaReferenceValidator areaReferenceValidator = new AreaReferenceValidator();

    @Test
    public void validateAreaReferenceTestWillSuccess() {

        Arrays.stream(TestHelper.SUCCESSFUL_AREA_REFERENCES).forEach(areaRef ->
                assertTrue(areaRef, this.areaReferenceValidator.validateAreaReference(areaRef)));
    }

    @Test
    public void validateAreaReferenceTestWillFail() {

        Arrays.stream(TestHelper.FAILING_AREA_REFERENCES)
                .forEach(areaRef -> assertFalse(areaRef, this.areaReferenceValidator.validateAreaReference(areaRef)));
    }

    @Test
    public void validateAreaReferenceTestWithNullAreaReferenceWillFail() {

        assertFalse(this.areaReferenceValidator.validateAreaReference(null));
    }

    @Test
    public void validateAreaReferenceTestWithEmptyAreaReferenceWillFail() {

        assertFalse(this.areaReferenceValidator.validateAreaReference(""));
    }


    @Test(expected = Test.None.class)
    public void validateAreaReferenceTestAndThrowWillSuccess() {

        Arrays.stream(TestHelper.SUCCESSFUL_AREA_REFERENCES)
                .forEach(this.areaReferenceValidator::validateAreaReferenceAndThrow);
    }

    @Test
    public void validateAreaReferenceTestAndThrowWillFail() {

        Arrays.stream(TestHelper.FAILING_AREA_REFERENCES).forEach(areaRef -> {

            exceptionRule.expect(MempoiException.class);
            exceptionRule.expectMessage(Errors.ERR_AREA_REFERENCE_NOT_VALID);

            this.areaReferenceValidator.validateAreaReferenceAndThrow(areaRef);
        });
    }

    @Test(expected = MempoiException.class)
    public void validateAreaReferenceTestAndThrowWithNullAreaReferenceWillFail() {

        this.areaReferenceValidator.validateAreaReferenceAndThrow(null);
    }

    @Test(expected = MempoiException.class)
    public void validateAreaReferenceTestAndThrowWithEmptyAreaReferenceWillFail() {

        this.areaReferenceValidator.validateAreaReferenceAndThrow("");
    }
}
