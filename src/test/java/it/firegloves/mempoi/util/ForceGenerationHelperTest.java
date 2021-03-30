package it.firegloves.mempoi.util;

import static org.junit.Assert.assertTrue;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.verifyZeroInteractions;
import static org.mockito.internal.verification.VerificationModeFactory.times;

import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.ForceGenerationUtils;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;
import org.slf4j.Logger;

public class ForceGenerationHelperTest {

    @Mock
    private Logger logger;

    private final String exceptionError = "WTF! Error!";
    private final String logError = "WTF! Error in the log!";

    @Rule
    public ExpectedException exceptionRule = ExpectedException.none();

    private boolean handlerTest;

    @Before
    public void setUp() {
        this.handlerTest = false;
        MockitoAnnotations.initMocks(this);
    }

    @Test
    public void withNoForceGenerationWillThrowException() {

        exceptionRule.expect(MempoiException.class);
        exceptionRule.expectMessage(exceptionError);

        ForceGenerationHelper.manageForceGeneration(
                new MempoiException(exceptionError),
                logError,
                logger);

        verifyZeroInteractions(logger);
    }

    @Test
    public void withForceGenerationWillLogAndExecuteHandler() {

        ForceGenerationUtils.executeTestWithForceGeneration(() -> {

            ForceGenerationHelper.manageForceGeneration(
                    new MempoiException(exceptionError),
                    logError,
                    logger,
                    () -> handlerTest = true);

            verify(logger, times(1)).error(logError);
            assertTrue(handlerTest);
        });
    }


    @Test
    public void withForceGenerationAndNoHandlerWillLog() {

        ForceGenerationUtils.executeTestWithForceGeneration(() -> {

            ForceGenerationHelper.manageForceGeneration(
                    new MempoiException(exceptionError),
                    logError,
                    logger);

            verify(logger, times(1)).error(logError);
        });
    }


    @Test
    public void withNoForceGenerationAndNoExceptionWillDoNothing() {

        ForceGenerationHelper.manageForceGeneration(
                logError,
                logger,
                () -> handlerTest = true);
    }
}
