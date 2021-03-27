/**
 * created by firegloves
 */

package it.firegloves.mempoi.exception;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import it.firegloves.mempoi.util.Errors;
import java.io.IOException;
import org.junit.Test;

public class MempoiExceptionTest {

    @Test
    public void mempoiException() {

        MempoiException e = new MempoiException();
        assertNotNull(e);
    }

    @Test
    public void mempoiExceptionErrMexTest() {

        MempoiException e = new MempoiException(Errors.ERR_FILE_TO_EXPORT_PARENT_DIR_CANT_BE_CREATED);
        assertEquals(Errors.ERR_FILE_TO_EXPORT_PARENT_DIR_CANT_BE_CREATED, e.getMessage());
    }

    @Test
    public void mempoiExceptionErrMexThrowableTest() {

        MempoiException e = new MempoiException(Errors.ERR_FILE_TO_EXPORT_PARENT_DIR_CANT_BE_CREATED, new IOException());
        assertEquals(Errors.ERR_FILE_TO_EXPORT_PARENT_DIR_CANT_BE_CREATED, e.getMessage());
        assertTrue(e.getCause() instanceof IOException);
    }

    @Test
    public void mempoiExceptionThrowableTest() {

        MempoiException e = new MempoiException(new IOException());
        assertTrue(e.getCause() instanceof IOException);
    }
}
