/**
 * created by firegloves
 */

package it.firegloves.mempoi.manager;

import static org.mockito.Mockito.times;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

import it.firegloves.mempoi.exception.MempoiException;
import java.lang.reflect.Constructor;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

public class ConnectionManagerTest {

    @Mock
    private ResultSet rs;
    @Mock
    private PreparedStatement prepStmt;

    @Before
    public void setup() {

        MockitoAnnotations.initMocks(this);
    }

    /******************************************************************************************************************
     * private constructor
     *****************************************************************************************************************/

    @Test(expected = IllegalStateException.class)
    public void privateConstructorTest() throws Throwable {

        Constructor c = ConnectionManager.class.getDeclaredConstructor();
        c.setAccessible(true);

        try {
            c.newInstance();
        } catch (Exception e) {
            throw e.getCause();
        }
    }


    /******************************************************************************************************************
     * closeResultSetAndPrepStmt
     *****************************************************************************************************************/

    @Test
    public void closeResultSetAndPrepStmtAllOpen() throws Exception {

        when(rs.isClosed()).thenReturn(false);
        when(prepStmt.isClosed()).thenReturn(false);

        ConnectionManager.closeResultSetAndPrepStmt(rs, prepStmt);

        verify(rs, times(1)).close();
        verify(prepStmt, times(1)).close();
    }

    @Test
    public void closeResultSetAndPrepStmtRsOpenStmtClosed() throws Exception {

        when(rs.isClosed()).thenReturn(false);
        when(prepStmt.isClosed()).thenReturn(true);

        ConnectionManager.closeResultSetAndPrepStmt(rs, prepStmt);

        verify(rs, times(1)).close();
        verify(prepStmt, times(0)).close();
    }

    @Test
    public void closeResultSetAndPrepStmtRsClosedStmtOpen() throws Exception {

        when(rs.isClosed()).thenReturn(true);
        when(prepStmt.isClosed()).thenReturn(false);

        ConnectionManager.closeResultSetAndPrepStmt(rs, prepStmt);

        verify(rs, times(0)).close();
        verify(prepStmt, times(1)).close();
    }

    @Test
    public void closeResultSetAndPrepStmtRsClosedStmtClosed() throws Exception {

        when(rs.isClosed()).thenReturn(true);
        when(prepStmt.isClosed()).thenReturn(true);

        ConnectionManager.closeResultSetAndPrepStmt(rs, prepStmt);

        verify(rs, times(0)).close();
        verify(prepStmt, times(0)).close();
    }

    @Test(expected = MempoiException.class)
    public void closeResultSetAndPrepStmtErrorClosing() throws Exception {

        when(rs.isClosed()).thenThrow(new SQLException());

        ConnectionManager.closeResultSetAndPrepStmt(rs, prepStmt);

        verify(rs, times(1)).close();
        verify(prepStmt, times(0)).close();
    }
}
