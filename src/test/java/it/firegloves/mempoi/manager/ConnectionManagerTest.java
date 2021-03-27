/**
 * created by firegloves
 */

package it.firegloves.mempoi.manager;

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
    ResultSet rs;
    @Mock
    PreparedStatement prepStmt;

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
    public void closeResultSetAndPrepStmt_allOpen() throws Exception {

        when(rs.isClosed()).thenReturn(false);
        when(prepStmt.isClosed()).thenReturn(false);

        ConnectionManager.closeResultSetAndPrepStmt(rs, prepStmt);
    }

    @Test
    public void closeResultSetAndPrepStmt_rsOpenStmtClosed() throws Exception {

        when(rs.isClosed()).thenReturn(false);
        when(prepStmt.isClosed()).thenReturn(true);

        ConnectionManager.closeResultSetAndPrepStmt(rs, prepStmt);
    }

    @Test
    public void closeResultSetAndPrepStmt_rsClosedStmtOpen() throws Exception {

        when(rs.isClosed()).thenReturn(true);
        when(prepStmt.isClosed()).thenReturn(false);

        ConnectionManager.closeResultSetAndPrepStmt(rs, prepStmt);
    }

    @Test
    public void closeResultSetAndPrepStmt_rsClosedStmtClosed() throws Exception {

        when(rs.isClosed()).thenReturn(true);
        when(prepStmt.isClosed()).thenReturn(true);

        ConnectionManager.closeResultSetAndPrepStmt(rs, prepStmt);
    }

    @Test(expected = MempoiException.class)
    public void closeResultSetAndPrepStmt_errorClosing() throws Exception {

        when(rs.isClosed()).thenThrow(new SQLException());

        ConnectionManager.closeResultSetAndPrepStmt(rs, prepStmt);
    }
}
