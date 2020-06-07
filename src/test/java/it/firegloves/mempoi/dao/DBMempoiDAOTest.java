package it.firegloves.mempoi.dao;

import it.firegloves.mempoi.dao.impl.DBMempoiDAO;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.exception.MempoiException;
import org.hamcrest.core.IsInstanceOf;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.sql.*;
import java.util.List;

import static org.junit.Assert.*;
import static org.mockito.Mockito.when;

public class DBMempoiDAOTest {

    @Mock
    private PreparedStatement prepStmt;
    @Mock
    private ResultSet resultSet;
    @Mock
    private ResultSetMetaData resultSetMetaData;

    @Before
    public void beforeTest() {
        MockitoAnnotations.initMocks(this);
    }


    @Test
    public void givenAPrepStmtReturnAResultSet() {

        try {
            when(prepStmt.executeQuery()).thenReturn(resultSet);
            when(prepStmt.toString()).thenReturn("SELECT * FROM Test");
        } catch (Exception e) {
            e.printStackTrace();
        }

        resultSet = DBMempoiDAO.getInstance().executeExportQuery(prepStmt);

        assertNotNull("DBMempoiDAO executing prepstmt not null", resultSet);
        assertThat("DBMempoiDAO executing prepstmt instance of ResultSet", resultSet, new IsInstanceOf(ResultSet.class));
    }


    @Test
    public void givenAResultSetReadMetadata() {

        MempoiColumn id = new MempoiColumn(Types.INTEGER, "id", 0);
        MempoiColumn name = new MempoiColumn(Types.VARCHAR, "name", 1);

        try {
            when(resultSet.getMetaData()).thenReturn(resultSetMetaData);
            when(resultSetMetaData.getColumnCount()).thenReturn(2);
            when(resultSetMetaData.getColumnType(1)).thenReturn(Types.INTEGER);
            when(resultSetMetaData.getColumnLabel(1)).thenReturn(id.getColumnName());
            when(resultSetMetaData.getColumnType(2)).thenReturn(Types.VARCHAR);
            when(resultSetMetaData.getColumnLabel(2)).thenReturn(name.getColumnName());
        } catch (Exception e) {
            e.printStackTrace();
        }

        List<MempoiColumn> columnList = DBMempoiDAO.getInstance().readMetadata(resultSet);

        assertNotNull("DBMempoiDAO executing readMetadata list not null", columnList);
        assertEquals("DBMempoiDAO executing readMetadata list size == 2", 2, columnList.size() );
        assertEquals("DBMempoiDAO executing readMetadata column 1", id, columnList.get(0));
        assertEquals("DBMempoiDAO index column 1", 0, columnList.get(0).getColIndex());
        assertEquals("DBMempoiDAO executing readMetadata column 2", name, columnList.get(1));
        assertEquals("DBMempoiDAO index column 2", 1, columnList.get(1).getColIndex());
    }

    @Test(expected = MempoiException.class)
    public void givenANullResultSetReadMetadata() {

        DBMempoiDAO.getInstance().readMetadata(null);
    }

    @Test(expected = MempoiException.class)
    public void givenAResultSetReadMetadataAndThrowException() throws SQLException {

        when(resultSet.getMetaData()).thenThrow(new SQLException());

        DBMempoiDAO.getInstance().readMetadata(this.resultSet);
    }

    @Test(expected = MempoiException.class)
    public void invalidPrepStmt() throws Exception {

        when(this.prepStmt.executeQuery()).thenThrow(new SQLException());

        DBMempoiDAO.getInstance().executeExportQuery(this.prepStmt);
    }
}
