package it.firegloves.mempoi.unit;

import it.firegloves.mempoi.dao.impl.DBMempoiDAO;
import it.firegloves.mempoi.domain.MempoiColumn;
import org.hamcrest.core.IsInstanceOf;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Types;
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

        MempoiColumn id = new MempoiColumn(Types.INTEGER, "id");
        MempoiColumn name = new MempoiColumn(Types.VARCHAR, "name");

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
        assertEquals("DBMempoiDAO executing readMetadata column 2", name, columnList.get(1));
    }
}
