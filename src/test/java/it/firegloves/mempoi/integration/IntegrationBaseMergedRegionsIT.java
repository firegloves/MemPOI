package it.firegloves.mempoi.integration;

import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.ConnectionManagerHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.After;
import org.junit.Before;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;

public abstract class IntegrationBaseMergedRegionsIT extends IntegrationBaseIT {

    Connection conn = null;
    PreparedStatement prepStmt = null;

    @Before
    public void init() {
        try {
            this.conn = ConnectionManagerHelper.getConnection();

            if (!this.outReportFolder.exists() && !this.outReportFolder.mkdirs()) {
                throw new MempoiException("Error in creating out report file folder: " + this.outReportFolder.getAbsolutePath() + ". Maybe permissions problem?");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /**
     * creates and returns the export query
     *
     * @param columns  the String[] of the columns needed in the export query
     * @param headers  the String[] of the columns needed in the export query
     * @param maxLimit the upper bound value of the limit statement, set it to -1 if limit is not desired -> NO_LIMITS
     * @return the resulting query
     */
    protected String createQuery(String[] columns, String[] headers, int maxLimit, String[] orderByCols) {

        StringBuilder sb = new StringBuilder("SELECT ");
        for (int i = 0; i < columns.length; i++) {
            sb.append("`").append(columns[i]).append("`");
            if (i < headers.length) {
                sb.append(" AS ").append("`").append(headers[i]).append("`");
            }
            sb.append(", ");
        }
        sb.delete(sb.length() - 2, sb.length());
        sb.append(" FROM " + TestHelper.TABLE_MERGED_REGIONS);

        if (null != orderByCols && orderByCols.length > 0) {
            sb.append(" ORDER BY ");
            String grpCols = Arrays.stream(orderByCols).collect(Collectors.joining(", "));
            sb.append(grpCols);
        }

        if (maxLimit > -1) {
            sb.append(" LIMIT 0, " + maxLimit);
        }

        return sb.toString();
    }


    @After
    public void afterTest() {

        try {
            if (!this.conn.isClosed()) {
                this.conn.close();
            }

            // prepared statement is closed by mempoi

        } catch (SQLException e) {
            e.printStackTrace();
        }
    }


    /**
     * create a standard PreparedStatement
     *
     * @return the created PreparedStatement
     * @throws SQLException
     */
    public PreparedStatement createStatement(String[] orderByCols) throws SQLException {
        return this.conn.prepareStatement(this.createQuery(TestHelper.COLUMNS, TestHelper.HEADERS, TestHelper.MAX_ROWS, orderByCols));
    }

    /**
     * create a standard PreparedStatement
     *
     * @return the created PreparedStatement
     * @throws SQLException
     */
    public PreparedStatement createStatement(String[] orderByCols, int maxRows) throws SQLException {
        return this.conn.prepareStatement(this.createQuery(TestHelper.COLUMNS, TestHelper.HEADERS, maxRows, orderByCols));
    }

}
