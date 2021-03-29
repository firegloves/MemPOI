package it.firegloves.mempoi.integration;

import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.ConnectionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import org.junit.After;
import org.junit.Before;

public abstract class IntegrationBaseIT {

    protected File outReportFolder = new File("out/report-files/");

    // in order to run tests you need to manually import resources/test_dump.sql first
    // it's a mysql 8+ dump
    // then adjust db connection string accordingly with your parameters

    private Connection conn = null;
    protected PreparedStatement prepStmt = null;

    @Before
    public void init() {
        try {
            this.conn = ConnectionHelper.getConnection();
            this.prepStmt = this.createStatement();

            if (!this.outReportFolder.exists() && !this.outReportFolder.mkdirs()) {
                throw new MempoiException("Error in creating out report file folder: " + this.outReportFolder.getAbsolutePath() + ". Maybe permissions problem?");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    protected String createQuery(String[] columns, String[] headers, int maxLimit) {
        return createQuery(TestHelper.TABLE_EXPORT_TEST, columns, headers, maxLimit);
    }

    /**
     * creates and returns the export query
     *
     * @param columns  the String[] of the columns needed in the export query
     * @param headers  the String[] of the columns needed in the export query
     * @param maxLimit the upper bound value of the limit statement, set it to -1 if limit is not desired -> NO_LIMITS
     * @return the resulting query
     */
    protected String createQuery(String tableName, String[] columns, String[] headers, int maxLimit) {

        StringBuilder sb = new StringBuilder("SELECT ");
        for (int i = 0; i < columns.length; i++) {
            sb.append("`").append(columns[i]).append("`");
            if (i < headers.length) {
                sb.append(" AS ").append("`").append(headers[i]).append("`");
            }
            sb.append(", ");
        }
        sb.delete(sb.length() - 2, sb.length());
        sb.append(" FROM " + tableName);

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
    public PreparedStatement createStatement() throws SQLException {
        return this.conn.prepareStatement(this.createQuery(TestHelper.COLUMNS, TestHelper.HEADERS, TestHelper.MAX_ROWS));
    }
}
