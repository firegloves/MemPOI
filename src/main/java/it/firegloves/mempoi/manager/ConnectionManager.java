package it.firegloves.mempoi.manager;

import it.firegloves.mempoi.exception.MempoiException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

public class ConnectionManager {

    /**
     * private constructor to avoid class instantiation
     */
    private ConnectionManager() {
        throw new IllegalStateException("Utility class");
    }


    /**
     * closes connection resultset and preparedstatement
     * @param resultSet the ResultSet to close
     * @param pStmt the PreparedStatement to close
     */
    public static void closeResultSetAndPrepStmt(ResultSet resultSet, PreparedStatement pStmt) {

        try {

            // result set
            if (null != resultSet && !resultSet.isClosed()) {
                resultSet.close();
            }

            // prep stmt
            if (null != pStmt && !pStmt.isClosed()) {
                pStmt.close();
            }

        } catch (SQLException e) {
            throw new MempoiException(e);
        }
    }
}
