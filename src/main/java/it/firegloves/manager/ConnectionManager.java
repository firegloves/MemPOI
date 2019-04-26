package it.firegloves.manager;

import it.firegloves.exception.MempoiException;

import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

public class ConnectionManager {

    /**
     * closes connection resultset and preparedstatement
     * @param resultSet
     * @param pStmt
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
