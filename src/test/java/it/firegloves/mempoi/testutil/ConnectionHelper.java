package it.firegloves.mempoi.testutil;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class ConnectionHelper {

    private static final String dbDriver = "mysql";
    private static final String dbHost = "localhost";
    private static final String dbPort = "3309";
    private static final String dbName = "mempoi";
    private static final String dbUser = "root";
    private static final String dbPassword = "mempassword";

    private ConnectionHelper() {}

    public static Connection getConnection() throws SQLException {
        return DriverManager.getConnection("jdbc:" + dbDriver + "://" + dbHost + ":" + dbPort + "/" + dbName, dbUser, dbPassword);
    }
}
