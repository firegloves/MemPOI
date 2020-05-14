package it.firegloves.mempoi.testutil;

import org.jeasy.props.annotations.Property;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

import static org.jeasy.props.PropertiesInjectorBuilder.aNewPropertiesInjector;

public class ConnectionManagerHelper {

    private static final String propFilename = "application.properties";

    @Property(source = propFilename, key = "db.driver")
    private static String dbDriver;
    @Property(source = propFilename, key = "db.host")
    private static String dbHost;
    @Property(source = propFilename, key = "db.port")
    private static String dbPort;
    @Property(source = propFilename, key = "db.name")
    private static String dbName;
    @Property(source = propFilename, key = "db.user")
    private static String dbUser;
    @Property(source = propFilename, key = "db.password")
    private static String dbPassword;

    public static ConnectionManagerHelper INSTANCE = new ConnectionManagerHelper();


    private ConnectionManagerHelper() {
        aNewPropertiesInjector().injectProperties(this);
    }


    public static Connection getConnection() throws SQLException {
        return DriverManager.getConnection("jdbc:" + dbDriver + "://" + dbHost + ":" + dbPort + "/" + dbName, dbUser, dbPassword);
    }
}
