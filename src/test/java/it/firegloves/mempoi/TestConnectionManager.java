package it.firegloves.mempoi;

import org.h2.jdbcx.JdbcDataSource;

import javax.sql.DataSource;
import java.sql.Connection;
import java.sql.PreparedStatement;

public class TestConnectionManager {

    private JdbcDataSource ds;
    private Connection connection;

    public TestConnectionManager() {
        ds = new JdbcDataSource();
        ds.setURL("jdbc:h2:mem:mempoi:_test;DB_CLOSE_DELAY=-1;MVCC=true");
        ds.setUser("sa");

        connection = ds.getConnection();
    }


    public PreparedStatement createPreparedStatement(String query) {
        return connection.prepareStatement(query);
    }
}
