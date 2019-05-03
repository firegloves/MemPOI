package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.exception.MempoiException;
import org.junit.After;
import org.junit.Before;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

public class FunctionalBaseTest {

    protected File outReportFolder = new File("out/report-files/");

    // in order to run tests you need to manually import resources/mempoi_export_test.sql first
    // it's a mysql 8+ dump
    // then adjust db connection string accordingly with your parameters

    Connection conn = null;
    PreparedStatement prepStmt = null;

    @Before
    public void init() {
        try {
            this.conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/mempoi", "root", "");
            this.prepStmt = this.conn.prepareStatement("SELECT id, creation_date AS `WONDERFUL DATE`, dateTime, timeStamp, name, valid, usefulChar, decimalOne FROM mempoi.export_test LIMIT 0, 10");
//            this.prepStmt = this.conn.prepareStatement("SELECT id, creation_date AS `WONDERFUL DATE`, dateTime, timeStamp, name, valid, usefulChar, decimalOne, bitTwo, doublone, floattone, interao, mediano, attempato, interuccio " +
//                    "FROM mempoi.export_test");


            if (! this.outReportFolder.exists() && ! this.outReportFolder.mkdirs()) {
                throw new MempoiException("Error in creating out report file folder: " + this.outReportFolder.getAbsolutePath() + ". Maybe permissions problem?");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @After
    public void afterTest() {

        try {
            if (! this.conn.isClosed()) {
                this.conn.close();
            }

            // prepared statement is closed by mempoi

        } catch (SQLException e) {
            e.printStackTrace();
        }
    }
}
