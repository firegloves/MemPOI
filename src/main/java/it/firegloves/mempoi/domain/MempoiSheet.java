package it.firegloves.mempoi.domain;

import java.sql.PreparedStatement;

public class MempoiSheet {

    /**
     * the prepared statement to execute to take export data
     */
    private PreparedStatement prepStmt;

    /**
     * the sheet name
     */
    private String sheetName;

    public MempoiSheet(PreparedStatement prepStmt) {
        this.prepStmt = prepStmt;
    }

    public MempoiSheet(PreparedStatement prepStmt, String sheetName) {
        this.prepStmt = prepStmt;
        this.sheetName = sheetName;
    }

    public PreparedStatement getPrepStmt() {
        return prepStmt;
    }

    public void setPrepStmt(PreparedStatement prepStmt) {
        this.prepStmt = prepStmt;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }
}
