package it.firegloves.mempoi.domain;

import org.apache.poi.ss.usermodel.Workbook;

public class MempoiTable {

    private Workbook workbook;
    private String areaReference;
    private String tableName;
    private String displayTableName;

    public Workbook getWorkbook() {
        return workbook;
    }

    public MempoiTable setWorkbook(Workbook workbook) {
        this.workbook = workbook;
        return this;
    }

    public String getAreaReference() {
        return areaReference;
    }

    public MempoiTable setAreaReference(String areaReference) {
        this.areaReference = areaReference;
        return this;
    }

    public String getTableName() {
        return tableName;
    }

    public MempoiTable setTableName(String tableName) {
        this.tableName = tableName;
        return this;
    }

    public String getDisplayTableName() {
        return displayTableName;
    }

    public MempoiTable setDisplayTableName(String displayTableName) {
        this.displayTableName = displayTableName;
        return this;
    }
}
