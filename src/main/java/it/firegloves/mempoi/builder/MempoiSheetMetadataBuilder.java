package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.domain.MempoiSheetMetadata;
import org.apache.poi.ss.SpreadsheetVersion;

public final class MempoiSheetMetadataBuilder {

    private SpreadsheetVersion spreadsheetVersion;
    private String sheetName;
    private int sheetIndex;
    private int totalRows;
    private int headerRowIndex;
    private int totalDataRows;
    private int firstDataRow;
    private int lastDataRow;
    private int subfooterRowIndex;
    private int totalColumns;
    private int firstDataColumn;
    private int lastDataColumn;
    private int firstTableRow;
    private int lastTableRow;
    private int firstTableColumn;
    private int lastTableColumn;
    private int firstPivotTableRow;
    private int lastPivotTableRow;
    private int firstPivotTableColumn;
    private int lastPivotTableColumn;

    private MempoiSheetMetadataBuilder() {
    }

    public static MempoiSheetMetadataBuilder aMempoiSheetMetadata() {
        return new MempoiSheetMetadataBuilder();
    }

    public MempoiSheetMetadataBuilder withSpreadsheetVersion(SpreadsheetVersion spreadsheetVersion) {
        this.spreadsheetVersion = spreadsheetVersion;
        return this;
    }

    public MempoiSheetMetadataBuilder withSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public MempoiSheetMetadataBuilder withSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
        return this;
    }

    public MempoiSheetMetadataBuilder withTotalRows(int totalRows) {
        this.totalRows = totalRows;
        return this;
    }

    public MempoiSheetMetadataBuilder withHeaderRowIndex(int headerRowIndex) {
        this.headerRowIndex = headerRowIndex;
        return this;
    }

    public MempoiSheetMetadataBuilder withTotalDataRows(int totalDataRows) {
        this.totalDataRows = totalDataRows;
        return this;
    }

    public MempoiSheetMetadataBuilder withFirstDataRow(int firstDataRow) {
        this.firstDataRow = firstDataRow;
        return this;
    }

    public MempoiSheetMetadataBuilder withLastDataRow(int lastDataRow) {
        this.lastDataRow = lastDataRow;
        return this;
    }

    public MempoiSheetMetadataBuilder withSubfooterRowIndex(int subfooterRowIndex) {
        this.subfooterRowIndex = subfooterRowIndex;
        return this;
    }

    public MempoiSheetMetadataBuilder withTotalColumns(int totalColumns) {
        this.totalColumns = totalColumns;
        return this;
    }

    public MempoiSheetMetadataBuilder withFirstDataColumn(int firstDataColumn) {
        this.firstDataColumn = firstDataColumn;
        return this;
    }

    public MempoiSheetMetadataBuilder withLastDataColumn(int lastDataColumn) {
        this.lastDataColumn = lastDataColumn;
        return this;
    }

    public MempoiSheetMetadataBuilder withFirstTableRow(int firstTableRow) {
        this.firstTableRow = firstTableRow;
        return this;
    }

    public MempoiSheetMetadataBuilder withLastTableRow(int lastTableRow) {
        this.lastTableRow = lastTableRow;
        return this;
    }

    public MempoiSheetMetadataBuilder withFirstTableColumn(int firstTableColumn) {
        this.firstTableColumn = firstTableColumn;
        return this;
    }

    public MempoiSheetMetadataBuilder withLastTableColumn(int lastTableColumn) {
        this.lastTableColumn = lastTableColumn;
        return this;
    }

    public MempoiSheetMetadataBuilder withFirstPivotTableRow(int firstPivotTableRow) {
        this.firstPivotTableRow = firstPivotTableRow;
        return this;
    }

    public MempoiSheetMetadataBuilder withLastPivotTableRow(int lastPivotTableRow) {
        this.lastPivotTableRow = lastPivotTableRow;
        return this;
    }

    public MempoiSheetMetadataBuilder withFirstPivotTableColumn(int firstPivotTableColumn) {
        this.firstPivotTableColumn = firstPivotTableColumn;
        return this;
    }

    public MempoiSheetMetadataBuilder withLastPivotTableColumn(int lastPivotTableColumn) {
        this.lastPivotTableColumn = lastPivotTableColumn;
        return this;
    }

    public MempoiSheetMetadata build() {
        MempoiSheetMetadata mempoiSheetMetadata = new MempoiSheetMetadata();
        mempoiSheetMetadata.setSpreadsheetVersion(spreadsheetVersion);
        mempoiSheetMetadata.setSheetName(sheetName);
        mempoiSheetMetadata.setSheetIndex(sheetIndex);
        mempoiSheetMetadata.setTotalRows(totalRows);
        mempoiSheetMetadata.setHeaderRowIndex(headerRowIndex);
        mempoiSheetMetadata.setTotalDataRows(totalDataRows);
        mempoiSheetMetadata.setFirstDataRow(firstDataRow);
        mempoiSheetMetadata.setLastDataRow(lastDataRow);
        mempoiSheetMetadata.setSubfooterRowIndex(subfooterRowIndex);
        mempoiSheetMetadata.setTotalColumns(totalColumns);
        mempoiSheetMetadata.setFirstDataColumn(firstDataColumn);
        mempoiSheetMetadata.setLastDataColumn(lastDataColumn);
        mempoiSheetMetadata.setFirstTableRow(firstTableRow);
        mempoiSheetMetadata.setLastTableRow(lastTableRow);
        mempoiSheetMetadata.setFirstTableColumn(firstTableColumn);
        mempoiSheetMetadata.setLastTableColumn(lastTableColumn);
        mempoiSheetMetadata.setFirstPivotTableRow(firstPivotTableRow);
        mempoiSheetMetadata.setLastPivotTableRow(lastPivotTableRow);
        mempoiSheetMetadata.setFirstPivotTableColumn(firstPivotTableColumn);
        mempoiSheetMetadata.setLastPivotTableColumn(lastPivotTableColumn);
        return mempoiSheetMetadata;
    }
}
