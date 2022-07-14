package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.domain.MempoiSheetMetadata;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

public final class MempoiSheetMetadataBuilder {

    private SpreadsheetVersion spreadsheetVersion;
    private String sheetName;
    private Integer sheetIndex;
    private Integer simpleTextHeaderRowIndex;
    private Integer simpleTextFooterRowIndex;
    private Integer colsHeaderRowIndex;
    private Integer totalRows;
    private Integer firstDataRow;
    private Integer lastDataRow;
    private Integer subfooterRowIndex;
    private Integer totalColumns;
    private Integer colsOffset;
    private Integer rowsOffset;
    private Integer firstDataColumn;
    private Integer lastDataColumn;
    private Integer firstTableRow;
    private Integer lastTableRow;
    private Integer firstTableColumn;
    private Integer lastTableColumn;
    private Integer firstPivotTablePositionRow;
    private Integer firstPivotTablePositionColumn;
    private Integer firstPivotTableSourceRow;
    private Integer lastPivotTableSourceRow;
    private Integer firstPivotTableSourceColumn;
    private Integer lastPivotTableSourceColumn;

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

    public MempoiSheetMetadataBuilder withSheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
        return this;
    }

    public MempoiSheetMetadataBuilder withSimpleTextHeaderRowIndex(Integer simpleTextHeaderRowIndex) {
        this.simpleTextHeaderRowIndex = simpleTextHeaderRowIndex;
        return this;
    }

    public MempoiSheetMetadataBuilder withSimpleTextFooterRowIndex(Integer simpleTextFooterRowIndex) {
        this.simpleTextFooterRowIndex = simpleTextFooterRowIndex;
        return this;
    }

    public MempoiSheetMetadataBuilder withColsHeaderRowIndex(Integer colsHeaderRowIndex) {
        this.colsHeaderRowIndex = colsHeaderRowIndex;
        return this;
    }

    public MempoiSheetMetadataBuilder withTotalRows(Integer totalRows) {
        this.totalRows = totalRows;
        return this;
    }

    public MempoiSheetMetadataBuilder withFirstDataRow(Integer firstDataRow) {
        this.firstDataRow = firstDataRow;
        return this;
    }

    public MempoiSheetMetadataBuilder withLastDataRow(Integer lastDataRow) {
        this.lastDataRow = lastDataRow;
        return this;
    }

    public MempoiSheetMetadataBuilder withSubfooterRowIndex(Integer subfooterRowIndex) {
        this.subfooterRowIndex = subfooterRowIndex;
        return this;
    }

    public MempoiSheetMetadataBuilder withTotalColumns(Integer totalColumns) {
        this.totalColumns = totalColumns;
        return this;
    }

    public MempoiSheetMetadataBuilder withColsOffset(Integer colsOffset) {
        this.colsOffset = colsOffset;
        return this;
    }

    public MempoiSheetMetadataBuilder withRowsOffset(Integer rowsOffset) {
        this.rowsOffset = rowsOffset;
        return this;
    }

    public MempoiSheetMetadataBuilder withFirstDataColumn(Integer firstDataColumn) {
        this.firstDataColumn = firstDataColumn;
        return this;
    }

    public MempoiSheetMetadataBuilder withLastDataColumn(Integer lastDataColumn) {
        this.lastDataColumn = lastDataColumn;
        return this;
    }

    public MempoiSheetMetadataBuilder withTableFromAreaReference(AreaReference areaReference) {
        this.firstTableRow = areaReference.getFirstCell().getRow();
        this.firstTableColumn = (int) areaReference.getFirstCell().getCol();
        this.lastTableRow = areaReference.getLastCell().getRow();
        this.lastTableColumn = (int) areaReference.getLastCell().getCol();
        return this;
    }


    public MempoiSheetMetadataBuilder withPivotTableInfo(AreaReference source, CellReference position) {
        return this.withPivotTableSourceFromAreaReference(source)
            .withFirstPivotTablePositionRow(position.getRow())
            .withFirstPivotTablePositionColumn((int) position.getCol());

    }

    public MempoiSheetMetadataBuilder withFirstPivotTablePositionRow(Integer firstPivotTablePositionRow) {
        this.firstPivotTablePositionRow = firstPivotTablePositionRow;
        return this;
    }

    public MempoiSheetMetadataBuilder withFirstPivotTablePositionColumn(Integer firstPivotTablePositionColumn) {
        this.firstPivotTablePositionColumn = firstPivotTablePositionColumn;
        return this;
    }

    public MempoiSheetMetadataBuilder withPivotTableSourceFromAreaReference(AreaReference areaReference) {
        this.firstPivotTableSourceRow = areaReference.getFirstCell().getRow();
        this.firstPivotTableSourceColumn = (int) areaReference.getFirstCell().getCol();
        this.lastPivotTableSourceRow = areaReference.getLastCell().getRow();
        this.lastPivotTableSourceColumn = (int) areaReference.getLastCell().getCol();
        return this;
    }

    public MempoiSheetMetadata build() {
        MempoiSheetMetadata mempoiSheetMetadata = new MempoiSheetMetadata();
        mempoiSheetMetadata.setSpreadsheetVersion(spreadsheetVersion);
        mempoiSheetMetadata.setSheetName(sheetName);
        mempoiSheetMetadata.setSheetIndex(sheetIndex);
        mempoiSheetMetadata.setTotalRows(totalRows
                - (simpleTextHeaderRowIndex != null ? simpleTextHeaderRowIndex : colsHeaderRowIndex)
                + (subfooterRowIndex != null ? 1 : 0));
        mempoiSheetMetadata.setSimpleTextHeaderRowIndex(simpleTextHeaderRowIndex);
        mempoiSheetMetadata.setSimpleTextFooterRowIndex(simpleTextFooterRowIndex);
        mempoiSheetMetadata.setHeaderRowIndex(colsHeaderRowIndex);
        mempoiSheetMetadata.setTotalDataRows(lastDataRow - colsHeaderRowIndex);
        mempoiSheetMetadata.setFirstDataRow(firstDataRow);
        mempoiSheetMetadata.setLastDataRow(lastDataRow);
        mempoiSheetMetadata.setSubfooterRowIndex(subfooterRowIndex);
        mempoiSheetMetadata.setTotalColumns(totalColumns);
        mempoiSheetMetadata.setColsOffset(colsOffset);
        mempoiSheetMetadata.setRowsOffset(rowsOffset);
        mempoiSheetMetadata.setFirstDataColumn(firstDataColumn);
        mempoiSheetMetadata.setLastDataColumn(lastDataColumn);
        mempoiSheetMetadata.setFirstTableRow(firstTableRow);
        mempoiSheetMetadata.setLastTableRow(lastTableRow);
        mempoiSheetMetadata.setFirstTableColumn(firstTableColumn);
        mempoiSheetMetadata.setLastTableColumn(lastTableColumn);
        mempoiSheetMetadata.setFirstPivotTableSourceRow(firstPivotTableSourceRow);
        mempoiSheetMetadata.setLastPivotTableSourceRow(lastPivotTableSourceRow);
        mempoiSheetMetadata.setFirstPivotTableSourceColumn(firstPivotTableSourceColumn);
        mempoiSheetMetadata.setLastPivotTableSourceColumn(lastPivotTableSourceColumn);
        mempoiSheetMetadata.setFirstPivotTablePositionRow(firstPivotTablePositionRow);
        mempoiSheetMetadata.setFirstPivotTablePositionColumn(firstPivotTablePositionColumn);
        return mempoiSheetMetadata;
    }
}
