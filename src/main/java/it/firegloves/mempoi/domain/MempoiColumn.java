package it.firegloves.mempoi.domain;

import it.firegloves.mempoi.domain.footer.MempoiSubFooterCell;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import it.firegloves.mempoi.pipeline.mempoicolumn.MempoiColumnElaborationStep;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Method;
import java.sql.ResultSet;
import java.sql.Types;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class MempoiColumn {

    private EExportDataType type;
    private CellStyle cellStyle;
    private String columnName;

    private Method rsAccessDataMethod;
    private Method cellSetValueMethod;

    private MempoiSubFooterCell subFooterCell;


    /**
     * pipeline pattern for elaborating MempoiColumn's generated data after workbook generation
     */
    private List<MempoiColumnElaborationStep> elaborationStepList = new ArrayList<>();


    public MempoiColumn(String columnName) {
        this.columnName = columnName;
    }

    public MempoiColumn(int sqlObjType, String columnName) {
        this.columnName = columnName;
        this.setType(sqlObjType);
    }


    public EExportDataType getType() {
        return type;
    }

    public void setType(EExportDataType type) {
        this.type = type;
    }

    public void setType(int sqlObjType) {
        this.type = this.getFieldTypeName(sqlObjType);
        this.setResultSetAccessMethod();
        this.setCellSetValueMethod();
    }


    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public MempoiSubFooterCell getSubFooterCell() {
        return subFooterCell;
    }

    public void setSubFooterCell(MempoiSubFooterCell subFooterCell) {
        this.subFooterCell = subFooterCell;
    }

    /**
     * read result set access method String from type and set the relative java.reflect.Method
     */
    private void setResultSetAccessMethod() {

        try {
            this.rsAccessDataMethod = ResultSet.class.getMethod(this.type.getRsAccessDataMethodName(), this.type.getRsAccessParamClass());
        } catch (NoSuchMethodException e) {
            throw new MempoiRuntimeException(e);
        }
    }

    /**
     * set Poi Cell set value method basing on the type return class
     */
    private void setCellSetValueMethod() {

        try {
            this.cellSetValueMethod = Cell.class.getMethod("setCellValue", this.type.getRsReturnClass());
        } catch (NoSuchMethodException e) {
            throw new MempoiRuntimeException(e);
        }
    }


    /**
     * basing on the sqlObjType returns the relative EExportDataType
     *
     * @param sqlObjType
     */
    private EExportDataType getFieldTypeName(int sqlObjType) {

        switch (sqlObjType) {

            case Types.BIGINT:
            case Types.DOUBLE:
                return EExportDataType.DOUBLE;
            case Types.DECIMAL:
            case Types.FLOAT:
            case Types.NUMERIC:
            case Types.REAL:
                return EExportDataType.FLOAT;
            case Types.INTEGER:
            case Types.SMALLINT:
            case Types.TINYINT:
                return EExportDataType.INT;
            case Types.CHAR:
            case Types.NCHAR:
            case Types.VARCHAR:
            case Types.NVARCHAR:
            case Types.LONGVARCHAR:
                return EExportDataType.TEXT;
            case Types.TIMESTAMP:
                return EExportDataType.TIMESTAMP;
            case Types.DATE:
                return EExportDataType.DATE;
            case Types.TIME:
                return EExportDataType.TIME;
            case Types.BIT:
            case Types.BOOLEAN:
                return EExportDataType.BOOLEAN;
            default:
                throw new MempoiRuntimeException("SQL TYPE NOT RECOGNIZED: " + sqlObjType);
        }
    }

    public Method getRsAccessDataMethod() {
        return rsAccessDataMethod;
    }

    public Method getCellSetValueMethod() {
        return cellSetValueMethod;
    }

    public void setCellSetValueMethod(Method cellSetValueMethod) {
        this.cellSetValueMethod = cellSetValueMethod;
    }

    public void addElaborationStep(MempoiColumnElaborationStep step) {
        this.elaborationStepList.add(step);
    }


    /**
     * applies strategy analysis
     *
     * @param cell  the Cell from which gain informations
     * @param value cell value of type T
     */
    public void elaborationStepListAnalyze(Cell cell, Object value) {
        this.elaborationStepList.stream().forEach(step -> step.analyze(cell, value));
    }

    /**
     * closes the analysis, often used to manage the point that the ResultSet was already full iterated
     *
     * @param lastRowNum last row num
     */
    public void elaborationStepListCloseAnalysis(int lastRowNum) {
        this.elaborationStepList.stream().forEach(step -> step.closeAnalysis(lastRowNum));
    }

    /**
     * applies strategy analysis
     *
     * @param mempoiSheet the MempoiSheet from which gain informations
     * @param workbook the Workbook from which get Sheet
     */
    public void elaborationStepListExecute(MempoiSheet mempoiSheet, Workbook workbook) {
        this.elaborationStepList.stream().forEach(step -> step.execute(mempoiSheet, workbook));
    }


    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (obj == null || getClass() != obj.getClass()) {
            return false;
        }
        final MempoiColumn other = (MempoiColumn) obj;
        return Objects.equals(this.columnName, other.columnName);

        // TODO check if the commented block was used

//                && Objects.equals(this.cellStyle, other.cellStyle)
//                && Objects.equals(this.type, other.type)
//                && Objects.equals(this.rsAccessDataMethod, other.rsAccessDataMethod)
//                && Objects.equals(this.cellSetValueMethod, other.cellSetValueMethod)
//                && Objects.equals(this.subFooterCell, other.subFooterCell);
    }

    @Override
    public int hashCode() {
        return Objects.hash(type, cellStyle, columnName, rsAccessDataMethod, cellSetValueMethod, subFooterCell);
    }
}
