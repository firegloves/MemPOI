package it.firegloves.mempoi.domain;

import it.firegloves.mempoi.datapostelaboration.mempoicolumn.MempoiColumnElaborationStep;
import it.firegloves.mempoi.domain.footer.MempoiSubFooterCell;
import it.firegloves.mempoi.exception.MempoiException;
import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Method;
import java.sql.ResultSet;
import java.sql.Types;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

@Data
@Accessors(chain = true)
public class MempoiColumn {

    private EExportDataType type;
    private CellStyle cellStyle;
    private String columnName;
    private int colIndex;

    /**
     * contains the ResultSet's method to
     */
    private Method rsAccessDataMethod;
    private Method cellSetValueMethod;

    private MempoiSubFooterCell subFooterCell;


    /**
     * pipeline pattern for elaborating MempoiColumn's data
     */
    private List<MempoiColumnElaborationStep> elaborationStepList = new ArrayList<>();


    public MempoiColumn(String columnName) {
        this.columnName = columnName;
    }

    public MempoiColumn(int sqlObjType, String columnName, int colIndex) {
        this.columnName = columnName;
        this.setType(sqlObjType);
        this.colIndex = colIndex;
    }

    public void setType(int sqlObjType) {
        this.type = this.getFieldTypeName(sqlObjType);
        this.setResultSetAccessMethod(this.type);
        this.setCellSetValueMethod(this.type);
    }

    /**
     * read result set access method String from type and set the relative java.reflect.Method
     */
    private void setResultSetAccessMethod(EExportDataType type) {

        try {
            this.rsAccessDataMethod = ResultSet.class.getMethod(type.getRsAccessDataMethodName(), type.getRsAccessParamClass());
        } catch (NoSuchMethodException e) {
            throw new MempoiException(e);
        }
    }

    /**
     * set Poi Cell set value method basing on the type return class
     */
    private void setCellSetValueMethod(EExportDataType type) {

        try {
            this.cellSetValueMethod = Cell.class.getMethod("setCellValue", type.getRsReturnClass());
        } catch (NoSuchMethodException e) {
            throw new MempoiException(e);
        }
    }


    /**
     * basing on the sqlObjType returns the relative EExportDataType
     *
     * @param sqlObjType
     */
    private EExportDataType getFieldTypeName(int sqlObjType) {

        switch (sqlObjType) {

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
            case Types.BIGINT:
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
                throw new MempoiException("SQL TYPE NOT RECOGNIZED: " + sqlObjType);
        }
    }

    public void addElaborationStep(MempoiColumnElaborationStep step) {
        if (null != step) {
            this.elaborationStepList.add(step);
        }
    }

    /**
     * applies strategy analysis
     *
     * @param cell  the Cell from which gain informations
     * @param value cell value of type T
     */
    public void elaborationStepListAnalyze(Cell cell, Object value) {
        this.elaborationStepList.forEach(step -> step.performAnalysis(cell, value));
    }

    /**
     * closes the analysis, often used to manage the point that the ResultSet was already full iterated
     *
     * @param lastRowNum last row num
     */
    public void elaborationStepListCloseAnalysis(int lastRowNum) {
        this.elaborationStepList.forEach(step -> step.closeAnalysis(lastRowNum));
    }

    /**
     * applies pipeline elaboration steps (NOT SXSSF - after workbook creation)
     *
     * @param mempoiSheet the MempoiSheet from which gain informations
     * @param workbook the Workbook from which get Sheet
     */
    public void elaborationStepListExecute(MempoiSheet mempoiSheet, Workbook workbook) {
        this.elaborationStepList.forEach(step -> step.execute(mempoiSheet, workbook));
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
    }

    @Override
    public int hashCode() {
        return Objects.hash(type, cellStyle, columnName, rsAccessDataMethod, cellSetValueMethod, subFooterCell);
    }
}