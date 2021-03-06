/**
 * represents a column of the export report
 * it contains every info required to accomplish the export
 */
package it.firegloves.mempoi.domain;

import it.firegloves.mempoi.datapostelaboration.mempoicolumn.MempoiColumnElaborationStep;
import it.firegloves.mempoi.domain.datatransformation.DataTransformationFunction;
import it.firegloves.mempoi.domain.footer.MempoiSubFooterCell;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.MempoiColumnStyleManager;
import it.firegloves.mempoi.util.Errors;
import java.lang.reflect.Method;
import java.sql.ResultSet;
import java.sql.Types;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

@Data
@Accessors(chain = true)
public class MempoiColumn {

    /**
     * the enum mapping ResultSet actions to the column data
     */
    private EExportDataType type;
    /**
     * the style to apply to every cell of the column
     */
    private CellStyle cellStyle;
    /**
     * the name of the column. will be used as column header value
     */
    private String columnName;
    /**
     * the column index in the report
     */
    private int colIndex;
    /**
     * the method to call on the ResultSet to extract the desired data
     */
    private Method rsAccessDataMethod;
    /**
     * the setter method to invoke on the Cell to set its value using the reflection
     */
    private Method cellSetValueMethod;
    /**
     * data needed to manage the subfooter of this column
     */
    private MempoiSubFooterCell subFooterCell;
    /**
     * contains the configuration of the current MempoiColumn
     */
    private MempoiColumnConfig mempoiColumnConfig;


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
     * By default, return the column name.
     * If a mempoiColumnConfig is set and a display name is configured then, the configured displayed name is returned.
     *
     * Note: empty String is allowed in the configuration.
     *
     * @return The column header cell value to display
     */
    public String getColumnDisplayName() {
        if(Objects.nonNull(mempoiColumnConfig)) {
            return mempoiColumnConfig.getColumnDisplayName().orElse(this.columnName);
        }
        return this.columnName;
    }


    /**
     * read result set access method String from type and set the relative java.reflect.Method
     */
    private void setResultSetAccessMethod(EExportDataType type) {

        try {
            this.rsAccessDataMethod = ResultSet.class
                    .getMethod(type.getRsAccessDataMethodName(), type.getRsAccessParamClass());
        } catch (NoSuchMethodException e) {
            throw new MempoiException(e);
        }
    }

    public MempoiColumnConfig getMempoiColumnConfig() {
        return mempoiColumnConfig;
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

            case Types.BIGINT:
                return EExportDataType.BIG_INTEGER;
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
            case TypesExtended.UUID:
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


    /**
     * basing on the received class type, returns the relative EExportDataType
     */
    private Optional<EExportDataType> getEExportDataTypeByParameterValue(Class<?> cellSetValueMethodParamClass) {

        String simpleName = cellSetValueMethodParamClass.getSimpleName();

        switch (simpleName) {

            case "Double":
            case "Long":
            case "BigInteger":
                return Optional.ofNullable(EExportDataType.DOUBLE);
            case "Float":
                return Optional.ofNullable(EExportDataType.FLOAT);
            case "Integer":
            case "Short":
                return Optional.ofNullable(EExportDataType.INT);
            case "String":
            case "Character":
                return Optional.ofNullable(EExportDataType.TEXT);
            case "Time":
            case "LocalDateTime":
                return Optional.ofNullable(EExportDataType.TIME);
            case "Timestamp":
                return Optional.ofNullable(EExportDataType.TIMESTAMP);
            case "Date":
            case "LocalDate":
                return Optional.ofNullable(EExportDataType.DATE);
            case "Boolean":
                return Optional.ofNullable(EExportDataType.BOOLEAN);
            default:
                throw new MempoiException("JAVA TYPE CLASS NOT RECOGNIZED: " + simpleName);
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
     * sets and configure the required data to process the MempoiColumnConfig
     *
     * @param mempoiColumnConfig
     * @param mempoiColumnStyleManager
     * @return
     */
    public MempoiColumn setMempoiColumnConfig(MempoiColumnConfig mempoiColumnConfig, MempoiColumnStyleManager mempoiColumnStyleManager) {
        this.mempoiColumnConfig = mempoiColumnConfig;

        if (null != this.mempoiColumnConfig) {

            this.mempoiColumnConfig.getDataTransformationFunction()
                    //if user has supplied a list of transformation functions then apply configurations
                    .ifPresent(dataTransformationFunction -> {


                        // get transform method of the DataTransformationFunction selecting the one typed by the type
                        // erasure. This is meant to avoid the Object function return type.
                        Method functionMethod = Arrays
                                .stream(dataTransformationFunction.getClass().getDeclaredMethods())
                                .filter(method ->
                                        DataTransformationFunction.TRANSFORM_METHOD_NAME.equals(method.getName()) &&
                                                !method.getReturnType().getSimpleName().equals("Object"))
                                .findFirst()
                                .orElseThrow(() -> new MempoiException(
                                        Errors.ERR_DATA_TRANSFORMATION_FUNCTION_METHOD_NOT_FOUND));

                        // get the final step return data type
                        EExportDataType eExportDataType = this
                                .getEExportDataTypeByParameterValue(functionMethod.getReturnType())
                                .orElseThrow(() -> new MempoiException(
                                        Errors.ERR_DATA_TRANSFORMATION_FUNCTION_EEXPORTDATATYPE_NOT_FOUND));

                        // set the export data type to be used to set the final cell value
                        this.setCellSetValueMethod(eExportDataType);
                        // set cell style basing of the new identified EExportDataType
                        mempoiColumnStyleManager.setMempoiColumnCellStyle(this, eExportDataType);
                    });

            // TODO TEST THIS
            if (null != mempoiColumnConfig.getCellStyle()) {
                this.setCellStyle(mempoiColumnConfig.getCellStyle());
            }
        }

        return this;
    }

    /**
     * applies pipeline elaboration steps (NOT SXSSF - after workbook creation)
     *
     * @param mempoiSheet the MempoiSheet from which gain informations
     * @param workbook    the Workbook from which get Sheet
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
