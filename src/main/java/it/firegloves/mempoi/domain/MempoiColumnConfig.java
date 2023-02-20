/**
 * represents the configuration of a MempoiColumn
 */
package it.firegloves.mempoi.domain;

import it.firegloves.mempoi.domain.datatransformation.DataTransformationFunction;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import java.util.Optional;
import lombok.Getter;
import lombok.Value;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@Value
public class MempoiColumnConfig {

    /**
     * used to bind the current MempoiColumnConfig to the relative MempoiColumn
     */
    @Getter
    String columnName;

    /**
     * used to bind the current MempoiColumnConfig to the relative MempoiColumn
     */
    String columnDisplayName;


    /**
     * DataTransformationFunction to apply to every value read by the DB for column identified by the columnName
     * property
     */
    DataTransformationFunction<?, ?> dataTransformationFunction;

    /**
     * CellStyle to apply to the column identified by the columnName property
     */
    CellStyle cellStyle;

    /**
     * if true the corresponding column will be ignored during the export process
     */
    boolean ignoreColumn;

    /**
     * specifies the order in which the column appear in the export
     */
    Integer positionOrder;

    private MempoiColumnConfig(String columnName,
            DataTransformationFunction<?, ?> dataTransformationFunction,
            CellStyle cellStyle, String columnDisplayName,
            boolean ignoreColumn,
            Integer positionOrder) {

        this.columnName = columnName;
        this.dataTransformationFunction = dataTransformationFunction;
        this.cellStyle = cellStyle;
        this.columnDisplayName = columnDisplayName;
        this.ignoreColumn = ignoreColumn;
        this.positionOrder = positionOrder;
    }


    public Optional<DataTransformationFunction<?, ?>> getDataTransformationFunction() {
        return Optional.ofNullable(dataTransformationFunction);
    }

    public Optional<String> getColumnDisplayName() {
        return Optional.ofNullable(columnDisplayName);
    }

    /*
     * BUILDER
     */
    public static class MempoiColumnConfigBuilder {

        private static final Logger logger = LoggerFactory.getLogger(MempoiColumnConfigBuilder.class);
        private static final String OVERRIDING_MEMPOI_TABLE_COLUMN_NAME = "A previously setted value of Excel Table column {} is about to be replaced";
        private String columnName;
        private DataTransformationFunction<?, ?> dataTransformationFunction;
        private CellStyle cellStyle;
        private String columnDisplayName;
        private boolean ignoreColumn;
        private Integer positionOrder;

        /**
         * private constructor to lower constructor visibility from outside forcing the use of the static Builder pattern
         */
        private MempoiColumnConfigBuilder() {
        }

        /**
         * static method to create a new MempoiBuilder containing the received
         *
         * @return the current MempoiColumnConfigBuilder
         */
        public static MempoiColumnConfigBuilder aMempoiColumnConfig() {
            return new MempoiColumnConfigBuilder();
        }

        public MempoiColumnConfigBuilder withColumnName(String columnName) {
            this.columnName = columnName;
            return this;
        }

        public MempoiColumnConfigBuilder withDataTransformationFunction(
                DataTransformationFunction<?, ?> dataTransformationFunction) {
            this.dataTransformationFunction = dataTransformationFunction;
            this.dataTransformationFunction.setColumnName(this.columnName);
            return this;
        }

        public MempoiColumnConfigBuilder withCellStyle(CellStyle cellStyle) {
            this.cellStyle = cellStyle;
            return this;
        }

        /**
         * Customize the column title
         *
         * @param columnDisplayName the column header text to export
         * @return the current MempoiColumnConfigBuilder
         */
        public MempoiColumnConfigBuilder withColumnDisplayName(String columnDisplayName) {
            if(this.columnDisplayName != null)
                logger.info(OVERRIDING_MEMPOI_TABLE_COLUMN_NAME, this.columnName);
            this.columnDisplayName = columnDisplayName;
            return this;
        }

        /**
         * if true the current column will be ignored during the export
         *
         * @param ignoreColumn the boolean to decide if ignore the current column or not
         * @return the current MempoiColumnConfigBuilder
         */
        public MempoiColumnConfigBuilder withIgnoreColumn(boolean ignoreColumn) {
            this.ignoreColumn = ignoreColumn;
            return this;
        }

        /**
         * set the order in which the column will appear in the resulting export
         *
         * @param positionOrder the order in which the column will appear in the resulting export
         * @return the current MempoiColumnConfigBuilder
         */
        public MempoiColumnConfigBuilder withPositionOrder(int positionOrder) {
            if (positionOrder < 0) {
                throw new MempoiException(
                        String.format(Errors.ERR_MEMPOI_COL_CONFIG_POSITION_ORDER_NOT_VALID, positionOrder));
            }
            this.positionOrder = positionOrder;
            return this;
        }

        public MempoiColumnConfig build() {

            if (StringUtils.isEmpty(columnName)) {
                throw new MempoiException(Errors.ERR_MEMPOI_COL_CONFIG_COLNAME_NOT_VALID);
            }

            return new MempoiColumnConfig(columnName, dataTransformationFunction, cellStyle, columnDisplayName,
                    ignoreColumn, positionOrder);
        }
    }
}
