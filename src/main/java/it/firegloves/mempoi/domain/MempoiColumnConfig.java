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
    @Getter
    Optional<String> columnDisplayName;

    /**
     * DataTransformationFunction to apply to every value read by the DB for column identified by the columnName
     * property
     */
    DataTransformationFunction<?, ?> dataTransformationFunction;

    /**
     * CellStyle to apply to the column identified by the columnName property
     */
    CellStyle cellStyle;

    private MempoiColumnConfig(String columnName,
            DataTransformationFunction<?, ?> dataTransformationFunction,
            CellStyle cellStyle, String columnDisplayName) {
        this.columnName = columnName;
        this.dataTransformationFunction = dataTransformationFunction;
        this.cellStyle = cellStyle;
        this.columnDisplayName = Optional.ofNullable(columnDisplayName);
    }


    public Optional<DataTransformationFunction<?, ?>> getDataTransformationFunction() {
        return Optional.ofNullable(dataTransformationFunction);
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

        public MempoiColumnConfig build() {

            if (StringUtils.isEmpty(columnName)) {
                throw new MempoiException(Errors.ERR_MEMPOI_COL_CONFIG_COLNAME_NOT_VALID);
            }

            return new MempoiColumnConfig(columnName, dataTransformationFunction, cellStyle, columnDisplayName);
        }
    }
}
