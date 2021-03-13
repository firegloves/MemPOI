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

@Value
public class MempoiColumnConfig {

    /**
     * used to bind the current MempoiColumnConfig to the relative MempoiColumn
     */
    @Getter
    String columnName;
    DataTransformationFunction<?> dataTransformationFunction;

    private MempoiColumnConfig(String columnName,
                              DataTransformationFunction<?> dataTransformationFunction) {
        this.columnName = columnName;
        this.dataTransformationFunction = dataTransformationFunction;
    }


    public Optional<DataTransformationFunction<?>> getDataTransformationFunction() {
        return Optional.ofNullable(dataTransformationFunction);
    }


    /*
     * BUILDER
     */
    public static class MempoiColumnConfigBuilder {

        String columnName;
        DataTransformationFunction<?> dataTransformationFunction;

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
                DataTransformationFunction<?> dataTransformationFunction) {
            this.dataTransformationFunction = dataTransformationFunction;
            this.dataTransformationFunction.setColumnName(this.columnName);
            return this;
        }


        public MempoiColumnConfig build() {

            if (StringUtils.isEmpty(columnName)) {
                throw new MempoiException(Errors.ERR_MEMPOI_COL_CONFIG_COLNAME_NOT_VALID);
            }

            return new MempoiColumnConfig(columnName, dataTransformationFunction);
        }
    }
}
