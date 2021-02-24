package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.datatransformation.DataTransformationFunction;

public class MempoiColumnConfigBuilder {
    String columnName;
    DataTransformationFunction<Object, ?> dataTransformationFunction;

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
            DataTransformationFunction<Object, ?> dataTransformationFunction) {
        this.dataTransformationFunction = dataTransformationFunction;
        return this;
    }


    public MempoiColumnConfig build() {
        return new MempoiColumnConfig(columnName, dataTransformationFunction);

    }
}
