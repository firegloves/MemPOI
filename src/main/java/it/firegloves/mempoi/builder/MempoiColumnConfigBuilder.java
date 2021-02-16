package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.datatransformation.DataTransformationChain;
import it.firegloves.mempoi.domain.datatransformation.DataTransformationFunction;

import java.util.List;

public class MempoiColumnConfigBuilder {
    String columnName;
    DataTransformationChain<Object, ?> dataTransformationChain;

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

    public MempoiColumnConfigBuilder withDataTransformationChain(
            DataTransformationChain<Object, ?> dataTransformationChain) {
        this.dataTransformationChain = dataTransformationChain;
        return this;
    }


    public MempoiColumnConfig build() {
        return new MempoiColumnConfig(columnName, dataTransformationChain);

    }
}
