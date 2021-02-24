/**
 * represents the configuration of a MempoiColumn
 */
package it.firegloves.mempoi.domain;

import it.firegloves.mempoi.domain.datatransformation.DataTransformationFunction;
import lombok.Getter;
import lombok.Value;

import java.util.Optional;

@Value
public class MempoiColumnConfig {

    /**
     * used to bind the current MempoiColumnConfig to the relative MempoiColumn
     */
    @Getter
    String columnName;
    DataTransformationFunction<Object,?> dataTransformationFunction;

    public MempoiColumnConfig(String columnName,
                              DataTransformationFunction<Object, ?> dataTransformationChain) {
        this.columnName = columnName;
        this.dataTransformationFunction = dataTransformationChain;
    }


    public Optional<DataTransformationFunction<Object, ?>> getDataTransformationFunction() {
        return Optional.ofNullable(dataTransformationFunction);
    }
}
