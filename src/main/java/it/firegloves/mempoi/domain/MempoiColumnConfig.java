/**
 * represents the configuration of a MempoiColumn
 */
package it.firegloves.mempoi.domain;

import java.util.List;
import java.util.Optional;

import it.firegloves.mempoi.domain.datatransformation.DataTransformationChain;
import it.firegloves.mempoi.domain.datatransformation.DataTransformationFunction;
import lombok.AccessLevel;
import lombok.Builder;
import lombok.Getter;
import lombok.Singular;
import lombok.Value;
import lombok.experimental.Accessors;

@Value
public class MempoiColumnConfig {

    /**
     * used to bind the current MempoiColumnConfig to the relative MempoiColumn
     */
    @Getter
    String columnName;
    DataTransformationChain<Object,?> dataTransformationChain;

    public MempoiColumnConfig(String columnName,
                              DataTransformationChain<Object, ?> dataTransformationChain) {
        this.columnName = columnName;
        this.dataTransformationChain = dataTransformationChain;
    }


    public Optional<DataTransformationChain<Object, ?>> getDataTransformationChain() {
        return Optional.ofNullable(dataTransformationChain);
    }
}
