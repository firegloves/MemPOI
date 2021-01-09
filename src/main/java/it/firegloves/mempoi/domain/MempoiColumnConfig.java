/**
 * represents the configuration of a MempoiColumn
 */
package it.firegloves.mempoi.domain;

import java.util.List;
import java.util.Optional;
import lombok.AccessLevel;
import lombok.Builder;
import lombok.Getter;
import lombok.Singular;
import lombok.Value;
import lombok.experimental.Accessors;

@Value
@Accessors(chain = true)
@Builder(setterPrefix = "with")
public class MempoiColumnConfig {

    /**
     * used to bind the current MempoiColumnConfig to the relative MempoiColumn
     */
    String columnName;

    @Singular("dataTransformationFunction")
    @Getter(AccessLevel.NONE)
    List<DataTransformationFunction<?, ?>> dataTransformationFunctionList;


    public Optional<List<DataTransformationFunction<?, ?>>> getDataTransformationFunctionList() {
        return Optional.ofNullable(dataTransformationFunctionList);
    }
}
