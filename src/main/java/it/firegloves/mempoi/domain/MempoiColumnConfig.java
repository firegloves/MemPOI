/**
 * represents the configuration of a MempoiColumn
 */
package it.firegloves.mempoi.domain;

import lombok.Builder;
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
}
