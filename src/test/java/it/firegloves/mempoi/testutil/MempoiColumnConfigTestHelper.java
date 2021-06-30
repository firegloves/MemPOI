package it.firegloves.mempoi.testutil;

import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiColumnConfig.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.domain.datatransformation.StringDataTransformationFunction;
import it.firegloves.mempoi.exception.MempoiException;

import java.sql.ResultSet;

public class MempoiColumnConfigTestHelper {

    public static final String COLUMN_NAME = "name";
    public static final int CELL_VALUE = 5;
    public static final StringDataTransformationFunction<Integer> STRING_DATA_TRANFORMATION_FUNCTION = new StringDataTransformationFunction<Integer>() {
        @Override
        public Integer transform(final ResultSet rs, String value) throws MempoiException {
            return CELL_VALUE;
        }
    };


    public static MempoiColumnConfig getTestMempoiColumnConfig() {
        return MempoiColumnConfigBuilder.aMempoiColumnConfig()
                .withColumnName(COLUMN_NAME)
                .withDataTransformationFunction(STRING_DATA_TRANFORMATION_FUNCTION)
                .build();
    }
}
