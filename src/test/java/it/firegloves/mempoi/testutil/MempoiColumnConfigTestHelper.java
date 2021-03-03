package it.firegloves.mempoi.testutil;

import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiColumnConfig.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.domain.datatransformation.DataTransformationFunction;
import it.firegloves.mempoi.exception.MempoiException;

public class MempoiColumnConfigTestHelper {

    public static final String COLUMN_NAME = "name";
    public static final int CELL_VALUE = 5;
    public static final DataTransformationFunction<Integer> DATA_TRANFORMATION_FUNCTION = new DataTransformationFunction<Integer>() {
        @Override
        public Integer transform(Object value) throws MempoiException {
            return CELL_VALUE;
        }
    };


    public static MempoiColumnConfig getTestMempoiColumnConfig() {
        return MempoiColumnConfigBuilder.aMempoiColumnConfig()
                .withColumnName(COLUMN_NAME)
                .withDataTransformationFunction(DATA_TRANFORMATION_FUNCTION)
                .build();
    }
}
