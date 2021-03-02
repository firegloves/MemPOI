package it.firegloves.mempoi.domain;


import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.domain.MempoiColumnConfig.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.MempoiColumnConfigTestHelper;
import org.junit.Test;

public class MempoiColumnConfigTest {

    @Test
    public void testFullBuilder() {

        MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                .withColumnName(MempoiColumnConfigTestHelper.COLUMN_NAME)
                .withDataTransformationFunction(MempoiColumnConfigTestHelper.DATA_TRANFORMATION_FUNCTION)
                .build();

        assertEquals(MempoiColumnConfigTestHelper.COLUMN_NAME, mempoiColumnConfig.getColumnName());
        assertEquals(MempoiColumnConfigTestHelper.DATA_TRANFORMATION_FUNCTION,
                mempoiColumnConfig.getDataTransformationFunction().get());
    }

    @Test
    public void shouldTnrowExceptionIfNullOrEmptyColumnNameIsPassed() {

        // null column name
        try {
            MempoiColumnConfigBuilder.aMempoiColumnConfig()
                    .build();
        } catch (Exception e) {
            assertEquals(MempoiException.class, e.getClass());
            assertEquals("An invalid column name has been set for a MempoiColumnConfig", e.getMessage());
        }

        // null column name
        try {
            MempoiColumnConfigBuilder.aMempoiColumnConfig()
                    .withColumnName("")
                    .build();
        } catch (Exception e) {
            assertEquals(MempoiException.class, e.getClass());
            assertEquals("An invalid column name has been set for a MempoiColumnConfig", e.getMessage());
        }
    }
}
