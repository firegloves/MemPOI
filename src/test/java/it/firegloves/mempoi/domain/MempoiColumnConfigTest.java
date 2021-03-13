package it.firegloves.mempoi.domain;


import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.domain.MempoiColumnConfig.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.MempoiColumnConfigTestHelper;
import org.junit.Test;

public class MempoiColumnConfigTest {

    @Test
    public void testFullBuilder() {

        MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                .withColumnName(MempoiColumnConfigTestHelper.COLUMN_NAME)
                .withDataTransformationFunction(MempoiColumnConfigTestHelper.STRING_DATA_TRANFORMATION_FUNCTION)
                .build();

        AssertionHelper.validateMempoiColumnConfig(MempoiColumnConfigTestHelper.getTestMempoiColumnConfig(),
                mempoiColumnConfig);
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
