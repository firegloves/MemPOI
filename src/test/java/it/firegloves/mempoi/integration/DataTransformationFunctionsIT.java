package it.firegloves.mempoi.integration;

import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.times;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.DataTransformationFunction;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.testutil.TestHelper;
import org.junit.Test;

public class DataTransformationFunctionsIT extends IntegrationBaseIT {

    @Test
    public void shouldExecuteTheReceivedDataTransformationFunctions() throws Exception {

        DataTransformationFunction<Object> dataTransformationFunctionOne = mock(DataTransformationFunction.class);
        DataTransformationFunction<Object> dataTransformationFunctionTwo = mock(DataTransformationFunction.class);
        when(dataTransformationFunctionOne.apply(any())).thenReturn(true);
        when(dataTransformationFunctionTwo.apply(any())).thenReturn(true);

        MempoiColumnConfig validColConfig = MempoiColumnConfig.builder().withColumnName("valid")
                .withDataTransformationFunction(dataTransformationFunctionOne)
                .withDataTransformationFunction(dataTransformationFunctionTwo)
                .build();

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .addMempoiColumnConfig(validColConfig)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .addMempoiSheet(mempoiSheet)
                .build();

        memPOI.prepareMempoiReportToByteArray().get();

        verify(dataTransformationFunctionOne, times(TestHelper.MAX_ROWS)).apply(any());
        verify(dataTransformationFunctionTwo, times(TestHelper.MAX_ROWS)).apply(any());
    }
}
