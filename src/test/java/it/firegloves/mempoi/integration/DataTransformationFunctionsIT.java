package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotEquals;
import static org.junit.Assert.assertNotNull;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.datatransformation.DataTransformationChain;
import it.firegloves.mempoi.domain.datatransformation.DataTransformationFunction;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;

import java.io.File;
import java.util.concurrent.CompletableFuture;

import org.junit.Test;

public class DataTransformationFunctionsIT extends IntegrationBaseIT {

    // TODO refine this test: what is testing? for now is only a quick try to check MempoiColumnConfig is debug mode
    @Test
    public void testWithoutMempoiColumnConfig() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "data_transformation_functions.xlsx");

        try {

            DataTransformationChain<Object, Integer> dataTransformationChain = new DataTransformationChain<>(
                    (Object o) -> "ciao").chainUp((String s) -> 5);
            MempoiColumnConfig nameConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig().withColumnName("name")
                    .withDataTransformationChain(dataTransformationChain).build();
            MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                    .addMempoiColumnConfig(nameConfig).build();

            MemPOI memPOI = MempoiBuilder.aMemPOI().withFile(fileDest).addMempoiSheet(mempoiSheet).build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper
                    .validateGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS,
                            null, new StandardStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
