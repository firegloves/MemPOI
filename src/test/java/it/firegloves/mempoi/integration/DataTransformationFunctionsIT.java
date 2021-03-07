package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiColumnConfig.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.MempoiColumnConfigTestHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.util.concurrent.CompletableFuture;
import org.junit.Test;

public class DataTransformationFunctionsIT extends IntegrationBaseIT {

    @Test
    public void shouldApplyTheDataTranformationFunctionIfSupplied() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "data_transformation_functions_applied.xlsx");

        try {
            MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigTestHelper.getTestMempoiColumnConfig();
            MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                    .addMempoiColumnConfig(mempoiColumnConfig).build();

            MemPOI memPOI = MempoiBuilder.aMemPOI().withFile(fileDest).addMempoiSheet(mempoiSheet).build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper
                    .validateGeneratedFileDataTransformationFunction(this.createStatement(), fileDest.getAbsolutePath(),
                            TestHelper.COLUMNS, TestHelper.HEADERS,
                            null, new StandardStyleTemplate(), 0);

        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
    }


    @Test
    public void shouldNOTApplyTheDataTranformationFunctionIfSuppliedWithWrongColumnName() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "data_transformation_functions_NOT_applied.xlsx");

        try {
            MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                    .withColumnName("not-existing")
                    .build();
            MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                    .addMempoiColumnConfig(mempoiColumnConfig).build();

            MemPOI memPOI = MempoiBuilder.aMemPOI().withFile(fileDest).addMempoiSheet(mempoiSheet).build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper
                    .validateGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS,
                            null, new StandardStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
    }
}
