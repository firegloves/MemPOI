package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiColumnConfig.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.datatransformation.DataTransformationFunction;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.MempoiColumnConfigTestHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.util.concurrent.CompletableFuture;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;

public class DataTransformationFunctionsIT extends IntegrationBaseIT {

    @Rule
    public ExpectedException expectedEx = ExpectedException.none();

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
                            TestHelper.COLUMNS, TestHelper.HEADERS, null, new StandardStyleTemplate(), 0,
                            MempoiColumnConfigTestHelper.CELL_VALUE);

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

    @Test
    public void shouldApplyTheDataTranformationFunctionIfSuppliedAndTheCastMethodIsInvokedWithValidValues()
            throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "data_transformation_functions_applied_with_cast.xlsx");

        try {
            MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                    .withColumnName(MempoiColumnConfigTestHelper.COLUMN_NAME)
                    .withDataTransformationFunction(new DataTransformationFunction<Integer>() {
                        @Override
                        public Integer transform(Object value) throws MempoiException {
                            String typedValue = super.<String>cast(value, String.class);
                            return MempoiColumnConfigTestHelper.CELL_VALUE;
                        }
                    })
                    .build();
            MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                    .addMempoiColumnConfig(mempoiColumnConfig).build();

            MemPOI memPOI = MempoiBuilder.aMemPOI().withFile(fileDest).addMempoiSheet(mempoiSheet).build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            AssertionHelper
                    .validateGeneratedFileDataTransformationFunction(this.createStatement(), fileDest.getAbsolutePath(),
                            TestHelper.COLUMNS, TestHelper.HEADERS, null, new StandardStyleTemplate(), 0,
                            MempoiColumnConfigTestHelper.CELL_VALUE);

        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
    }


    @Test(expected = MempoiException.class)
    public void shouldThrowExceptionIfDataTransformationFunctionSuppliedAndTheCastMethodIsInvokedWithInvalidValues() {

        MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                .withColumnName(MempoiColumnConfigTestHelper.COLUMN_NAME)
                .withDataTransformationFunction(new DataTransformationFunction<Integer>() {
                    @Override
                    public Integer transform(Object value) throws MempoiException {
                        Integer typedValue = super.<Integer>cast(value, Integer.class);
                        return typedValue - 1;
                    }
                })
                .build();
        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                .addMempoiColumnConfig(mempoiColumnConfig).build();

        MemPOI memPOI = MempoiBuilder.aMemPOI().addMempoiSheet(mempoiSheet).build();

        memPOI.prepareMempoiReportToByteArray().join();

    }
}
