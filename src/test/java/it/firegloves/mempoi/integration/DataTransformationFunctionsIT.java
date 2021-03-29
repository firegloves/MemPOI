package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiColumnConfig.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.datatransformation.BooleanDataTransformationFunction;
import it.firegloves.mempoi.domain.datatransformation.DateDataTransformationFunction;
import it.firegloves.mempoi.domain.datatransformation.DoubleDataTransformationFunction;
import it.firegloves.mempoi.domain.datatransformation.StringDataTransformationFunction;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.MempoiColumnConfigTestHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;

public class DataTransformationFunctionsIT extends IntegrationBaseIT {

    @Rule
    public ExpectedException expectedEx = ExpectedException.none();
    private final String strValue = "terrrrribbblee!";

    /*****************************************************************************
     * SHOULD APPLY DATA TRANSFORMATION FUNCTIONS
     ****************************************************************************/

    @Test
    public void shouldApplyStringDataTranformationFunctionIfSupplied() {

        List<String> colNameList = Arrays.asList("name", "usefulChar");

        colNameList.forEach(colName -> {

            File fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                    "string_" + colName + "_data_transf_function.xlsx");

            try {
                this.prepStmt = createStatement();

                MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                        .withColumnName(colName)
                        .withDataTransformationFunction(new StringDataTransformationFunction<Integer>() {
                            @Override
                            public Integer transform(String value) throws MempoiException {
                                return MempoiColumnConfigTestHelper.CELL_VALUE;
                            }
                        })
                        .build();

                MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                        .addMempoiColumnConfig(mempoiColumnConfig).build();

                MemPOI memPOI = MempoiBuilder.aMemPOI().withFile(fileDest).addMempoiSheet(mempoiSheet).build();

                CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
                assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

                this.assertOnGeneratedFileDataTransformationFunction(this.createStatement(), fileDest.getAbsolutePath(),
                        TestHelper.HEADERS_2, new StandardStyleTemplate(),
                        (double) MempoiColumnConfigTestHelper.CELL_VALUE, Double.class,
                        colName);

            } catch (Exception e) {
                AssertionHelper.failAssertion(e);
            }
        });
    }

    @Test
    public void shouldApplyBooleanDataTranformationFunctionIfSupplied() {

        List<String> colNameList = Arrays.asList("valid", "bitTwo");

        colNameList.forEach(colName -> {

            File fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                    "boolean_" + colName + "_data_transf_function.xlsx");

            try {
                this.prepStmt = createStatement();

                MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                        .withColumnName(colName)
                        .withDataTransformationFunction(new BooleanDataTransformationFunction<Integer>() {
                            @Override
                            public Integer transform(Boolean value) throws MempoiException {
                                return MempoiColumnConfigTestHelper.CELL_VALUE;
                            }
                        })
                        .build();

                MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                        .addMempoiColumnConfig(mempoiColumnConfig).build();

                MemPOI memPOI = MempoiBuilder.aMemPOI().withFile(fileDest).addMempoiSheet(mempoiSheet).build();

                CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
                assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

                this.assertOnGeneratedFileDataTransformationFunction(this.createStatement(), fileDest.getAbsolutePath(),
                        TestHelper.HEADERS_2, new StandardStyleTemplate(),
                        (double) MempoiColumnConfigTestHelper.CELL_VALUE, Double.class,
                        colName);

            } catch (Exception e) {
                AssertionHelper.failAssertion(e);
            }
        });
    }

    @Test
    public void shouldApplyDateDataTranformationFunctionIfSupplied() {

        List<String> colNameList = Arrays.asList("creation_date", "dateTime", "STAMPONE", "attempato");

        colNameList.forEach(colName -> {

            File fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                    "date_" + colName + "_data_transf_function.xlsx");

            try {
                this.prepStmt = createStatement();

                MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                        .withColumnName(colName)
                        .withDataTransformationFunction(new DateDataTransformationFunction<Integer>() {
                            @Override
                            public Integer transform(Date value) throws MempoiException {
                                return MempoiColumnConfigTestHelper.CELL_VALUE;
                            }
                        })
                        .build();

                MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                        .addMempoiColumnConfig(mempoiColumnConfig).build();

                MemPOI memPOI = MempoiBuilder.aMemPOI().withFile(fileDest).addMempoiSheet(mempoiSheet).build();

                CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
                assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

                this.assertOnGeneratedFileDataTransformationFunction(this.createStatement(), fileDest.getAbsolutePath(),
                        TestHelper.HEADERS_2, new StandardStyleTemplate(),
                        (double) MempoiColumnConfigTestHelper.CELL_VALUE, Double.class,
                        colName);

            } catch (Exception e) {
                AssertionHelper.failAssertion(e);
            }
        });
    }


    @Test
    public void shouldApplyDoubleDataTranformationFunctionIfSupplied() {

        List<String> colNameList = Arrays
                .asList("decimalOne", "doublone", "floattone", "interao", "mediano", "interuccio");

        colNameList.forEach(colName -> {

            File fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                    "double_" + colName + "_data_transf_function.xlsx");

            try {
                this.prepStmt = createStatement();

                MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                        .withColumnName(colName)
                        .withDataTransformationFunction(new DoubleDataTransformationFunction<String>() {
                            @Override
                            public String transform(Double value) throws MempoiException {
                                return strValue;
                            }
                        })
                        .build();

                MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                        .addMempoiColumnConfig(mempoiColumnConfig).build();

                MemPOI memPOI = MempoiBuilder.aMemPOI().withFile(fileDest).addMempoiSheet(mempoiSheet).build();

                CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
                assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

                this.assertOnGeneratedFileDataTransformationFunction(this.createStatement(), fileDest.getAbsolutePath(),
                        TestHelper.HEADERS_2, new StandardStyleTemplate(), strValue, String.class, colName);

            } catch (Exception e) {
                AssertionHelper.failAssertion(e);
            }
        });
    }


    /*****************************************************************************
     * NULL VALUES
     ****************************************************************************/

    @Test(expected = Test.None.class)
    public void givenANullValueReadByDBTheDateDataTranformationFunctionShouldReceiveNull() {

        List<String> colNameList = Arrays.asList("creation_date", "dateTime", "STAMPONE", "attempato");

        colNameList.forEach(colName -> {

            try {
                this.prepStmt = createNullValuesStatement();

                MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                        .withColumnName(colName)
                        .withDataTransformationFunction(new DateDataTransformationFunction<Integer>() {
                            @Override
                            public Integer transform(Date value) throws MempoiException {
                                if (null != value) {
                                    throw new MempoiException(
                                            "data transformation function did not receive a null value");
                                }
                                return 5;
                            }
                        })
                        .build();

                MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                        .addMempoiColumnConfig(mempoiColumnConfig).build();

                MemPOI memPOI = MempoiBuilder.aMemPOI().addMempoiSheet(mempoiSheet).build();

                memPOI.prepareMempoiReportToByteArray().get();

            } catch (Exception e) {
                AssertionHelper.failAssertion(e);
            }
        });
    }

    @Test(expected = Test.None.class)
    public void givenANullValueReadByDBTheBooleanDataTranformationFunctionShouldReceivePrimitiveDefault() {

        List<String> colNameList = Arrays.asList("valid", "bitTwo");

        colNameList.forEach(colName -> {

            try {
                this.prepStmt = createNullValuesStatement();

                MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                        .withColumnName(colName)
                        .withDataTransformationFunction(new BooleanDataTransformationFunction<Integer>() {
                            @Override
                            public Integer transform(Boolean value) throws MempoiException {
                                if (value) {
                                    throw new MempoiException(
                                            "boolean data transformation function did not receive primitive default value");
                                }
                                return 5;
                            }
                        })
                        .build();

                MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                        .addMempoiColumnConfig(mempoiColumnConfig).build();

                MemPOI memPOI = MempoiBuilder.aMemPOI().addMempoiSheet(mempoiSheet).build();

                memPOI.prepareMempoiReportToByteArray().get();

            } catch (Exception e) {
                AssertionHelper.failAssertion(e);
            }
        });
    }

    @Test(expected = Test.None.class)
    public void givenANullValueReadByDBTheBooleanDataTranformationFunctionShouldReceiveNullIfNullValuesOverPrimitiveDetaultOnes() {

        List<String> colNameList = Arrays.asList("valid", "bitTwo");

        colNameList.forEach(colName -> {

            try {
                this.prepStmt = createNullValuesStatement();

                MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                        .withColumnName(colName)
                        .withDataTransformationFunction(new BooleanDataTransformationFunction<Integer>() {
                            @Override
                            public Integer transform(Boolean value) throws MempoiException {
                                if (null != value) {
                                    throw new MempoiException(
                                            "boolean data transformation function did not receive null value");
                                }
                                return 5;
                            }
                        })
                        .build();

                MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                        .addMempoiColumnConfig(mempoiColumnConfig).build();

                MemPOI memPOI = MempoiBuilder.aMemPOI().addMempoiSheet(mempoiSheet)
                        .withNullValuesOverPrimitiveDetaultOnes(true).build();

                memPOI.prepareMempoiReportToByteArray().get();

            } catch (Exception e) {
                AssertionHelper.failAssertion(e);
            }
        });
    }

    @Test(expected = Test.None.class)
    public void givenANullValueReadByDBTheStringDataTranformationFunctionShouldReceiveNullValue() {

        List<String> colNameList = Arrays.asList("name", "usefulChar");

        colNameList.forEach(colName -> {

            try {
                this.prepStmt = createNullValuesStatement();

                MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                        .withColumnName(colName)
                        .withDataTransformationFunction(new StringDataTransformationFunction<Integer>() {
                            @Override
                            public Integer transform(String value) throws MempoiException {
                                if (null != value) {
                                    throw new MempoiException(
                                            "string data transformation function did not receive a null value");
                                }
                                return 5;
                            }
                        })
                        .build();

                MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                        .addMempoiColumnConfig(mempoiColumnConfig).build();

                MemPOI memPOI = MempoiBuilder.aMemPOI().addMempoiSheet(mempoiSheet)
                        .withNullValuesOverPrimitiveDetaultOnes(true).build();

                memPOI.prepareMempoiReportToByteArray().get();

            } catch (Exception e) {
                AssertionHelper.failAssertion(e);
            }
        });
    }

    @Test(expected = Test.None.class)
    public void givenANullValueReadByDBTheDoubleDataTranformationFunctionShouldReceivePrimitiveDefault() {

        List<String> colNameList = Arrays
                .asList("decimalOne", "doublone", "floattone", "interao", "mediano", "interuccio");

        colNameList.forEach(colName -> {

            try {
                this.prepStmt = createNullValuesStatement();

                MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                        .withColumnName(colName)
                        .withDataTransformationFunction(new DoubleDataTransformationFunction<String>() {
                            @Override
                            public String transform(Double value) throws MempoiException {
                                if (value != 0d) {
                                    throw new MempoiException(
                                            "double data transformation function did not receive primitive value");
                                }
                                return strValue;
                            }
                        })
                        .build();

                MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                        .addMempoiColumnConfig(mempoiColumnConfig).build();

                MemPOI memPOI = MempoiBuilder.aMemPOI().addMempoiSheet(mempoiSheet).build();

                memPOI.prepareMempoiReportToByteArray().get();

            } catch (Exception e) {
                AssertionHelper.failAssertion(e);
            }
        });
    }

    @Test(expected = Test.None.class)
    public void givenANullValueReadByDBTheDoubleDataTranformationFunctionShouldReceiveNullIfNullValuesOverPrimitiveDetaultOnesIsTrue() {

        List<String> colNameList = Arrays
                .asList("decimalOne", "doublone", "floattone", "interao", "mediano", "interuccio");

        colNameList.forEach(colName -> {

            try {
                this.prepStmt = createNullValuesStatement();

                MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                        .withColumnName(colName)
                        .withDataTransformationFunction(new DoubleDataTransformationFunction<String>() {
                            @Override
                            public String transform(Double value) throws MempoiException {
                                if (null != value) {
                                    throw new MempoiException(
                                            "data transformation function did not receive a null value");
                                }
                                return strValue;
                            }
                        })
                        .build();

                MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                        .addMempoiColumnConfig(mempoiColumnConfig).build();

                MemPOI memPOI = MempoiBuilder.aMemPOI().addMempoiSheet(mempoiSheet)
                        .withNullValuesOverPrimitiveDetaultOnes(true).build();

                memPOI.prepareMempoiReportToByteArray().get();

            } catch (Exception e) {
                AssertionHelper.failAssertion(e);
            }
        });
    }

    /*****************************************************************************
     * OTHERS
     ****************************************************************************/

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
                    .assertOnGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS_2,
                            TestHelper.HEADERS_2,
                            null, new StandardStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
    }


    @Test(expected = MempoiException.class)
    public void shouldThrowExceptionIfTheDataTransformationFunctionSuppliedAcceptATypeNotEqualsToTheOneReturnedByDB() {

        MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                .withColumnName(MempoiColumnConfigTestHelper.COLUMN_NAME)
                .withDataTransformationFunction(new DoubleDataTransformationFunction<Integer>() {
                    @Override
                    public Integer transform(Double value) throws MempoiException {
                        return 1;
                    }
                })
                .build();
        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                .addMempoiColumnConfig(mempoiColumnConfig).build();

        MemPOI memPOI = MempoiBuilder.aMemPOI().addMempoiSheet(mempoiSheet).build();

        memPOI.prepareMempoiReportToByteArray().join();
    }


    /*****************************************************************************
     * UTILITIES
     ****************************************************************************/


    @Override
    public PreparedStatement createStatement() throws SQLException {
        return this.conn.prepareStatement(
                this.createQuery(TestHelper.TABLE_EXPORT_TEST, TestHelper.COLUMNS_2, TestHelper.HEADERS_2, -1));
    }

    public PreparedStatement createNullValuesStatement() throws SQLException {
        return this.conn.prepareStatement(
                this.createQuery(TestHelper.TABLE_NULL_VALUES, TestHelper.COLUMNS_2, TestHelper.HEADERS_2, 100));
    }


    /**
     * opens the received generated xlsx file and applies generic asserts
     *
     * @param prepStmt       the PreparedStatement to execute to validate data
     * @param fileToValidate the absolute filename of the xlsx file on which apply the generic asserts
     * @param headers        the array of headers name required
     */
    private void assertOnGeneratedFileDataTransformationFunction(PreparedStatement prepStmt,
            String fileToValidate, String[] headers, StyleTemplate styleTemplate, Object transformedValue,
            Class transformedValueCastClass, String transformedColumnName) {

        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {

            ResultSet rs = prepStmt.executeQuery();

            Workbook wb = WorkbookFactory.create(inp);

            Sheet sheet = wb.getSheetAt(0);

            // validates header row
            AssertionHelper.assertOnHeaderRow(sheet.getRow(0), headers,
                    null != styleTemplate ? styleTemplate.getHeaderCellStyle(wb) : null);

            // validates data rows
            for (int r = 1; rs.next(); r++) {
                validateGeneratedFileDataRowDataTransformationFunction(rs, sheet.getRow(r), headers, styleTemplate, wb,
                        transformedValue, transformedValueCastClass, transformedColumnName);
            }

        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }


    /**
     * validate one Row of the generic export (created with createStatement()) against the resulting ResultSet generated
     * by the execution of the same method
     *
     * @param rs            the ResultSet against which validate the Row
     * @param row           the Row to validate against the ResultSet
     * @param headers       the array of columns name, useful to retrieve data from the ResultSet
     * @param styleTemplate StyleTemplate to get styles to validate
     * @param wb            the curret Workbook
     */
    private void validateGeneratedFileDataRowDataTransformationFunction(ResultSet rs, Row row, String[] headers,
            StyleTemplate styleTemplate, Workbook wb, Object transformedValue, Class transformedValueCastClass,
            String transformedColumnName) {

        try {

            assertEquals(rs.getInt(TestHelper.COLUMNS_2[0]), (int) row.getCell(0).getNumericCellValue());
            this.assertOnDateColumn(rs, row, 1, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);
            this.assertOnDateColumn(rs, row, 2, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);
            this.assertOnDateColumn(rs, row, 3, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);
            this.assertOnStringColumn(rs, row, 4, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);
            this.assertOnBooleanColumn(rs, row, 5, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);
            this.assertOnStringColumn(rs, row, 6, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);
            this.assertOnDoubleColumn(rs, row, 7, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);
            this.assertOnBooleanColumn(rs, row, 8, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);
            this.assertOnDoubleColumn(rs, row, 9, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);
            this.assertOnFloatColumn(rs, row, 10, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);
            this.assertOnIntegerColumn(rs, row, 11, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);
            this.assertOnIntegerColumn(rs, row, 12, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);
            this.assertOnDateColumn(rs, row, 13, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);
            this.assertOnIntegerColumn(rs, row, 14, transformedColumnName, transformedValue, transformedValueCastClass,
                    styleTemplate, wb);

        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }

    }

    private void assertOnIntegerColumn(ResultSet rs, Row row, int columnIndex, String transformedColumnName,
            Object transformedValue, Class transformedValueCastClass, StyleTemplate styleTemplate, Workbook wb)
            throws SQLException {

        if (TestHelper.HEADERS_2[columnIndex].equals(transformedColumnName)) {
            assertEquals(transformedValueCastClass.cast(transformedValue),
                    row.getCell(columnIndex).getStringCellValue());
            AssertionHelper.assertOnCellStyle(row.getCell(columnIndex).getCellStyle(),
                    styleTemplate.getCommonDataCellStyle(wb));
        } else {
            assertEquals(rs.getInt(TestHelper.HEADERS_2[columnIndex]),
                    (int) row.getCell(columnIndex).getNumericCellValue());
            AssertionHelper.assertOnCellStyle(row.getCell(columnIndex).getCellStyle(),
                    styleTemplate.getIntegerCellStyle(wb));
        }
    }

    private void assertOnDoubleColumn(ResultSet rs, Row row, int columnIndex, String transformedColumnName,
            Object transformedValue, Class transformedValueCastClass, StyleTemplate styleTemplate, Workbook wb)
            throws SQLException {

        if (TestHelper.HEADERS_2[columnIndex].equals(transformedColumnName)) {
            assertEquals(transformedValueCastClass.cast(transformedValue),
                    row.getCell(columnIndex).getStringCellValue());
            AssertionHelper.assertOnCellStyle(row.getCell(columnIndex).getCellStyle(),
                    styleTemplate.getCommonDataCellStyle(wb));
        } else {
            assertEquals(rs.getDouble(TestHelper.HEADERS_2[columnIndex]),
                    row.getCell(columnIndex).getNumericCellValue(), 0.1);
            AssertionHelper.assertOnCellStyle(row.getCell(columnIndex).getCellStyle(),
                    styleTemplate.getIntegerCellStyle(wb));
        }
    }

    private void assertOnFloatColumn(ResultSet rs, Row row, int columnIndex, String transformedColumnName,
            Object transformedValue, Class transformedValueCastClass, StyleTemplate styleTemplate, Workbook wb)
            throws SQLException {

        if (TestHelper.HEADERS_2[columnIndex].equals(transformedColumnName)) {
            assertEquals(transformedValueCastClass.cast(transformedValue),
                    row.getCell(columnIndex).getStringCellValue());
            AssertionHelper.assertOnCellStyle(row.getCell(columnIndex).getCellStyle(),
                    styleTemplate.getCommonDataCellStyle(wb));
        } else {
            assertEquals(rs.getFloat(TestHelper.HEADERS_2[columnIndex]),
                    (int) row.getCell(columnIndex).getNumericCellValue(), 0.1);
            AssertionHelper.assertOnCellStyle(row.getCell(columnIndex).getCellStyle(),
                    styleTemplate.getFloatingPointCellStyle(wb));
        }
    }

    private void assertOnDateColumn(ResultSet rs, Row row, int columnIndex, String transformedColumnName,
            Object transformedValue, Class transformedValueCastClass, StyleTemplate styleTemplate, Workbook wb)
            throws SQLException {

        if (TestHelper.HEADERS_2[columnIndex].equals(transformedColumnName)) {
            assertEquals(transformedValueCastClass.cast(transformedValue),
                    row.getCell(columnIndex).getNumericCellValue());
            AssertionHelper.assertOnCellStyle(row.getCell(columnIndex).getCellStyle(),
                    styleTemplate.getIntegerCellStyle(wb));
        } else {
            assertEquals(rs.getTimestamp(TestHelper.HEADERS_2[columnIndex]).getTime(),
                    row.getCell(columnIndex).getDateCellValue().getTime());
            AssertionHelper.assertOnCellStyle(row.getCell(columnIndex).getCellStyle(),
                    styleTemplate.getDateCellStyle(wb));
        }
    }

    private void assertOnBooleanColumn(ResultSet rs, Row row, int columnIndex, String transformedColumnName,
            Object transformedValue, Class transformedValueCastClass, StyleTemplate styleTemplate, Workbook wb)
            throws SQLException {

        if (TestHelper.HEADERS_2[columnIndex].equals(transformedColumnName)) {
            assertEquals(transformedValueCastClass.cast(transformedValue),
                    row.getCell(columnIndex).getNumericCellValue());
            AssertionHelper.assertOnCellStyle(row.getCell(columnIndex).getCellStyle(),
                    styleTemplate.getIntegerCellStyle(wb));
        } else {
            assertEquals(rs.getBoolean(TestHelper.HEADERS_2[columnIndex]),
                    row.getCell(columnIndex).getBooleanCellValue());
            AssertionHelper.assertOnCellStyle(row.getCell(columnIndex).getCellStyle(),
                    styleTemplate.getCommonDataCellStyle(wb));
        }
    }

    private void assertOnStringColumn(ResultSet rs, Row row, int columnIndex, String transformedColumnName,
            Object transformedValue, Class transformedValueCastClass, StyleTemplate styleTemplate, Workbook wb)
            throws SQLException {

        if (TestHelper.HEADERS_2[columnIndex].equals(transformedColumnName)) {
            assertEquals(transformedValueCastClass.cast(transformedValue),
                    row.getCell(columnIndex).getNumericCellValue());
            AssertionHelper.assertOnCellStyle(row.getCell(columnIndex).getCellStyle(),
                    styleTemplate.getIntegerCellStyle(wb));
        } else {
            assertEquals(rs.getString(TestHelper.HEADERS_2[columnIndex]),
                    row.getCell(columnIndex).getStringCellValue());
            AssertionHelper.assertOnCellStyle(row.getCell(columnIndex).getCellStyle(),
                    styleTemplate.getCommonDataCellStyle(wb));
        }
    }
}
