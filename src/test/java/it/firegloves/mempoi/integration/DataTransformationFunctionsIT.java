package it.firegloves.mempoi.integration;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.datatransformation.DataTransformationFunction;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;

import static org.junit.Assert.assertEquals;

public class DataTransformationFunctionsIT extends IntegrationBaseIT {

    // TODO refine this test: what is testing? for now is only a quick try to check MempoiColumnConfig is debug mode
    @Test
    public void testWithoutMempoiColumnConfig() throws InterruptedException, ExecutionException, SQLException {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "data_transformation_functions.xlsx");

        try {

            //            DataTransformationFunction<Object, Integer> dataTransformationFunction = (Object o) -> 5;
            final int cellValue = 5;
            DataTransformationFunction<Object, Integer> dataTransformationFunction =
                    new DataTransformationFunction<Object, Integer>() {
                @Override
                public Integer transform(Object value) throws MempoiException {
                    return cellValue;
                }
            };
            String columnName = "name";
            MempoiColumnConfig nameConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig().withColumnName(columnName)
                    .withDataTransformationFunction(dataTransformationFunction).build();
            MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt)
                    .addMempoiColumnConfig(nameConfig).build();

            MemPOI memPOI = MempoiBuilder.aMemPOI().withFile(fileDest).addMempoiSheet(mempoiSheet).build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            validateGeneratedFile(fileDest.getAbsolutePath(), columnName, cellValue);

        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
    }

    public static void validateGeneratedFile(String fileToValidate, String columnName, int columnValue) {

        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {

            Workbook wb = WorkbookFactory.create(inp);

            Sheet sheet = wb.getSheetAt(0);
            int idx = -1;
            for(int kakka = 0; kakka < TestHelper.COLUMNS.length; kakka++) {
                Cell cell = sheet.getRow(0).getCell(kakka);
                if (cell.getStringCellValue().equals(columnName)) {
                    idx = cell.getColumnIndex();
                    break;
                }
            }
            for(int i = 1; i< TestHelper.MAX_ROWS; i++){
               Assert.assertEquals(columnValue, sheet.getRow(i).getCell(idx).getNumericCellValue(), 0);
            }
        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }
}
