package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.PreparedStatement;
import java.util.concurrent.CompletableFuture;

import static org.junit.Assert.assertEquals;

public class GroupByTest extends FunctionalBaseGroupByTest {

    @Test
    public void testWithAnimals() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_group_by.xlsx");

        try {
            PreparedStatement prepStmt = this.createStatement(new String[] { "name" }, 200_000);

            // dogs sheet
            MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Grouped by name")
                    .withPrepStmt(prepStmt)
                    .withGroupByColumns(new String[] {"name"})
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
//                    .withDebug(true)
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new XSSFWorkbook())
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }
    }

}
