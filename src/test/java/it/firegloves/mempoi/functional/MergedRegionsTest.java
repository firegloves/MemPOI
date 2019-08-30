package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.sql.PreparedStatement;
import java.util.concurrent.CompletableFuture;

import static org.junit.Assert.assertEquals;

public class MergedRegionsTest extends FunctionalBaseMergedRegionsTest {

    @Test
    public void testWithFileAndMergedRegionsHSSF() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_merged_regions_HSSF.xlsx");

        try {
            PreparedStatement prepStmt = this.createStatement(null, 60_001);    // TODO create tests that exceed HSSF limits and try to manage it

            // dogs sheet
            MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Merged regions name column")
                    .withPrepStmt(prepStmt)
                    .withMergedRegionColumns(new String[] { "name" })
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
//                    .withDebug(true)
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new HSSFWorkbook())
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }
    }



    @Test
    public void testWithFileAndMergedRegionsSXSSF() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file_and_merged_regions_SXSSF.xlsx");

        try {
            PreparedStatement prepStmt = this.createStatement(null, 200_000);    // TODO create tests that exceed HSSF limits and try to manage it

            // dogs sheet
            MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Merged regions name column")
                    .withPrepStmt(prepStmt)
                    .withMergedRegionColumns(new String[] { "name" })
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
//                    .withDebug(true)
                    .withFile(fileDest)
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withWorkbook(new SXSSFWorkbook(200))
                    .addMempoiSheet(sheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }
    }

}
