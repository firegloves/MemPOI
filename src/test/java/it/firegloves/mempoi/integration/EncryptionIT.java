package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotEquals;
import static org.junit.Assert.assertNotNull;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiEncryption;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.io.FileOutputStream;
import java.util.concurrent.CompletableFuture;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class EncryptionIT extends IntegrationBaseIT {

    @Test
    public void shouldEncryptWhenEncryptionConfigIsAvailable() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "encryption.xlsx");

        try {

            MempoiEncryption mempoiEncryption = MempoiEncryption.builder()
                    .withPassword("faciolo")
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withWorkbook(new SXSSFWorkbook())
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withMempoiEncryption(mempoiEncryption)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            // TODO open the encrypted doc to validate using password

            AssertionHelper
                    .validateGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS,
                            null, new StandardStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void shouldEncryptBinaryWorkbook() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "encryption_binary_file.xls");

        try {

            MempoiEncryption mempoiEncryption = MempoiEncryption.builder()
                    .withPassword("faciolo")
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withWorkbook(new HSSFWorkbook())
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withMempoiEncryption(mempoiEncryption)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            // TODO open the encrypted doc to validate using password

            AssertionHelper
                    .validateGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS,
                            null, new StandardStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void shouldEncryptXmlBasedWorkbook() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "encryption_xml_based.xlsx");

        try {

            MempoiEncryption mempoiEncryption = MempoiEncryption.builder()
                    .withPassword("faciolo")
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withWorkbook(new XSSFWorkbook())
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withMempoiEncryption(mempoiEncryption)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            // TODO open the encrypted doc to validate using password

            AssertionHelper
                    .validateGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS,
                            null, new StandardStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
