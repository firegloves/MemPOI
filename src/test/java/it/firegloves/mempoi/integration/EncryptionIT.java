package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiEncryption;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;
import java.sql.ResultSet;
import java.util.concurrent.CompletableFuture;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class EncryptionIT extends IntegrationBaseIT {

    private final String password = "mempassword";

    @Test
    public void shouldEncryptBinaryWorkbook() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "encryption_binary_file.xls");

        MempoiEncryption mempoiEncryption = MempoiEncryption.MempoiEncryptionBuilder.aMempoiEncryption()
                .withPassword(password)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .withWorkbook(new HSSFWorkbook())
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .withMempoiEncryption(mempoiEncryption)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        Biff8EncryptionKey.setCurrentUserPassword(password);
        POIFSFileSystem fs = new POIFSFileSystem(new File(fileDest.getAbsolutePath()), true);
        HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);

        ResultSet rs = this.createStatement().executeQuery();
        Sheet sheet = wb.getSheetAt(0);
        StandardStyleTemplate standardStyleTemplate = new StandardStyleTemplate();
        // validates header row
        AssertionHelper
                .validateHeaderRow(sheet.getRow(0), TestHelper.HEADERS, standardStyleTemplate.getHeaderCellStyle(wb));
        // validates data rows
        for (int r = 1; rs.next(); r++) {
            AssertionHelper
                    .validateGeneratedFileDataRow(rs, sheet.getRow(r), TestHelper.HEADERS, standardStyleTemplate, wb);
        }

        Biff8EncryptionKey.setCurrentUserPassword(null);
    }


    @Test
    public void shouldEncryptXmlBasedWorkbook() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "encryption_xml_based.xlsx");

        MempoiEncryption mempoiEncryption = MempoiEncryption.MempoiEncryptionBuilder.aMempoiEncryption()
                .withPassword(password)
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
        XSSFWorkbook wb = (XSSFWorkbook) decryptXmlBasedFile(fileDest);

        ResultSet rs = this.createStatement().executeQuery();
        Sheet sheet = wb.getSheetAt(0);
        StandardStyleTemplate standardStyleTemplate = new StandardStyleTemplate();
        // validates header row
        AssertionHelper
                .validateHeaderRow(sheet.getRow(0), TestHelper.HEADERS, standardStyleTemplate.getHeaderCellStyle(wb));
        // validates data rows
        for (int r = 1; rs.next(); r++) {
            AssertionHelper
                    .validateGeneratedFileDataRow(rs, sheet.getRow(r), TestHelper.HEADERS, standardStyleTemplate, wb);
        }
    }


    private Workbook decryptXmlBasedFile(File fileDest) throws IOException {

        POIFSFileSystem filesystem = new POIFSFileSystem(new File(fileDest.getAbsolutePath()), true);
        EncryptionInfo info = new EncryptionInfo(filesystem);
        Decryptor d = Decryptor.getInstance(info);
        try {
            if (!d.verifyPassword(password)) {
                throw new RuntimeException("Unable to process: document is encrypted");
            }
            InputStream dataStream = d.getDataStream(filesystem);
            return WorkbookFactory.create(dataStream);

        } catch (GeneralSecurityException ex) {
            throw new RuntimeException("Unable to process encrypted document", ex);
        }
    }

}
