package it.firegloves.mempoi.manager;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiEncryption;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import lombok.AllArgsConstructor;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@AllArgsConstructor
public class FileManager {

    private static final Logger logger = LoggerFactory.getLogger(FileManager.class);

    /**
     * contains the workbook configurations
     */
    private final WorkbookConfig workbookConfig;


    /**
     * write the workbook in a temporary file. in this way we can use SXSSF, apply cell formulas and evaluate them. otherwise on large dataset it will fails because SXSSF keep in memory only a few rows
     *
     * @return the written file
     */
    public File writeTempFile() {

        // TODO encrypt temp file too

        File tmpFile = new File(
                System.getProperty("java.io.tmpdir") + File.separator + "mempoi_temp_" + System.currentTimeMillis()
                        + ".xlsx");

        this.writeFile(tmpFile);
        logger.debug("MemPOI temp file created: {}", tmpFile.getAbsolutePath());

        return tmpFile;
    }


    /**
     * write report file
     *
     * @param file the file with path to which write exported data
     * @return created file name with path
     * @throws MempoiException if write operation fails
     */
    public String createFinalFile(File file) {
        try {
            // checks path consistency
            if (!file.getAbsoluteFile().getParentFile().exists()) {
                if (file.getAbsoluteFile().getParentFile().mkdirs()) {
                    logger.debug("CREATED FILE TO EXPORT PARENT DIR: {}", file.getParentFile().getAbsolutePath());
                } else {
                    throw new FileNotFoundException(String.format(Errors.ERR_FILE_TO_EXPORT_PARENT_DIR_CANT_BE_CREATED,
                            file.getParentFile().getAbsolutePath()));
                }
            }

            logger.debug("writing final file {}", file.getAbsolutePath());
            ByteArrayOutputStream byteArrayOutputStream = this.writeWorkbookToByteArrayOutputStream();

            if (null != this.workbookConfig.getMempoiEncryption()) {
                byteArrayOutputStream = this.encrypt(new ByteArrayInputStream(byteArrayOutputStream.toByteArray()));
            }

            this.writeFile(file, new ByteArrayInputStream(byteArrayOutputStream.toByteArray()));

            logger.debug("written final file {}", file.getAbsolutePath());

            return file.getAbsolutePath();

        } catch (Exception e) {
            throw new MempoiException(e);
        } finally {
            this.closeWorkbook();
        }
    }


    /**
     * writes the the workbook in the WorkbookConfig in the received file
     *
     * @param file the file where write the workbook
     */
    private void writeFile(File file, ByteArrayInputStream inputStream) {

        try {
            // writes data to file
            try (FileOutputStream outputStream = new FileOutputStream(file)) {
                IOUtils.copy(inputStream, outputStream);
            }
        } catch (Exception e) {
            throw new MempoiException(e);
        } finally {
            this.closeWorkbook();
        }
    }

    /**
     * writes the the workbook in the WorkbookConfig in the received file
     *
     * @param file the file where write the workbook
     */
    private void writeFile(File file) {

        try {
            // writes data to file
            try (FileOutputStream outputStream = new FileOutputStream(file)) {
                this.workbookConfig.getWorkbook().write(outputStream);
            }
        } catch (Exception e) {
            throw new MempoiException(e);
        } finally {
            this.closeWorkbook();
        }
    }


    /**
     * writes the the workbook in the WorkbookConfig in the received file
     */
    private ByteArrayOutputStream writeWorkbookToByteArrayOutputStream() {

        try {
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                this.workbookConfig.getWorkbook().write(bos);
                return bos;
            }
        } catch (Exception e) {
            throw new MempoiException(e);
        } finally {
            this.closeWorkbook();
        }
    }


    /**
     * write report to byte array
     *
     * @return the byte array corresponding to the poi export
     * @throws MempoiException
     */
    public byte[] writeToByteArray() {

        try {

            // writes data to file
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                this.workbookConfig.getWorkbook().write(bos);
                return bos.toByteArray();
            }

        } catch (Exception e) {
            throw new MempoiException(e);
        } finally {
            this.closeWorkbook();
        }
    }


    /**
     * closes the current workbook
     */
    private void closeWorkbook() {

        if (null == this.workbookConfig.getWorkbook()) {
            logger.warn(Errors.WARN_NULL_WB_NOT_CLOSED);
            throw new MempoiException(Errors.ERR_WORKBOOK_NULL);
        }

        // deletes temp poi file
        if (this.workbookConfig.getWorkbook() instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook) this.workbookConfig.getWorkbook()).dispose();
        }

        // closes the workbook
        try {
            this.workbookConfig.getWorkbook().close();
        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }

    /**
     * @param input
     * @return
     * @throws Exception
     */
    private ByteArrayOutputStream encrypt(InputStream input) throws Exception {

        return this.isXmlBasedWoorkbook(workbookConfig.getWorkbook()) ?
                this.encryptXmlBasedWorkbook(workbookConfig.getMempoiEncryption(), input) :
                this.encryptBinaryWorkbook(workbookConfig.getMempoiEncryption(), input);

    }


    private ByteArrayOutputStream encryptXmlBasedWorkbook(MempoiEncryption mempoiEncryption, InputStream input)
            throws Exception {

        POIFSFileSystem poiFs = new POIFSFileSystem();

        Encryptor enc = mempoiEncryption.getEncryptionInfo().getEncryptor();
        enc.confirmPassword(mempoiEncryption.getPassword());

        try (OPCPackage opc = OPCPackage.open(input);
                OutputStream os = enc.getDataStream(poiFs)) {
            opc.save(os);
        }

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        poiFs.writeFilesystem(byteArrayOutputStream);
        return byteArrayOutputStream;
    }


    private ByteArrayOutputStream encryptBinaryWorkbook(MempoiEncryption mempoiEncryption, InputStream input)
            throws Exception {

        Biff8EncryptionKey.setCurrentUserPassword(mempoiEncryption.getPassword());
        POIFSFileSystem poiFs = new POIFSFileSystem(input);
        HSSFWorkbook hwb = new HSSFWorkbook(poiFs.getRoot(), true);
        Biff8EncryptionKey.setCurrentUserPassword(null);
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        hwb.write(byteArrayOutputStream);
        return byteArrayOutputStream;
    }


    /**
     * check if the received workbook is xml based
     *
     * @param workbook the workbook to check
     * @return true if the received workbook is xml based, false otherwise
     */
    private boolean isXmlBasedWoorkbook(Workbook workbook) {
        return !(workbook instanceof HSSFWorkbook);
    }
}
