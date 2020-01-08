package it.firegloves.mempoi.manager;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class FileManager {

    private static final Logger logger = LoggerFactory.getLogger(FileManager.class);

    /**
     * contains the workbook configurations
     */
    private WorkbookConfig workbookConfig;


    public FileManager(WorkbookConfig workbookConfig) {
        this.workbookConfig = workbookConfig;
    }


    /**
     * write the workbook in a temporary file. in this way we can use SXSSF, apply cell formulas and evaluate them. otherwise on large dataset it will fails because SXSSF keep in memory only a few rows
     *
     * @return the written file
     */
    public File writeTempFile() {

        File tmpFile = new File(System.getProperty("java.io.tmpdir") + "mempoi_temp_" + System.currentTimeMillis() + ".xlsx");

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
                    throw new FileNotFoundException(String.format(Errors.ERR_FILE_TO_EXPORT_PARENT_DIR_CANT_BE_CREATED, file.getParentFile().getAbsolutePath()));
                }
            }

            logger.debug("writing final file {}", file.getAbsolutePath());
            this.writeFile(file);
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
}