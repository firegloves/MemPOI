/**
 * library entry point
 */
package it.firegloves.mempoi;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.exception.MempoiException;

import java.util.concurrent.CompletableFuture;


public class MemPOI {

    private WorkbookConfig workbookConfig;

    public MemPOI() {
    }

    public MemPOI(WorkbookConfig workbookConfig) {
        this.workbookConfig = workbookConfig;
    }

    /**
     * @return the current WorkbookConfig containing report configurations
     */
    public WorkbookConfig getWorkbookConfig() {
        return workbookConfig;
    }

    /**
     * set a WorkbookConfig containing report configurations
     * @param workbookConfig the WorkbookConfig to use for generating export
     */
    public void setWorkbookConfig(WorkbookConfig workbookConfig) {
        this.workbookConfig = workbookConfig;
    }

    /**
     * exports data into a POI report. saves exported data into the File identified by the file variable
     *
     * @return a CompletableFuture containing the string of the file name (absolute path) of the generated file
     */
    public CompletableFuture<String> prepareMempoiReportToFile() {

        return CompletableFuture.supplyAsync(() -> {

            if (null == this.workbookConfig.getFile()) {
                throw new MempoiException("Error: report to file requested but no file was specified in the config");
            }

            try {
                Strategos strategos = new Strategos(this.workbookConfig);
                return strategos.generateMempoiReportToFile(this.workbookConfig.getSheetList(), this.workbookConfig.getFile());
            } catch (Exception e) {
                throw new MempoiException(e);
            }
        });
    }


    /**
     * exports data into a POI report. return the byte array of the exported data
     *
     * @return a CompletableFuture containing the byte arrat of the exported data
     */
    public CompletableFuture<byte[]> prepareMempoiReportToByteArray() {

        return CompletableFuture.supplyAsync(() -> {

            try {
                Strategos strategos = new Strategos(this.workbookConfig);
                return strategos.generateMempoiReportToByteArray();
            } catch (Exception e) {
                throw new MempoiException(e);
            }
        });
    }
}