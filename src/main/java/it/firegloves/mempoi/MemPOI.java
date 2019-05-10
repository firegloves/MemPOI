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

    public WorkbookConfig getWorkbookConfig() {
        return workbookConfig;
    }

    public void setWorkbookConfig(WorkbookConfig workbookConfig) {
        this.workbookConfig = workbookConfig;
    }

    /**
     * exports data into a POI report. saves exported data into the File identified by the file variable
     *
     * @return a CompletableFuture containing the string of the file name (absolute path) of the generated file
     */
    public CompletableFuture<String> prepareMempoiReportToFile() throws MempoiException {

        return CompletableFuture.supplyAsync(() -> {

            if (null == this.workbookConfig.getFile()) {
                throw new MempoiException("Error: report to file requested but no file was specified in the config");
            }

            Strategos strategos = new Strategos(this.workbookConfig);
            return strategos.generateMempoiReportToFile(this.workbookConfig.getSheetList(), this.workbookConfig.getFile());
        });
    }


    /**
     * exports data into a POI report. return the byte array of the exported data
     *
     * @return a CompletableFuture containing the byte arrat of the exported data
     */
    public CompletableFuture<byte[]> prepareMempoiReportToByteArray() throws MempoiException {

        return CompletableFuture.supplyAsync(() -> {

            Strategos strategos = new Strategos(this.workbookConfig);
            return strategos.generateMempoiReportToByteArray();
        });
    }
}