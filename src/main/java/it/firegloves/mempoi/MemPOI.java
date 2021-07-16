/**
 * Library entry point
 * A MemPOI represents the report you want to generate
 * You can create it using the relative builder, that will place every required setting in its workbookConfig variable
 * After creating the desired MemPOI you can ask for the report using its 2 methods
 * Those methods ask both to {@link it.firegloves.mempoi.strategos.Strategos}  to coordinate the export process
 */
package it.firegloves.mempoi;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiReport;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.strategos.Strategos;
import java.util.concurrent.CompletableFuture;


public class MemPOI {

    private WorkbookConfig workbookConfig;

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
     * exports data into a POI report.
     * saves exported data into a File if a file path has been set using the MempoiBuilder, otherwise into a byte array
     *
     * @return a CompletableFuture containing the MempoiReport object (exported data info and metadata)
     */
    public CompletableFuture<MempoiReport> prepareMempoiReport() {

        return CompletableFuture.supplyAsync(() -> {

            try {
                Strategos strategos = new Strategos(this.workbookConfig);
                return strategos.composeMempoiReport(this.workbookConfig.getSheetList());
            } catch (MempoiException e) {
                throw e;
            } catch (Exception e) {
                throw new MempoiException(e);
            }
        });
    }

    /**
     * exports data into a POI report. saves exported data into the File identified by the file variable
     *
     * @return a CompletableFuture containing the string of the file name (absolute path) of the generated file
     */
    @Deprecated
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
    @Deprecated
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
