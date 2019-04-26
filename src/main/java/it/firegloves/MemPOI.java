/**
 * library entry point
 */
package it.firegloves;

import it.firegloves.domain.MempoiSheet;
import it.firegloves.domain.footer.MempoiFooter;
import it.firegloves.domain.footer.MempoiSubFooter;
import it.firegloves.exception.MempoiException;
import it.firegloves.styles.MempoiReportStyler;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.util.List;
import java.util.concurrent.CompletableFuture;


public class MemPOI {

    private List<MempoiSheet> mempoiSheetList;
    private MempoiSubFooter mempoiSubFooter;
    private MempoiFooter mempoiFooter;
    private MempoiReportStyler mempoiReportStyler;
    private Workbook workbook;
    private boolean adjustColumnWidth;
    private File file;
    private boolean evaluateCellFormulas;

    public MemPOI() {
    }

    public List<MempoiSheet> getMempoiSheetList() {
        return mempoiSheetList;
    }

    public void setMempoiSheetList(List<MempoiSheet> mempoiSheetList) {
        this.mempoiSheetList = mempoiSheetList;
    }

    public MempoiReportStyler getMempoiReportStyler() {
        return mempoiReportStyler;
    }

    public void setMempoiReportStyler(MempoiReportStyler mempoiReportStyler) {
        this.mempoiReportStyler = mempoiReportStyler;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public boolean isAdjustColumnWidth() {
        return adjustColumnWidth;
    }

    public void setAdjustColumnWidth(boolean adjustColumnWidth) {
        this.adjustColumnWidth = adjustColumnWidth;
    }

    public File getFile() {
        return file;
    }

    public void setFile(File file) {
        this.file = file;
    }

    public MempoiSubFooter getMempoiSubFooter() {
        return mempoiSubFooter;
    }

    public void setMempoiSubFooter(MempoiSubFooter mempoiSubFooter) {
        this.mempoiSubFooter = mempoiSubFooter;
    }

    public MempoiFooter getMempoiFooter() {
        return mempoiFooter;
    }

    public void setMempoiFooter(MempoiFooter mempoiFooter) {
        this.mempoiFooter = mempoiFooter;
    }

    public boolean isEvaluateCellFormulas() {
        return evaluateCellFormulas;
    }

    public void setEvaluateCellFormulas(boolean evaluateCellFormulas) {
        this.evaluateCellFormulas = evaluateCellFormulas;
    }

    /**
     * exports data into a POI report. saves exported data into the File identified by the file variable
     *
     * @return a CompletableFuture containing the string of the file name (absolute path) of the generated file
     */
    public CompletableFuture<String> prepareMempoiReportToFile() {

        return CompletableFuture.supplyAsync(() -> {

            if (null == this.file) {
                throw new MempoiException("File export requested while target file not specified");
            }

            Strategos strategos = new Strategos(this.workbook, this.mempoiReportStyler, this.adjustColumnWidth, this.mempoiSubFooter, this.mempoiFooter, this.evaluateCellFormulas);
            return strategos.generateMempoiReportToFile(this.mempoiSheetList, this.file);
        });
    }


    /**
     * exports data into a POI report. return the byte array of the exported data
     *
     * @return a CompletableFuture containing the byte arrat of the exported data
     */
    public CompletableFuture<byte[]> prepareMempoiReportToByteArray() {

        return CompletableFuture.supplyAsync(() -> {

            Strategos strategos = new Strategos(this.workbook, this.mempoiReportStyler, this.adjustColumnWidth, this.mempoiSubFooter, this.mempoiFooter, this.evaluateCellFormulas);
            return strategos.generateMempoiReportToByteArray(this.mempoiSheetList);
        });
    }
}