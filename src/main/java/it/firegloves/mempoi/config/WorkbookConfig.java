package it.firegloves.mempoi.config;

import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.FormulaSubFooter;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.styles.MempoiStyler;
import org.apache.poi.ss.usermodel.Workbook;
import java.io.File;
import java.util.List;

public class WorkbookConfig {

    /**
     * the sub footer fallback to apply to the report. if null => no sub footer is appended to the report.  you could also specify a particular MempoiSubFooter for the singles sheets
     */
    private MempoiSubFooter mempoiSubFooter;

    /**
     * the footer fallback to apply to the report. if null => no footer is appended to the report.  you could also specify a particular MempoiFooter for the singles sheets
     */
    private MempoiFooter mempoiFooter;

    /**
     * the workbook fallback report styler containing desired output styles. you could also specify a particular reportStyler for the singles sheets
     */
    private MempoiStyler reportStyler;

    /**
     * the workbook to create the report
     */
    private Workbook workbook;

    /**
     * the optional file where to save the report
     */
    private File file;

    /**
     * if true mempoi tries to adjust all columns width accordingly with column data cells length
     */
    private boolean adjustColSize;

    /**
     * true if come cell formulas have to be evaluated
     * this true condition implies that:
     *  - if the workbook is a SXSSFWorkbook
     *    - the workbook is written on a temp file
     *    - the file is reopenened
     *    - the cell formulas are evaluated
     *    - the workbook is saved again on the final file
     * otherwise the workbook is written normally
     */
    private boolean hasFormulasToEvaluate;

    /**
     * by default MemPOI forces Excel to evaluate cell formulas when it opens the report
     * but if this var is true MemPOI tries to evaluate cell formulas at runtime instead
     */
    private boolean evaluateCellFormulas;


    /**
     * the MempoiSheet list equal to the sheet to create in the report
     */
    private List<MempoiSheet> sheetList;


    public WorkbookConfig() {
    }

    public WorkbookConfig(MempoiSubFooter mempoiSubFooter, MempoiFooter mempoiFooter, MempoiStyler reportStyler, Workbook workbook, boolean adjustColSize, boolean evaluateCellFormulas, List<MempoiSheet> sheetList) {
        this.mempoiSubFooter = mempoiSubFooter;
        this.mempoiFooter = mempoiFooter;
        this.reportStyler = reportStyler;
        this.workbook = workbook;
        this.adjustColSize = adjustColSize;
        this.evaluateCellFormulas = evaluateCellFormulas;
        this.sheetList = sheetList;
    }

    public WorkbookConfig(MempoiSubFooter mempoiSubFooter, MempoiFooter mempoiFooter, MempoiStyler reportStyler, Workbook workbook, boolean adjustColSize, boolean evaluateCellFormulas, List<MempoiSheet> sheetList, File file) {
        this.mempoiSubFooter = mempoiSubFooter;
        this.mempoiFooter = mempoiFooter;
        this.reportStyler = reportStyler;
        this.workbook = workbook;
        this.adjustColSize = adjustColSize;
        this.evaluateCellFormulas = evaluateCellFormulas;
        this.sheetList = sheetList;
        this.file = file;
    }

    /**
     * defines if the current workbook has to process some cell formulas
     */
    private void setHasFormulasToEvaluate() {
        this.hasFormulasToEvaluate = null != this.mempoiSubFooter && this.mempoiSubFooter instanceof FormulaSubFooter;
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

    public MempoiStyler getReportStyler() {
        return reportStyler;
    }

    public void setReportStyler(MempoiStyler reportStyler) {
        this.reportStyler = reportStyler;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public boolean isAdjustColSize() {
        return adjustColSize;
    }

    public void setAdjustColSize(boolean adjustColSize) {
        this.adjustColSize = adjustColSize;
    }

    public boolean isHasFormulasToEvaluate() {
        return hasFormulasToEvaluate;
    }

    public void setHasFormulasToEvaluate(boolean hasFormulasToEvaluate) {
        this.hasFormulasToEvaluate = hasFormulasToEvaluate;
    }

    public List<MempoiSheet> getSheetList() {
        return sheetList;
    }

    public void setSheetList(List<MempoiSheet> sheetList) {
        this.sheetList = sheetList;
    }

    public boolean isEvaluateCellFormulas() {
        return evaluateCellFormulas;
    }

    public void setEvaluateCellFormulas(boolean evaluateCellFormulas) {
        this.evaluateCellFormulas = evaluateCellFormulas;
        this.setHasFormulasToEvaluate();
    }

    public File getFile() {
        return file;
    }

    public void setFile(File file) {
        this.file = file;
    }
}
