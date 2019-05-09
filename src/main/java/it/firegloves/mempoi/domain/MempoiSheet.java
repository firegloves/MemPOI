package it.firegloves.mempoi.domain;

import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.styles.MempoiStyler;

import java.sql.PreparedStatement;
import java.util.Optional;

public class MempoiSheet {

    /**
     * the prepared statement to execute to take export data
     */
    private PreparedStatement prepStmt;

    /**
     * the sheet name
     */
    private String sheetName;

    /**
     * the report styler containing desired output styles for the current sheet
     */
    private MempoiStyler reportStyler;

    /**
     * the footer to apply to the sheet. if null => no footer is appended to the report
     */
    private MempoiFooter mempoiFooter;

    /**
     * the sub footer to apply to the sheet. if null => no sub footer is appended to the report
     */
    private MempoiSubFooter mempoiSubFooter;


    public MempoiSheet(PreparedStatement prepStmt) {
        this.prepStmt = prepStmt;
    }

    public MempoiSheet(PreparedStatement prepStmt, String sheetName) {
        this.prepStmt = prepStmt;
        this.sheetName = sheetName;
    }

    public MempoiSheet(PreparedStatement prepStmt, String sheetName, MempoiStyler reportStyler, MempoiFooter mempoiFooter, MempoiSubFooter mempoiSubFooter) {
        this.prepStmt = prepStmt;
        this.sheetName = sheetName;
        this.reportStyler = reportStyler;
        this.mempoiFooter = mempoiFooter;
        this.mempoiSubFooter = mempoiSubFooter;
    }

    public PreparedStatement getPrepStmt() {
        return prepStmt;
    }

    public void setPrepStmt(PreparedStatement prepStmt) {
        this.prepStmt = prepStmt;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Optional<MempoiStyler> getReportStyler() {
        return Optional.ofNullable(reportStyler);
    }

    public void setReportStyler(MempoiStyler reportStyler) {
        this.reportStyler = reportStyler;
    }

    public Optional<MempoiFooter> getMempoiFooter() {
        return Optional.ofNullable(mempoiFooter);
    }

    public void setMempoiFooter(MempoiFooter mempoiFooter) {
        this.mempoiFooter = mempoiFooter;
    }

    public Optional<MempoiSubFooter> getMempoiSubFooter() {
        return Optional.ofNullable(mempoiSubFooter);
    }

    public void setMempoiSubFooter(MempoiSubFooter mempoiSubFooter) {
        this.mempoiSubFooter = mempoiSubFooter;
    }
}
