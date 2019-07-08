package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.sql.PreparedStatement;

public final class MempoiSheetBuilder {

    private String sheetName;
    private MempoiStyler sheetStyler;

    // style variables
    private Workbook workbook;
    private StyleTemplate styleTemplate;
    private CellStyle headerCellStyle;
    private CellStyle subFooterCellStyle;
    private CellStyle commonDataCellStyle;
    private CellStyle dateCellStyle;
    private CellStyle datetimeCellStyle;
    private CellStyle numberCellStyle;
    private MempoiFooter mempoiFooter;
    private MempoiSubFooter mempoiSubFooter;
    private PreparedStatement prepStmt;

    private MempoiSheetBuilder() {
    }

    public static MempoiSheetBuilder aMempoiSheet() {
        return new MempoiSheetBuilder();
    }

    public MempoiSheetBuilder withSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public MempoiSheetBuilder withSheetStyler(MempoiStyler sheetStyler) {
        this.sheetStyler = sheetStyler;
        return this;
    }

    public MempoiSheetBuilder withWorkbook(Workbook workbook) {
        this.workbook = workbook;
        return this;
    }

    public MempoiSheetBuilder withStyleTemplate(StyleTemplate styleTemplate) {
        this.styleTemplate = styleTemplate;
        return this;
    }

    public MempoiSheetBuilder withHeaderCellStyle(CellStyle headerCellStyle) {
        this.headerCellStyle = headerCellStyle;
        return this;
    }

    public MempoiSheetBuilder withSubFooterCellStyle(CellStyle subFooterCellStyle) {
        this.subFooterCellStyle = subFooterCellStyle;
        return this;
    }

    public MempoiSheetBuilder withCommonDataCellStyle(CellStyle commonDataCellStyle) {
        this.commonDataCellStyle = commonDataCellStyle;
        return this;
    }

    public MempoiSheetBuilder withDateCellStyle(CellStyle dateCellStyle) {
        this.dateCellStyle = dateCellStyle;
        return this;
    }

    public MempoiSheetBuilder withDatetimeCellStyle(CellStyle datetimeCellStyle) {
        this.datetimeCellStyle = datetimeCellStyle;
        return this;
    }

    public MempoiSheetBuilder withNumberCellStyle(CellStyle numberCellStyle) {
        this.numberCellStyle = numberCellStyle;
        return this;
    }

    public MempoiSheetBuilder withMempoiFooter(MempoiFooter mempoiFooter) {
        this.mempoiFooter = mempoiFooter;
        return this;
    }

    public MempoiSheetBuilder withMempoiSubFooter(MempoiSubFooter mempoiSubFooter) {
        this.mempoiSubFooter = mempoiSubFooter;
        return this;
    }

    public MempoiSheetBuilder withPrepStmt(PreparedStatement prepStmt) {
        this.prepStmt = prepStmt;
        return this;
    }

    public MempoiSheet build() {
        MempoiSheet mempoiSheet = new MempoiSheet(prepStmt);
        mempoiSheet.setSheetName(sheetName);
        mempoiSheet.setSheetStyler(sheetStyler);
        mempoiSheet.setWorkbook(workbook);
        mempoiSheet.setStyleTemplate(styleTemplate);
        mempoiSheet.setHeaderCellStyle(headerCellStyle);
        mempoiSheet.setSubFooterCellStyle(subFooterCellStyle);
        mempoiSheet.setCommonDataCellStyle(commonDataCellStyle);
        mempoiSheet.setDateCellStyle(dateCellStyle);
        mempoiSheet.setDatetimeCellStyle(datetimeCellStyle);
        mempoiSheet.setNumberCellStyle(numberCellStyle);
        mempoiSheet.setMempoiFooter(mempoiFooter);
        mempoiSheet.setMempoiSubFooter(mempoiSubFooter);
        return mempoiSheet;
    }
}
