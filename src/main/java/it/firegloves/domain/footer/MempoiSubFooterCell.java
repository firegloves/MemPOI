package it.firegloves.domain.footer;

import org.apache.poi.ss.usermodel.CellStyle;

public class MempoiSubFooterCell {

    protected String colLetter;
    protected String value;
    protected boolean cellFormula;
    protected CellStyle style;

    public MempoiSubFooterCell() {
    }

    public MempoiSubFooterCell(CellStyle style) {
        this.style = style;
    }

    public MempoiSubFooterCell(String value, boolean cellFormula, CellStyle style) {
        this.value = value;
        this.cellFormula = cellFormula;
        this.style = style;
    }

    public String getColLetter() {
        return colLetter;
    }

    public void setColLetter(String colLetter) {
        this.colLetter = colLetter;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public boolean isCellFormula() {
        return cellFormula;
    }

    public void setCellFormula(boolean cellFormula) {
        this.cellFormula = cellFormula;
    }

    public CellStyle getStyle() {
        return style;
    }

    public void setStyle(CellStyle style) {
        this.style = style;
    }
}
