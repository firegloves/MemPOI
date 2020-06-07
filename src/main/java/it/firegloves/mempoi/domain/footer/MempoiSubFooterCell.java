package it.firegloves.mempoi.domain.footer;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.CellStyle;

@Data
@Accessors(chain = true)
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

}