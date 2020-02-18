package it.firegloves.mempoi.domain.footer;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.Workbook;

@Data
@Accessors(chain = true)
public class MempoiFooter {

    protected Workbook workbook;

    protected String leftText;
    protected String centerText;
    protected String rightText;

    public MempoiFooter(Workbook workbook) {
        this.workbook = workbook;
    }
}