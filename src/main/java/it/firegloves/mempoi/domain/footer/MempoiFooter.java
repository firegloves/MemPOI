package it.firegloves.mempoi.domain.footer;

import org.apache.poi.ss.usermodel.Workbook;

public class MempoiFooter {

    protected Workbook workbook;

    protected String leftText;
    protected String centerText;
    protected String rightText;

    public MempoiFooter(Workbook workbook) {
        this.workbook = workbook;
    }

    public String getLeftText() {
        return leftText;
    }

    public void setLeftText(String leftText) {
        this.leftText = leftText;
    }

    public String getCenterText() {
        return centerText;
    }

    public void setCenterText(String centerText) {
        this.centerText = centerText;
    }

    public String getRightText() {
        return rightText;
    }

    public void setRightText(String rightText) {
        this.rightText = rightText;
    }
}
