package it.firegloves.mempoi.domain.footer;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Workbook;

public class StandardMempoiFooter extends MempoiFooter {

    public final String MEMPOI_URL = "https://github.com/firegloves/MemPOI";

    protected Hyperlink mempoiLink;

    public StandardMempoiFooter(Workbook workbook, String centerText) {
        super(workbook);
        super.leftText = "Powered by MemPOI";
        this.setMempoiLink();
        super.centerText = centerText;
        super.rightText = "Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages();
    }

    public void setMempoiLink() {
        this.mempoiLink = this.workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
        this.mempoiLink.setAddress(MEMPOI_URL);
    }
}
