package it.firegloves.styles.template;

import org.apache.poi.ss.usermodel.*;

public class StandardStyleTemplate extends HueStyleTemplate {

    private static final short HEADER_CELL_BG_COLOR_INDEX = IndexedColors.GREEN.getIndex();
    private static final short HEADER_FONT_COLOR_INDEX = IndexedColors.BLACK.getIndex();
    private static final short SUB_FOOTER_CELL_BG_COLOR_INDEX = IndexedColors.SEA_GREEN.getIndex();

    public StandardStyleTemplate() {
        super();
        this.setHeaderCellBgColorIndex(HEADER_CELL_BG_COLOR_INDEX);
        this.setHeaderFontColorIndex(HEADER_FONT_COLOR_INDEX);
        this.setSubFooterCellBgColorIndex(SUB_FOOTER_CELL_BG_COLOR_INDEX);
    }

}
