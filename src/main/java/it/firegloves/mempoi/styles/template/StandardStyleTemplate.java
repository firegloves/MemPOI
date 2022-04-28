package it.firegloves.mempoi.styles.template;

import org.apache.poi.ss.usermodel.IndexedColors;

public class StandardStyleTemplate extends HueStyleTemplate {

    private static final short TEXT_HEADER_CELL_BG_COLOR_INDEX = IndexedColors.DARK_GREEN.getIndex();
    private static final short HEADER_CELL_BG_COLOR_INDEX = IndexedColors.GREEN.getIndex();
    private static final short HEADER_FONT_COLOR_INDEX = IndexedColors.BLACK.getIndex();
    private static final short SUB_FOOTER_CELL_BG_COLOR_INDEX = IndexedColors.SEA_GREEN.getIndex();

    public StandardStyleTemplate() {
        super();
        this.setSimpleTextHeaderCellBgColorIndex(TEXT_HEADER_CELL_BG_COLOR_INDEX);
        this.setColsHeaderCellBgColorIndex(HEADER_CELL_BG_COLOR_INDEX);
        this.setColsHeaderFontColorIndex(HEADER_FONT_COLOR_INDEX);
        this.setSubFooterCellBgColorIndex(SUB_FOOTER_CELL_BG_COLOR_INDEX);
    }

}
