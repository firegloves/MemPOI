package it.firegloves.mempoi.styles.template;

import org.apache.poi.ss.usermodel.IndexedColors;

public class PanegiriconStyleTemplate extends HueStyleTemplate {

    private static final short TEXT_HEADER_AND_FOOTER_CELL_BG_COLOR_INDEX = IndexedColors.DARK_TEAL.getIndex();
    private static final short HEADER_CELL_BG_COLOR_INDEX = IndexedColors.DARK_RED.getIndex();
    private static final short HEADER_FONT_COLOR_INDEX = IndexedColors.WHITE.getIndex();
    private static final short COMMON_CELL_BG_COLOR_INDEX = IndexedColors.TEAL.getIndex();
    private static final short COMMON_CELL_FONT_COLOR_INDEX = IndexedColors.WHITE.getIndex();
    private static final short SUB_FOOTER_CELL_BG_COLOR_INDEX = IndexedColors.RED.getIndex();

    public PanegiriconStyleTemplate() {
        super();
        this.setSimpleTextHeaderCellBgColorIndex(TEXT_HEADER_AND_FOOTER_CELL_BG_COLOR_INDEX);
        this.setColsHeaderCellBgColorIndex(HEADER_CELL_BG_COLOR_INDEX);
        this.setColsHeaderFontColorIndex(HEADER_FONT_COLOR_INDEX);
        this.setCommonCellBgColorIndex(COMMON_CELL_BG_COLOR_INDEX);
        this.setCommonFontColorIndex(COMMON_CELL_FONT_COLOR_INDEX);
        this.setSubFooterCellBgColorIndex(SUB_FOOTER_CELL_BG_COLOR_INDEX);
        this.setSimpleTextFooterCellBgColorIndex(TEXT_HEADER_AND_FOOTER_CELL_BG_COLOR_INDEX);
    }
}
