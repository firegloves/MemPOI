package it.firegloves.mempoi.styles.template;

import org.apache.poi.ss.usermodel.IndexedColors;

public class SummerStyleTemplate extends HueStyleTemplate {

    private static final short TEXT_HEADER_AND_FOOTER_CELL_BG_COLOR_INDEX = IndexedColors.BROWN.getIndex();
    private static final short HEADER_CELL_BG_COLOR_INDEX = IndexedColors.LIGHT_ORANGE.getIndex();
    private static final short HEADER_FONT_COLOR_INDEX = IndexedColors.BLACK.getIndex();
    private static final short COMMON_CELL_BG_COLOR_INDEX = IndexedColors.TAN.getIndex();
    private static final short SUB_FOOTER_CELL_BG_COLOR_INDEX = IndexedColors.GOLD.getIndex();

    public SummerStyleTemplate() {
        super();
        this.setSimpleTextHeaderCellBgColorIndex(TEXT_HEADER_AND_FOOTER_CELL_BG_COLOR_INDEX);
        this.setColsHeaderCellBgColorIndex(HEADER_CELL_BG_COLOR_INDEX);
        this.setColsHeaderFontColorIndex(HEADER_FONT_COLOR_INDEX);
        this.setCommonCellBgColorIndex(COMMON_CELL_BG_COLOR_INDEX);
        this.setSubFooterCellBgColorIndex(SUB_FOOTER_CELL_BG_COLOR_INDEX);
        this.setSubFooterFontColorIndex(HEADER_FONT_COLOR_INDEX);
        this.setSimpleTextFooterCellBgColorIndex(TEXT_HEADER_AND_FOOTER_CELL_BG_COLOR_INDEX);
    }

}
