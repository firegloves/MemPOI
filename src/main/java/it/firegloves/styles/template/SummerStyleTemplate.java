package it.firegloves.styles.template;

import org.apache.poi.ss.usermodel.*;

public class SummerStyleTemplate extends HueStyleTemplate {

    private static final short HEADER_CELL_BG_COLOR_INDEX = IndexedColors.LIGHT_ORANGE.getIndex();
    private static final short HEADER_FONT_COLOR_INDEX = IndexedColors.BLACK.getIndex();
    private static final short COMMON_CELL_BG_COLOR_INDEX = IndexedColors.TAN.getIndex();
    private static final short SUB_FOOTER_CELL_BG_COLOR_INDEX = IndexedColors.GOLD.getIndex();

    public SummerStyleTemplate() {
        super();
        this.setHeaderCellBgColorIndex(HEADER_CELL_BG_COLOR_INDEX);
        this.setHeaderFontColorIndex(HEADER_FONT_COLOR_INDEX);
        this.setCommonCellBgColorIndex(COMMON_CELL_BG_COLOR_INDEX);
        this.setSubFooterCellBgColorIndex(SUB_FOOTER_CELL_BG_COLOR_INDEX);
        this.setSubFooterFontColorIndex(HEADER_FONT_COLOR_INDEX);
    }

}
