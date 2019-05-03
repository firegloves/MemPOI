package it.firegloves.mempoi.styles.template;

import org.apache.poi.ss.usermodel.IndexedColors;

public class PurpleStyleTemplate extends HueStyleTemplate {

    private static final short HEADER_CELL_BG_COLOR_INDEX = IndexedColors.VIOLET.getIndex();
    private static final short COMMON_CELL_BG_COLOR_INDEX = IndexedColors.LAVENDER.getIndex();
//    private static final short COMMON_CELL_FONT_COLOR_INDEX = IndexedColors.DARK_TEAL.getIndex();
    private static final short SUB_FOOTER_CELL_BG_COLOR_INDEX = IndexedColors.PLUM.getIndex();

    public PurpleStyleTemplate() {
        super();
        this.setHeaderCellBgColorIndex(HEADER_CELL_BG_COLOR_INDEX);
        this.setCommonCellBgColorIndex(COMMON_CELL_BG_COLOR_INDEX);
//        this.setCommonFontColorIndex(COMMON_CELL_FONT_COLOR_INDEX);
        this.setSubFooterCellBgColorIndex(SUB_FOOTER_CELL_BG_COLOR_INDEX);
    }

}
