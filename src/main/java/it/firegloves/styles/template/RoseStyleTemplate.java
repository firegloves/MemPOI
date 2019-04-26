package it.firegloves.styles.template;

import org.apache.poi.ss.usermodel.IndexedColors;

public class RoseStyleTemplate extends HueStyleTemplate {

    private static final short HEADER_CELL_BG_COLOR_INDEX = IndexedColors.PINK.getIndex();
    private static final short COMMON_CELL_BG_COLOR_INDEX = IndexedColors.ROSE.getIndex();
//    private static final short COMMON_CELL_FONT_COLOR_INDEX = IndexedColors.DARK_TEAL.getIndex();
    private static final short SUB_FOOTER_CELL_BG_COLOR_INDEX = IndexedColors.PINK1.getIndex();

    public RoseStyleTemplate() {
        super();
        this.setHeaderCellBgColorIndex(HEADER_CELL_BG_COLOR_INDEX);
        this.setCommonCellBgColorIndex(COMMON_CELL_BG_COLOR_INDEX);
//        this.setCommonFontColorIndex(COMMON_CELL_FONT_COLOR_INDEX);
        this.setSubFooterCellBgColorIndex(SUB_FOOTER_CELL_BG_COLOR_INDEX);
    }

}
