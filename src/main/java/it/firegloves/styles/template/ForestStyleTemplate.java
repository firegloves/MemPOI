package it.firegloves.styles.template;

import org.apache.poi.ss.usermodel.IndexedColors;

public class ForestStyleTemplate extends HueStyleTemplate {

    private static final short HEADER_CELL_BG_COLOR_INDEX = IndexedColors.GREEN.getIndex();
    private static final short COMMON_CELL_BG_COLOR_INDEX = IndexedColors.LIME.getIndex();
    private static final short SUB_FOOTER_CELL_BG_COLOR_INDEX = IndexedColors.SEA_GREEN.getIndex();

    public ForestStyleTemplate() {
        super();
        this.setHeaderCellBgColorIndex(HEADER_CELL_BG_COLOR_INDEX);
        this.setCommonCellBgColorIndex(COMMON_CELL_BG_COLOR_INDEX);
        this.setSubFooterCellBgColorIndex(SUB_FOOTER_CELL_BG_COLOR_INDEX);
    }
}
