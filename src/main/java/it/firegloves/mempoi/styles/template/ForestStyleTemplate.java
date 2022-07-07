package it.firegloves.mempoi.styles.template;

import org.apache.poi.ss.usermodel.IndexedColors;

public class ForestStyleTemplate extends HueStyleTemplate {

    private static final short TEXT_HEADER_AND_FOOTER_CELL_BG_COLOR_INDEX = IndexedColors.DARK_GREEN.getIndex();
    private static final short HEADER_CELL_BG_COLOR_INDEX = IndexedColors.GREEN.getIndex();
    private static final short COMMON_CELL_BG_COLOR_INDEX = IndexedColors.LIME.getIndex();
    private static final short SUB_FOOTER_CELL_BG_COLOR_INDEX = IndexedColors.SEA_GREEN.getIndex();

    public ForestStyleTemplate() {
        super();
        this.setSimpleTextHeaderCellBgColorIndex(TEXT_HEADER_AND_FOOTER_CELL_BG_COLOR_INDEX);
        this.setColsHeaderCellBgColorIndex(HEADER_CELL_BG_COLOR_INDEX);
        this.setCommonCellBgColorIndex(COMMON_CELL_BG_COLOR_INDEX);
        this.setSubFooterCellBgColorIndex(SUB_FOOTER_CELL_BG_COLOR_INDEX);
        this.setSimpleTextFooterCellBgColorIndex(TEXT_HEADER_AND_FOOTER_CELL_BG_COLOR_INDEX);
    }
}
