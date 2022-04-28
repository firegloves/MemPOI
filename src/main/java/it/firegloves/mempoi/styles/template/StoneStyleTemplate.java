package it.firegloves.mempoi.styles.template;

import org.apache.poi.ss.usermodel.IndexedColors;

public class StoneStyleTemplate extends HueStyleTemplate {

    private static final short TEXT_HEADER_CELL_BG_COLOR_INDEX = IndexedColors.GREY_40_PERCENT.getIndex();
    private static final short TEXT_HEADER_FONT_COLOR_INDEX = IndexedColors.BLACK.getIndex();
    private static final short HEADER_CELL_BG_COLOR_INDEX = IndexedColors.GREY_80_PERCENT.getIndex();
    private static final short COMMON_CELL_BG_COLOR_INDEX = IndexedColors.GREY_25_PERCENT.getIndex();
    private static final short SUB_FOOTER_CELL_BG_COLOR_INDEX = IndexedColors.GREY_50_PERCENT.getIndex();

    public StoneStyleTemplate() {
        super();
        this.setSimpleTextHeaderCellBgColorIndex(TEXT_HEADER_CELL_BG_COLOR_INDEX);
        this.setSimpleTextHeaderFontColorIndex(TEXT_HEADER_FONT_COLOR_INDEX);
        this.setColsHeaderCellBgColorIndex(HEADER_CELL_BG_COLOR_INDEX);
        this.setCommonCellBgColorIndex(COMMON_CELL_BG_COLOR_INDEX);
        this.setSubFooterCellBgColorIndex(SUB_FOOTER_CELL_BG_COLOR_INDEX);
    }
}
