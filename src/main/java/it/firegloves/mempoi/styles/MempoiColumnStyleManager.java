/**
 * this class is responsible to map mempoi columns to the desired cell styles
 */

package it.firegloves.mempoi.styles;

import it.firegloves.mempoi.domain.EExportDataType;
import it.firegloves.mempoi.domain.MempoiColumn;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.EnumSet;
import java.util.HashMap;
import java.util.List;

public class MempoiColumnStyleManager {

    public static final EnumSet<EExportDataType> DATE_STYLER_TYPES = EnumSet.of(EExportDataType.DATE);
    public static final EnumSet<EExportDataType> DATETIME_STYLER_TYPES = EnumSet.of(EExportDataType.TIME, EExportDataType.TIMESTAMP);
    public static final EnumSet<EExportDataType> NUMBER_STYLER_TYPES = EnumSet.of(EExportDataType.INT, EExportDataType.DOUBLE, EExportDataType.FLOAT);

    private HashMap<EnumSet<EExportDataType>, CellStyle> cellStylerMap;

    private MempoiStyler reportStyler;

    public MempoiColumnStyleManager(MempoiStyler reportStyler) {
        this.reportStyler = reportStyler;
        this.initCellStyleMap();
    }

    /**
     * populates strategy pattern's map to apply a style to the cells
     */
    private void initCellStyleMap() {
        this.cellStylerMap = new HashMap<>();
        this.cellStylerMap.put(DATE_STYLER_TYPES, this.reportStyler.getDateCellStyle());
        this.cellStylerMap.put(DATETIME_STYLER_TYPES, this.reportStyler.getDatetimeCellStyle());
        this.cellStylerMap.put(NUMBER_STYLER_TYPES, this.reportStyler.getNumberCellStyle());
    }


    /**
     * @param type
     * @return the corresponding EnumSet<EExportDataType>
     */
    private EnumSet<EExportDataType> getEExportDataTypeEnumSet(EExportDataType type) {

        // TODO refactor

        if (DATE_STYLER_TYPES.contains(type)) {
            return DATE_STYLER_TYPES;
        } else if (DATETIME_STYLER_TYPES.contains(type)) {
            return DATETIME_STYLER_TYPES;
        } else if (NUMBER_STYLER_TYPES.contains(type)) {
            return NUMBER_STYLER_TYPES;
        } else {
            return null;
        }
    }


    /**
     * populate the received MempoiColumn list with the corresponding CellStyle
     *
     * @param mempoiColumnList MempoiColumn list to which apply CellStyle
     */
    public void setMempoiColumnListStyler(List<MempoiColumn> mempoiColumnList) {

        mempoiColumnList.stream().forEach(mc -> {

            // if column name == 'id' => no style
            if (mc.getColumnName().equalsIgnoreCase("id")) {
                mc.setCellStyle(this.reportStyler.getCommonDataCellStyle());
            } else {
                // else default style
                mc.setCellStyle(this.cellStylerMap.getOrDefault(this.getEExportDataTypeEnumSet(mc.getType()), this.reportStyler.getCommonDataCellStyle()));
            }
        });
    }

}