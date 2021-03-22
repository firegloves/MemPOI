/**
 * this class is responsible to map mempoi columns to the desired cell styles
 */

package it.firegloves.mempoi.styles;

import it.firegloves.mempoi.domain.EExportDataType;
import it.firegloves.mempoi.domain.MempoiColumn;
import java.util.Arrays;
import java.util.Collections;
import java.util.EnumSet;
import java.util.HashMap;
import java.util.List;
import java.util.Set;
import org.apache.poi.ss.usermodel.CellStyle;

public class MempoiColumnStyleManager {

    private static final Set<EExportDataType> dateStylerTypes = EnumSet.of(EExportDataType.DATE);
    private static final Set<EExportDataType> datetimeStylerTypes = EnumSet
            .of(EExportDataType.TIME, EExportDataType.TIMESTAMP);
    private static final Set<EExportDataType> integerStylerTypes = EnumSet
            .of(EExportDataType.INT, EExportDataType.BIG_INTEGER);
    private static final Set<EExportDataType> floatingPointStylerTypes = EnumSet
            .of(EExportDataType.DOUBLE, EExportDataType.FLOAT);
    private final List<Set<EExportDataType>> stylerTypesSet = Arrays
            .asList(dateStylerTypes, datetimeStylerTypes, integerStylerTypes, floatingPointStylerTypes);

    private HashMap<Set<EExportDataType>, CellStyle> cellStylerMap;

    private final MempoiStyler mempoiStyler;


    public MempoiColumnStyleManager(MempoiStyler mempoiStyler) {
        this.mempoiStyler = mempoiStyler;
        this.initCellStyleMap();
    }

    /**
     * populates strategy pattern's map to apply a style to the cells
     */
    private void initCellStyleMap() {
        this.cellStylerMap = new HashMap<>();
        this.cellStylerMap.put(dateStylerTypes, this.mempoiStyler.getDateCellStyle());
        this.cellStylerMap.put(datetimeStylerTypes, this.mempoiStyler.getDatetimeCellStyle());
        this.cellStylerMap.put(floatingPointStylerTypes, this.mempoiStyler.getFloatingPointCellStyle());
        this.cellStylerMap.put(integerStylerTypes, this.mempoiStyler.getIntegerCellStyle());
    }


    /**
     * @param type
     * @return the corresponding EnumSet<EExportDataType>
     */
    private Set<EExportDataType> getEExportDataTypeEnumSet(EExportDataType type) {

        return stylerTypesSet.stream().filter(e -> e.contains(type)).findFirst().orElseGet(Collections::emptySet);
    }


    /**
     * populate the received MempoiColumn list with the corresponding CellStyle
     *
     * @param mempoiColumnList MempoiColumn list to which apply CellStyle
     */
    public MempoiColumnStyleManager setMempoiColumnListStyler(List<MempoiColumn> mempoiColumnList) {

        mempoiColumnList.forEach(mc -> {

            // if column name == 'id' -> no style
            if (mc.getColumnName().equalsIgnoreCase("id")) {
                mc.setCellStyle(this.mempoiStyler.getIntegerCellStyle());
            } else {
                // else default style
                mc.setCellStyle(
                        this.cellStylerMap.getOrDefault(this.getEExportDataTypeEnumSet(mc.getType()), this.mempoiStyler
                                .getCommonDataCellStyle()));
            }
        });

        return this;
    }

    /**
     * set the CellStyle of the received MempoiColumn using the received EExportDataType to identify the needed one
     * this method is used to granularly override the MempoiColumn cell style starting by an EExportDataType
     *
     * @param mempoiColumn the MempoiColumn on which set the cell style
     * @param exportDataType the EExportDataType to use to identify the required style
     */
    public void setMempoiColumnCellStyle(MempoiColumn mempoiColumn, EExportDataType exportDataType) {

        // else default style
        mempoiColumn.setCellStyle(
                this.cellStylerMap.getOrDefault(
                        this.getEExportDataTypeEnumSet(exportDataType),
                        this.mempoiStyler.getCommonDataCellStyle()));
    }

    /**
     * check if the received EExportDataType belongs to the numeric list
     *
     * @param type the EExportDataType to check
     * @return true if the received EExportDataType belongs to the numeric list, false otherwise
     */
    public static boolean isNumericType(EExportDataType type) {
        return floatingPointStylerTypes.contains(type) || integerStylerTypes.contains(type);
    }
}
