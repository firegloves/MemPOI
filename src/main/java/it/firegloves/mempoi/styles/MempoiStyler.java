package it.firegloves.mempoi.styles;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.CellStyle;

@Data
@Accessors(chain = true)
@AllArgsConstructor
@NoArgsConstructor
public class MempoiStyler {

    // simple text header
    private CellStyle simpleTextHeaderCellStyle;

    // cols header
    private CellStyle colsHeaderCellStyle;

    // common data
    private CellStyle commonDataCellStyle;

    // date
    private CellStyle dateCellStyle;

    // date time
    private CellStyle datetimeCellStyle;

    // integer number
    private CellStyle integerCellStyle;

    // floating point number
    private CellStyle floatingPointCellStyle;

    // sub footer
    private CellStyle subFooterCellStyle;
}
