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

    // header
    private CellStyle headerCellStyle;

    // common data
    private CellStyle commonDataCellStyle;

    // date
    private CellStyle dateCellStyle;

    // date time
    private CellStyle datetimeCellStyle;

    // number
    private CellStyle numberCellStyle;

    // sub footer
    private CellStyle subFooterCellStyle;
}
