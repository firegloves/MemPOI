package it.firegloves.mempoi.domain.pivottable;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

@Data
@Accessors(chain = true)
public class MempoiPivotTable {

    private final Workbook workbook;
    private final MempoiPivotTableSource source;
    private final CellReference position;
    private final String[] rowLabelColumns;
    private final String[] columnLabelColumns;
    private final String[] reportFilterColumns;

}
