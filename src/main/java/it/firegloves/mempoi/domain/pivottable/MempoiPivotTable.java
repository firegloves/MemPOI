package it.firegloves.mempoi.domain.pivottable;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Table;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;

import java.util.List;
import java.util.Map;

@Data
@Accessors(chain = true)
public class MempoiPivotTable {

    private final Workbook workbook;
    private final MempoiPivotTableSource source;
    private final CellReference position;
    private final List<String> rowLabelColumns;
    private final Map<DataConsolidateFunction, List<String>> columnLabelColumns;
    private final List<String> reportFilterColumns;

    /**
     * reference to the XSSFPivotTable generated with the current MempoiPivotTable
     * DON'T POPULATE IT MANUALLY
     */
    private XSSFPivotTable pivotTable;
}
