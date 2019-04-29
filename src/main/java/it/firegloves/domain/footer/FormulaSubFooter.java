/**
 * this sub footer shows an average for all the numeric columns in the report
 */

package it.firegloves.domain.footer;

import it.firegloves.domain.EExportDataType;
import it.firegloves.domain.MempoiColumn;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.util.List;


public abstract class FormulaSubFooter implements MempoiSubFooter {



    /**
     * the formula to apply to the subfooter cell
     */
    protected String formula;

    @Override
    public void setColumnSubFooter(Workbook workbook, List<MempoiColumn> mempoiColumnList, CellStyle subFooterCellStyle, int firstDataRowIndex, int lastDataRowIndex) {

        // customize subfooter cell style
        this.customizeSubFooterCellStyle(workbook, subFooterCellStyle);

        int len = mempoiColumnList.size();

        for (int i=0; i<len; i++) {

            MempoiColumn mc = mempoiColumnList.get(i);

            if (! mc.getColumnName().equalsIgnoreCase("id") && EExportDataType.NUMBER_STYLER_TYPES.contains(mc.getType())) {

                String colLetter = CellReference.convertNumToColString(i);
                mempoiColumnList.get(i).setSubFooterCell(new MempoiSubFooterCell(this.getFormula(colLetter, firstDataRowIndex, lastDataRowIndex), true, subFooterCellStyle));

            } else {
                mempoiColumnList.get(i).setSubFooterCell(new MempoiSubFooterCell(subFooterCellStyle));
            }
        }
    }


    /**
     * applies some customization to the SubFooter CellStyle
     * @param workbook
     * @param subFooterCellStyle
     */
    protected abstract void customizeSubFooterCellStyle(Workbook workbook, CellStyle subFooterCellStyle);

    /**
     * method that returns the formula to apply to the cell
     * @param colLetter the column letter to eventually add to the formula
     * @param firstDataRowIndex the first data row index (headers and subheaders are not counted)
     * @param lastDataRowIndex the last data row index (footers and subfooters are not counted)
     * @return the computed cell formula
     */
    protected abstract String getFormula(String colLetter, int firstDataRowIndex, int lastDataRowIndex);
}
