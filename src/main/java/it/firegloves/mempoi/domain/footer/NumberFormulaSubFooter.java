/**
 * this sub footer shows an average for all the numeric columns in the report
 */

package it.firegloves.mempoi.domain.footer;

import it.firegloves.mempoi.styles.StandardDataFormat;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;


public abstract class NumberFormulaSubFooter extends FormulaSubFooter {

    @Override
    protected void customizeSubFooterCellStyle(Workbook workbook, CellStyle subFooterCellStyle) {

        subFooterCellStyle.setAlignment(HorizontalAlignment.RIGHT);
        subFooterCellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(StandardDataFormat.STANDARD_FLOATING_NUMBER_FORMAT.getFormat()));
    }


}
