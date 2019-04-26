/**
 * this sub footer shows an average for all the numeric columns in the report
 */

package it.firegloves.domain.footer;

import it.firegloves.styles.template.StyleTemplate;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;


public abstract class NumberFormulaSubFooter extends FormulaSubFooter {

    @Override
    protected void customizeSubFooterCellStyle(Workbook workbook, CellStyle subFooterCellStyle) {

        subFooterCellStyle.setAlignment(HorizontalAlignment.RIGHT);
        subFooterCellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(StyleTemplate.STANDARD_NUMBER_FORMAT));
    }


}
