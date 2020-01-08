/**
 * this sub footer shows an average for all the numeric columns in the report
 */

package it.firegloves.mempoi.domain.footer;

public class NumberAverageSubFooter extends NumberFormulaSubFooter {

    @Override
    protected String getFormula(String colLetter, int firstDataRowIndex, int lastDataRowIndex) {
        return "AVERAGE(" + colLetter + firstDataRowIndex + ":" + colLetter + lastDataRowIndex + ")";
    }
}
