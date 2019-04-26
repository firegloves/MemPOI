/**
 * this sub footer shows the minimum vlaue for all the numeric columns in the report
 */

package it.firegloves.domain.footer;

public class NumberMinSubFooter extends NumberFormulaSubFooter {


    @Override
    protected String getFormula(String colLetter, int firstDataRowIndex, int lastDataRowIndex) {
        return "MIN(" + colLetter + firstDataRowIndex + ":" + colLetter + lastDataRowIndex + ")";
    }
}
