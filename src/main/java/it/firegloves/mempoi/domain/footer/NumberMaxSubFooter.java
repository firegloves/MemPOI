/**
 * this sub footer shows the maximum vlaue for all the numeric columns in the report
 */

package it.firegloves.mempoi.domain.footer;

public class NumberMaxSubFooter extends NumberFormulaSubFooter {


    @Override
    protected String getFormula(String colLetter, int firstDataRowIndex, int lastDataRowIndex) {
        return "MAX(" + colLetter + firstDataRowIndex + ":" + colLetter + lastDataRowIndex + ")";
    }
}
