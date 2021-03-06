/**
 * pipeline pattern's step for elaborating MempoiColumn's generated data after workbook generation
 */

package it.firegloves.mempoi.datapostelaboration.mempoicolumn;

import it.firegloves.mempoi.domain.MempoiSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

public interface MempoiColumnElaborationStep<T> {

    /**
     * receives a value of type Cell, analyzes it and collects information
     *
     * @param cell  the Cell from which gain informations
     * @param value cell value of type T
     */
    void performAnalysis(Cell cell, T value);

    /**
     * closes the analysis, often used to manage the point that the ResultSet was already full iterated
     *
     * @param currRowNum current row num
     */
    void closeAnalysis(int currRowNum);

    /**
     * applies step logic to the Workbook modifying it
     *
     * @param mempoiSheet the MempoiSheet from which gain information
     * @param workbook the Workbook from which get Sheet
     */
    void execute(MempoiSheet mempoiSheet, Workbook workbook);
}
