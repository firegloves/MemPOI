/**
 * Strategy pattern for specific MempoiColumn logic
 */

package it.firegloves.mempoi.strategy.mempoicolumn;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

public interface MempoiColumnStrategy<T> {

    /**
     * receives a value of type Cell, analyzes it and collects informations
     * @param cell the Cell from which gain informations
     * @param value cell value of type T
     */
    void analyze(Cell cell, T value);

    /**
     * applies strategy logic to the Workbook modifying it
     * @param sheet the sheet on which operate
     */
    void execute(Sheet sheet);
}
