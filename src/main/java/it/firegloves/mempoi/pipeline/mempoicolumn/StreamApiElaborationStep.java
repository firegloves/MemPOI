/**
 * pipeline pattern's step for elaborating MempoiColumn's into a HSSF or XSSF
 */

package it.firegloves.mempoi.pipeline.mempoicolumn;

import it.firegloves.mempoi.pipeline.mempoicolumn.abstractfactory.MempoiColumnElaborationStep;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public abstract class StreamApiElaborationStep implements MempoiColumnElaborationStep {

    /**
     * reference to the SXSSFWorkbook that is being generated
     */
    protected SXSSFWorkbook workbook;



}