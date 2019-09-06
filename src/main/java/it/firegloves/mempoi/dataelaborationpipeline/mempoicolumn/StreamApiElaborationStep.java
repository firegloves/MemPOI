/**
 * pipeline pattern's step for elaborating MempoiColumn's into a HSSF or XSSF
 */

package it.firegloves.mempoi.dataelaborationpipeline.mempoicolumn;

import it.firegloves.mempoi.Strategos;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;

public abstract class StreamApiElaborationStep<T> implements MempoiColumnElaborationStep<T> {

    private static final Logger logger = LoggerFactory.getLogger(Strategos.class);

    /**
     * reference to the SXSSFWorkbook that is being generated
     */
    protected SXSSFWorkbook workbook;

    /**
     * this variable contains the SXSSFWorkbook in memory max rows number
     */
    protected int maxRowsNum;



    public StreamApiElaborationStep(SXSSFWorkbook workbook) {
        this.workbook = workbook;
        this.maxRowsNum = this.workbook.getRandomAccessWindowSize();
    }

    /**
     * if workbook's randomAccessWindowSize == -1 => manages rows flush
     * @param sheet
     *
     * @return true if rows have been flushed
     */
    protected boolean manageFlush(SXSSFSheet sheet) {

        if (this.maxRowsNum == -1) {

            try {
                sheet.flushRows(1);     // 1 remaining because of analysis => previous row is the last one, the current one belongs to the new set
                logger.debug("SXSSFSheet in memory rows flushed");
                return true;
            } catch (IOException e) {
                throw new MempoiRuntimeException(e);
            }
        }

        return false;
    }
}