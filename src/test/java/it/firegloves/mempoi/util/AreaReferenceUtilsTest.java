package it.firegloves.mempoi.util;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.TestHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

public class AreaReferenceUtilsTest {

    @Test
    public void shouldSkipTheFirstRow() {
        final Workbook wb = new SXSSFWorkbook();
        final AreaReference areaReference = new AreaReference(TestHelper.AREA_REFERENCE, wb.getSpreadsheetVersion());
        final AreaReference current = AreaReferenceUtils.skipFirstRow(areaReference, wb);

        assertEquals(current.formatAsString(), "A2:F6");
    }

    @Test
    public void shouldSkipTheExpectedRowsAndCols() {
        final Workbook wb = new SXSSFWorkbook();
        final AreaReference areaReference = new AreaReference(TestHelper.AREA_REFERENCE, wb.getSpreadsheetVersion());
        final AreaReference current = AreaReferenceUtils.skipRowsAndCols(2, 4, areaReference, wb);

        assertEquals("E3:J6", current.formatAsString());
    }

    @Test
    public void shouldNotSkipWhenReceiving0AsOffsets() {
        final Workbook wb = new SXSSFWorkbook();
        final AreaReference areaReference = new AreaReference(TestHelper.AREA_REFERENCE, wb.getSpreadsheetVersion());
        final AreaReference current = AreaReferenceUtils.skipRowsAndCols(0, 0, areaReference, wb);

        assertEquals(current.formatAsString(), "A1:F6");
    }

    @Test
    public void shouldNotSkipWhenReceivingNegativeValueAsOffsets() {
        final Workbook wb = new SXSSFWorkbook();
        final AreaReference areaReference = new AreaReference(TestHelper.AREA_REFERENCE, wb.getSpreadsheetVersion());
        final AreaReference current = AreaReferenceUtils.skipRowsAndCols(-1, -1, areaReference, wb);

        assertEquals(current.formatAsString(), "A1:F6");
    }

    @Test(expected = MempoiException.class)
    public void shouldThrowTheExpectedExceptionWhileSkippingAnInvalidAreaReference() {
        AreaReferenceUtils.skipRowsAndCols(1, 1, null, null);
    }

    @Test(expected = MempoiException.class)
    public void shouldThrowTheExpectedExceptionWhileSkippingAnInvalidWorkbook() {
        final Workbook wb = new SXSSFWorkbook();
        final AreaReference areaReference = new AreaReference(TestHelper.AREA_REFERENCE, wb.getSpreadsheetVersion());
        AreaReferenceUtils.skipRowsAndCols(1, 1, areaReference, null);
    }

}
