package it.firegloves.mempoi.domain;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.exception.MempoiException;
import org.junit.Test;

public class MempoiReportTest {

    public static final String FILE = "mypath/myreport.xslx";
    public static final byte[] BYTES = "mypath/myreport.xslx".getBytes();
    public static final int SHEET_INDEX = 1;
    public static final MempoiSheetMetadata MEMPOI_SHEET_METADATA = new MempoiSheetMetadata()
            .setSheetIndex(SHEET_INDEX);

    @Test
    public void shouldCorrectlyCreateAMempoiReportWithFile() {

        MempoiReport mempoiReport = new MempoiReport()
                .setFile(FILE)
                .addMempoiSheetMetadata(SHEET_INDEX, MEMPOI_SHEET_METADATA);

        assertEquals(FILE, mempoiReport.getFile());
        assertEquals(1, mempoiReport.getMempoiSheetMetadataMap().size());
        assertEquals(MEMPOI_SHEET_METADATA, mempoiReport.getMempoiSheetMetadataMap().get(SHEET_INDEX));
    }

    @Test
    public void shouldCorrectlyCreateAMempoiReportWithByteArray() {

        MempoiReport mempoiReport = new MempoiReport()
                .setBytes(BYTES)
                .addMempoiSheetMetadata(SHEET_INDEX, MEMPOI_SHEET_METADATA);

        assertEquals(BYTES, mempoiReport.getBytes());
        assertEquals(1, mempoiReport.getMempoiSheetMetadataMap().size());
        assertEquals(MEMPOI_SHEET_METADATA, mempoiReport.getMempoiSheetMetadataMap().get(SHEET_INDEX));
    }

    @Test(expected = MempoiException.class)
    public void shouldThrowExceptionIfBothBytesAndFileArePopulated() {

        new MempoiReport()
                .setBytes(BYTES)
                .setFile(FILE);
    }

    @Test(expected = MempoiException.class)
    public void shouldThrowExceptionIfBothFileAndBytesArePopulated() {

        new MempoiReport()
                .setFile(FILE)
                .setBytes(BYTES);
    }

    @Test(expected = MempoiException.class)
    public void shouldThrowExceptionIfAddingAMempoiSheetMetadataWithIndexLowerThanZero() {

        new MempoiReport().addMempoiSheetMetadata(-1, MEMPOI_SHEET_METADATA);
    }

    @Test(expected = MempoiException.class)
    public void shouldThrowExceptionIfAddingANullMempoiSheetMetadata() {

        new MempoiReport().addMempoiSheetMetadata(SHEET_INDEX, null);
    }
}
