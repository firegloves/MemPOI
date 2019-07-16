package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.util.concurrent.CompletableFuture;

import static org.junit.Assert.*;

public class CommonTest extends FunctionalBaseTest {

    @Test
    public void testWithoutStyler() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_con_file.xlsx");

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            super.validateGeneratedFile(this.createStatement(), fut.get(), COLUMNS, HEADERS, null, new StandardStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void testWithByteArrayAndNumberStyler() {

        try {

            SXSSFWorkbook workbook = new SXSSFWorkbook();

            CellStyle numberCellStyle = workbook.createCellStyle();
            numberCellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("#.##0,00"));


            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withNumberCellStyle(numberCellStyle)
                    .build();

            CompletableFuture<byte[]> fut = memPOI.prepareMempoiReportToByteArray();
            assertNotNull("not null byte array", fut.get());
            assertNotEquals("not empty byte array", 0, fut.get().length);

            File destFile = new File(this.outReportFolder.getAbsolutePath(), "test_with_byte_array_and_number_styler.xlsx");
            try (FileOutputStream fos = new FileOutputStream(destFile)) {
                fos.write(fut.get());
            }

            super.validateGeneratedFile(this.createStatement(), destFile.getAbsolutePath(), COLUMNS, HEADERS, null, null);

            // TODO add header overriden style check

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void testWithByteArray() {

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();

            CompletableFuture<byte[]> fut = memPOI.prepareMempoiReportToByteArray();
            assertNotNull("not null byte array", fut.get());
            assertNotEquals("not empty byte array", 0, fut.get().length);

            File destFile = new File(this.outReportFolder.getAbsolutePath(), "test_with_byte_array_and_number_styler.xlsx");
            try (FileOutputStream fos = new FileOutputStream(destFile)) {
                fos.write(fut.get());
            }

            super.validateGeneratedFile(this.createStatement(), destFile.getAbsolutePath(), COLUMNS, HEADERS, null, new StandardStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void testWithMultipleSheets() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_multiple_sheets.xlsx");

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt, "Dogs sheet"))
                    .addMempoiSheet(new MempoiSheet(conn.prepareStatement(super.createQuery(COLUMNS_2, HEADERS_2, NO_LIMITS)), "Cats sheet"))
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            // validate first sheet
            super.validateGeneratedFile(this.createStatement(), fut.get(), COLUMNS, HEADERS, null, new StandardStyleTemplate());
            // validate second sheet
            super.validateSecondPrepStmtSheet(conn.prepareStatement(super.createQuery(COLUMNS_2, HEADERS_2, NO_LIMITS)), fut.get(), 1, HEADERS_2, true, new StandardStyleTemplate());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void testWithFileAndStyles() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_file_and_styles.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
            headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle dateCellStyle = workbook.createCellStyle();
            dateCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
            dateCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            dateCellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy/MM/dd"));

            CellStyle datetimeCellStyle = workbook.createCellStyle();
            datetimeCellStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
            datetimeCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            datetimeCellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy/mm/dd hh:mm:ss"));

            CellStyle commonDataCellStyle = workbook.createCellStyle();
            commonDataCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            commonDataCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withHeaderCellStyle(headerCellStyle)
                    .withNumberCellStyle(workbook.createCellStyle())
                    .withDateCellStyle(dateCellStyle)
                    .withDatetimeCellStyle(datetimeCellStyle)
                    .withCommonDataCellStyle(commonDataCellStyle)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            // TODO validates overriden styles

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
