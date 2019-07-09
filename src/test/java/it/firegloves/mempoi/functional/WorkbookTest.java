package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.util.concurrent.CompletableFuture;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.junit.Assert.assertEquals;

public class WorkbookTest extends FunctionalBaseTest {

    /**********************************************************************************************
     * HSSFWorkbook
     *********************************************************************************************/

    @Test
    public void testWithHSSFWorkbook() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_HSSFWorkbook.xlsx");

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(new HSSFWorkbook())
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void testWithHSSFWorkbookAndStylesAndSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_HSSFWorkbook_and_styles_and_subfooter.xlsx");
        HSSFWorkbook workbook = new HSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void testWithHSSFWorkbookAndStylesAndSubfooterAndEvaluateformulas() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_HSSFWorkbook_and_styles_and_subfooter_and_evaluateformulas.xlsx");
        HSSFWorkbook workbook = new HSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void testWithHSSFWorkbookAndCustomStyles() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_HSSFWorkbook_and_custom_styles.xlsx");
        HSSFWorkbook workbook = new HSSFWorkbook();

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
                    // custom header cell style
                    .withHeaderCellStyle(headerCellStyle)
                    // no style for number fields
                    .withNumberCellStyle(workbook.createCellStyle())
                    // custom date cell style
                    .withDateCellStyle(dateCellStyle)
                    // custom datetime cell style
                    .withDatetimeCellStyle(datetimeCellStyle)
                    // custom common data cell style
                    .withCommonDataCellStyle(commonDataCellStyle)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    /**********************************************************************************************
     * XSSFWorkbook
     *********************************************************************************************/

    @Test
    public void testWithXSSFWorkbook() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_XSSFWorkbook.xlsx");

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(new XSSFWorkbook())
                    .withFile(fileDest)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void testWithXSSFWorkbookAndStylesAndSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_XSSFWorkbook_and_styles_and_subfooter.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void ttestWithXSSFWorkbookAndStylesAndSubfooterAndEvaluateformulas() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_XSSFWorkbook_and_styles_and_subfooter_and_evaluateformulas.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void testWithXSSFWorkbookAndCustomStyles() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_XSSFWorkbook_and_custom_styles.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();

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
                    // custom header cell style
                    .withHeaderCellStyle(headerCellStyle)
                    // no style for number fields
                    .withNumberCellStyle(workbook.createCellStyle())
                    // custom date cell style
                    .withDateCellStyle(dateCellStyle)
                    // custom datetime cell style
                    .withDatetimeCellStyle(datetimeCellStyle)
                    // custom common data cell style
                    .withCommonDataCellStyle(commonDataCellStyle)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    /**********************************************************************************************
     * SXSSFWorkbook
     *********************************************************************************************/

    @Test
    public void testWithSXSSFWorkbook() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_SXSSFWorkbook.xlsx");

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(new SXSSFWorkbook())
                    .withFile(fileDest)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void testWithSXSSFWorkbookAndStylesAndSubfooter() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_SXSSFWorkbook_and_styles_and_subfooter.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void testWithSXSSFWorkbookAndStylesAndSubfooterAndEvaluateformulas() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_SXSSFWorkbook_and_styles_and_subfooter_and_evaluateformulas.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        try {

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new ForestStyleTemplate())
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withEvaluateCellFormulas(true)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    @Test
    public void testWithSXSSFWorkbookAndCustomStyles() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_SXSSFWorkbook_and_custom_styles.xlsx");
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
                    // custom header cell style
                    .withHeaderCellStyle(headerCellStyle)
                    // no style for number fields
                    .withNumberCellStyle(workbook.createCellStyle())
                    // custom date cell style
                    .withDateCellStyle(dateCellStyle)
                    // custom datetime cell style
                    .withDatetimeCellStyle(datetimeCellStyle)
                    // custom common data cell style
                    .withCommonDataCellStyle(commonDataCellStyle)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
