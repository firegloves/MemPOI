package it.firegloves.mempoi.unit;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.NumberMinSubFooter;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.domain.footer.StandardMempoiFooter;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.StoneStyleTemplate;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.io.File;
import java.sql.PreparedStatement;

import static org.junit.Assert.*;

public class MempoiBuilderTest {

   @Mock
   private PreparedStatement prepStmt;

   @Before
   public void prepare() {
      MockitoAnnotations.initMocks(this);
   }

   @Test
   public void mempoiBuilderFullPopulated() {

      File fileDest = new File("file.xlsx");
      SXSSFWorkbook workbook = new SXSSFWorkbook();

      CellStyle headerCellStyle = workbook.createCellStyle();
      headerCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
      headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

      MemPOI memPOI = new MempoiBuilder()
              .setWorkbook(workbook)
              .setFile(fileDest)
              .setAdjustColumnWidth(true)
              .addMempoiSheet(new MempoiSheet(prepStmt, "test name"))
              .setStyleTemplate(new ForestStyleTemplate())
              .setHeaderCellStyle(headerCellStyle)
              .setMempoiSubFooter(new NumberSumSubFooter())
              .setEvaluateCellFormulas(true)
              .build();

      assertNotNull("MemPOIBuilder returns a not null MemPOI", memPOI);
      assertNotNull("MemPOI workbookconfig not null", memPOI.getWorkbookConfig());
      assertNotNull("MemPOI file not null", memPOI.getWorkbookConfig().getFile());
      assertTrue("MemPOI adjustColumnWidth true ", memPOI.getWorkbookConfig().isAdjustColSize());
      assertNotNull("MemPOI workbook not null", memPOI.getWorkbookConfig().getWorkbook());
      assertNotNull("MemPOI mempoiSheetList not null", memPOI.getWorkbookConfig().getSheetList());
      assertEquals("MemPOI mempoiSheetList size 1", 1, memPOI.getWorkbookConfig().getSheetList().size());
      assertNotNull("MemPOI first mempoiSheet not null", memPOI.getWorkbookConfig().getSheetList().get(0));
      assertNotNull("MemPOI first mempoiSheet sheet name not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetName());
      assertNotEquals("MemPOI first mempoiSheet sheet name not empty", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetName(), "");
      assertNotNull("MemPOI first mempoiSheet prepStmt not null", memPOI.getWorkbookConfig().getSheetList().get(0).getPrepStmt());
      assertNotNull("MemPOI mempoiReportStyler not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler());
      assertNotNull("MemPOI mempoiReportStyler CommonDataCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getCommonDataCellStyle());
      assertNotNull("MemPOI mempoiReportStyler DateCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getDateCellStyle());
      assertNotNull("MemPOI mempoiReportStyler DatetimeCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getDatetimeCellStyle());
      assertNotNull("MemPOI mempoiReportStyler HeaderCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getHeaderCellStyle());
      assertNotNull("MemPOI mempoiReportStyler NumberCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getNumberCellStyle());
      assertNotNull("MemPOI mempoiReportStyler SubFooterCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getSubFooterCellStyle());
      assertTrue("MemPOI evaluateCellFormulas true", memPOI.getWorkbookConfig().isEvaluateCellFormulas());
      assertTrue("MemPOI evaluateCellFormulas true", memPOI.getWorkbookConfig().isHasFormulasToEvaluate());
   }




   @Test
   public void mempoiBuilderMinimumPopulated() {

      MemPOI memPOI = new MempoiBuilder()
              .addMempoiSheet(new MempoiSheet(prepStmt))
              .build();

      assertNotNull("MemPOIBuilder returns a not null MemPOI", memPOI);
      assertNotNull("MemPOI workbookconfig not null", memPOI.getWorkbookConfig());
      assertNull("MemPOI file null", memPOI.getWorkbookConfig().getFile());
      assertFalse("MemPOI adjustColumnWidth false", memPOI.getWorkbookConfig().isAdjustColSize());
      assertNotNull("MemPOI workbook not null", memPOI.getWorkbookConfig().getWorkbook());
      assertNotNull("MemPOI mempoiSheetList not null", memPOI.getWorkbookConfig().getSheetList());
      assertEquals("MemPOI mempoiSheetList size 1", 1, memPOI.getWorkbookConfig().getSheetList().size());
      assertNotNull("MemPOI first mempoiSheet not null", memPOI.getWorkbookConfig().getSheetList().get(0));
      assertNull("MemPOI first mempoiSheet sheet name null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetName());
      assertNotNull("MemPOI first mempoiSheet prepStmt not null", memPOI.getWorkbookConfig().getSheetList().get(0).getPrepStmt());
      this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler(), "mempoiReportStyler");
      assertFalse("MemPOI evaluateCellFormulas false", memPOI.getWorkbookConfig().isEvaluateCellFormulas());
      assertFalse("MemPOI evaluateCellFormulas true", memPOI.getWorkbookConfig().isHasFormulasToEvaluate());
   }



   @Test
   public void mempoiBuilderFullSheetlistPopulated() {

      SXSSFWorkbook workbook = new SXSSFWorkbook();

      CellStyle cellStyle = workbook.createCellStyle();
      cellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
      cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

      MempoiSheet sheet1 = MempoiSheetBuilder.aMempoiSheet()
              .withPrepStmt(prepStmt)
              .withSheetName("Sheet 1")
              .withWorkbook(workbook)
              .withStyleTemplate(new StoneStyleTemplate())
              .withHeaderCellStyle(cellStyle)
              .withSubFooterCellStyle(cellStyle)
              .withCommonDataCellStyle(cellStyle)
              .withDateCellStyle(cellStyle)
              .withDatetimeCellStyle(cellStyle)
              .withNumberCellStyle(cellStyle)
              .withMempoiFooter(new StandardMempoiFooter(workbook, "title"))
              .withMempoiSubFooter(new NumberSumSubFooter())
              .build();

      MempoiSheet sheet2 = new MempoiSheet(prepStmt);

      MempoiSheet sheet3 = new MempoiSheet(prepStmt, "Sheet 3");
      sheet3.setDateCellStyle(cellStyle);

      MempoiSheet sheet4 = new MempoiSheet(prepStmt);
      sheet4.setMempoiSubFooter(new NumberMinSubFooter());

      MempoiSheet sheet5 = new MempoiSheet(prepStmt);
      sheet5.setMempoiFooter(new StandardMempoiFooter(workbook, "title"));

      MemPOI memPOI = new MempoiBuilder()
              .addMempoiSheet(sheet1)
              .addMempoiSheet(sheet2)
              .addMempoiSheet(sheet3)
              .addMempoiSheet(sheet4)
              .addMempoiSheet(sheet5)
              .build();

      assertNotNull("MemPOIBuilder returns a not null MemPOI", memPOI);
      assertNotNull("MemPOI workbookconfig not null", memPOI.getWorkbookConfig());
      assertNotNull("sheetlist not null", memPOI.getWorkbookConfig().getSheetList());
      assertEquals("sheetlist size 5", 5, memPOI.getWorkbookConfig().getSheetList().size());

      int i = 0;
      String sheetName = "sheet 1";
      assertNotNull(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt());
      assertNotNull(sheetName + " not null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName());
      assertNotEquals(sheetName + " not empty title", "", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName());
      this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
      assertTrue(sheetName + " null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
      assertTrue(sheetName + " null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

      i = 1;
      sheetName = "sheet 2";
      assertNotNull(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt());
      assertNull(sheetName + " null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName());
      assertNotNull(sheetName + " not null styler", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler());
      this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
      assertFalse(sheetName + " null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
      assertFalse(sheetName + " null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

      i = 2;
      sheetName = "sheet 3";
      assertNotNull(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt());
      assertNotNull(sheetName + " not null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName());
      assertNotEquals(sheetName + " not empty title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), "");
      this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
      assertFalse(sheetName + " null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
      assertFalse(sheetName + " null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

      i = 3;
      sheetName = "sheet 4";
      assertNotNull(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt());
      assertNull(sheetName + " not null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName());
      assertNotEquals(sheetName + " not empty title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), "");
      this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
      assertTrue(sheetName + " not null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
      assertFalse(sheetName + " null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

      i = 4;
      sheetName = "sheet 5";
      assertNotNull(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt());
      assertNull(sheetName + " not null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName());
      assertNotEquals(sheetName + " not empty title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), "");
      this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
      assertFalse(sheetName + " null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
      assertTrue(sheetName + " not null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

   }


   /**
    * generics asserts for MempoiStyler
    * @param styler
    * @param stylerName
    */
   private void assertStyles(MempoiStyler styler, String stylerName) {

      assertNotNull(stylerName + " not null", styler);
      assertNotNull(stylerName + " CommonDataCellStyle not null", styler.getCommonDataCellStyle());
      assertNotNull(stylerName + " DateCellStyle not null", styler.getDateCellStyle());
      assertNotNull(stylerName + " DatetimeCellStyle not null", styler.getDatetimeCellStyle());
      assertNotNull(stylerName + " HeaderCellStyle not null", styler.getHeaderCellStyle());
      assertNotNull(stylerName + " NumberCellStyle not null", styler.getNumberCellStyle());
      assertNotNull(stylerName + " SubFooterCellStyle not null", styler.getSubFooterCellStyle());
   }
}
