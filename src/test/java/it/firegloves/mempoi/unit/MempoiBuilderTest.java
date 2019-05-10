package it.firegloves.mempoi.unit;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
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

import static org.hamcrest.CoreMatchers.notNullValue;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.junit.Assert.*;

public class MempoiBuilderTest {

   @Mock
   private PreparedStatement prepStmt;

   @Before
   public void prepare() {
      MockitoAnnotations.initMocks(this);
   }

   @Test
   public void mempoi_builder_full_populated() {

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

      assertThat("MemPOIBuilder returns a not null MemPOI", memPOI, notNullValue());
      assertThat("MemPOI workbookconfig not null", memPOI.getWorkbookConfig(), notNullValue());
      assertThat("MemPOI file not null", memPOI.getWorkbookConfig().getFile(), notNullValue());
      assertEquals("MemPOI adjustColumnWidth true ", memPOI.getWorkbookConfig().isAdjustColSize(), true);
      assertThat("MemPOI workbook not null", memPOI.getWorkbookConfig().getWorkbook(), notNullValue());
      assertThat("MemPOI mempoiSheetList not null", memPOI.getWorkbookConfig().getSheetList(), notNullValue());
      assertEquals("MemPOI mempoiSheetList size 1", memPOI.getWorkbookConfig().getSheetList().size(), 1);
      assertThat("MemPOI first mempoiSheet not null", memPOI.getWorkbookConfig().getSheetList().get(0), notNullValue());
      assertThat("MemPOI first mempoiSheet sheet name not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetName(), notNullValue());
      assertNotEquals("MemPOI first mempoiSheet sheet name not empty", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetName(), "");
      assertThat("MemPOI first mempoiSheet prepStmt not null", memPOI.getWorkbookConfig().getSheetList().get(0).getPrepStmt(), notNullValue());
      assertThat("MemPOI mempoiReportStyler not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler(), notNullValue());
      assertThat("MemPOI mempoiReportStyler CommonDataCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getCommonDataCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler DateCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getDateCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler DatetimeCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getDatetimeCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler HeaderCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getHeaderCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler NumberCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getNumberCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler SubFooterCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getSubFooterCellStyle(), notNullValue());
      assertEquals("MemPOI evaluateCellFormulas true", memPOI.getWorkbookConfig().isEvaluateCellFormulas(), true);
      assertEquals("MemPOI evaluateCellFormulas true", memPOI.getWorkbookConfig().isHasFormulasToEvaluate(), true);
   }




   @Test
   public void mempoi_builder_minimum_populated() {

      MemPOI memPOI = new MempoiBuilder()
              .addMempoiSheet(new MempoiSheet(prepStmt))
              .build();

      assertThat("MemPOIBuilder returns a not null MemPOI", memPOI, notNullValue());
      assertThat("MemPOI workbookconfig not null", memPOI.getWorkbookConfig(), notNullValue());
      assertThat("MemPOI file null", memPOI.getWorkbookConfig().getFile(), nullValue());
      assertEquals("MemPOI adjustColumnWidth false", memPOI.getWorkbookConfig().isAdjustColSize(), false);
      assertThat("MemPOI workbook not null", memPOI.getWorkbookConfig().getWorkbook(), notNullValue());
      assertThat("MemPOI mempoiSheetList not null", memPOI.getWorkbookConfig().getSheetList(), notNullValue());
      assertEquals("MemPOI mempoiSheetList size 1", memPOI.getWorkbookConfig().getSheetList().size(), 1);
      assertThat("MemPOI first mempoiSheet not null", memPOI.getWorkbookConfig().getSheetList().get(0), notNullValue());
      assertThat("MemPOI first mempoiSheet sheet name null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetName(), nullValue());
      assertThat("MemPOI first mempoiSheet prepStmt not null", memPOI.getWorkbookConfig().getSheetList().get(0).getPrepStmt(), notNullValue());
      this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler(), "mempoiReportStyler");
      assertEquals("MemPOI evaluateCellFormulas false", memPOI.getWorkbookConfig().isEvaluateCellFormulas(), false);
      assertEquals("MemPOI evaluateCellFormulas true", memPOI.getWorkbookConfig().isHasFormulasToEvaluate(), false);
   }



   @Test
   public void mempoi_builder_full_sheetlist_populated() {

      SXSSFWorkbook workbook = new SXSSFWorkbook();

      CellStyle cellStyle = workbook.createCellStyle();
      cellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
      cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

      MempoiSheet sheet1 = new MempoiSheet(prepStmt, "Sheet 1", workbook, new StoneStyleTemplate(),
              cellStyle, cellStyle, cellStyle, cellStyle, cellStyle, cellStyle,
              new StandardMempoiFooter(workbook, "title"), new NumberSumSubFooter());

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

      assertThat("MemPOIBuilder returns a not null MemPOI", memPOI, notNullValue());
      assertThat("MemPOI workbookconfig not null", memPOI.getWorkbookConfig(), notNullValue());
      assertThat("sheetlist not null", memPOI.getWorkbookConfig().getSheetList(), notNullValue());
      assertEquals("sheetlist size 5", memPOI.getWorkbookConfig().getSheetList().size(), 5);

      int i = 0;
      String sheetName = "sheet 1";
      assertThat(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt(), notNullValue());
      assertThat(sheetName + " not null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), notNullValue());
      assertNotEquals(sheetName + " not empty title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), "");
      this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
      assertTrue(sheetName + " null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
      assertTrue(sheetName + " null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

      i = 1;
      sheetName = "sheet 2";
      assertThat(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt(), notNullValue());
      assertThat(sheetName + " null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), nullValue());
      assertThat(sheetName + " not null styler", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), notNullValue());
      this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
      assertFalse(sheetName + " null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
      assertFalse(sheetName + " null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

      i = 2;
      sheetName = "sheet 3";
      assertThat(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt(), notNullValue());
      assertThat(sheetName + " not null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), notNullValue());
      assertNotEquals(sheetName + " not empty title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), "");
      this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
      assertFalse(sheetName + " null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
      assertFalse(sheetName + " null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

      i = 3;
      sheetName = "sheet 4";
      assertThat(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt(), notNullValue());
      assertThat(sheetName + " not null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), nullValue());
      assertNotEquals(sheetName + " not empty title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), "");
      this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
      assertTrue(sheetName + " not null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
      assertFalse(sheetName + " null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

      i = 4;
      sheetName = "sheet 5";
      assertThat(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt(), notNullValue());
      assertThat(sheetName + " not null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), nullValue());
      assertNotEquals(sheetName + " not empty title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), "");
      this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
      assertFalse(sheetName + " null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
      assertTrue(sheetName + " not null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

   }



   private void assertStyles(MempoiStyler styler, String stylerName) {

      assertThat(stylerName + " not null", styler, notNullValue());
      assertThat(stylerName + " CommonDataCellStyle not null", styler.getCommonDataCellStyle(), notNullValue());
      assertThat(stylerName + " DateCellStyle not null", styler.getDateCellStyle(), notNullValue());
      assertThat(stylerName + " DatetimeCellStyle not null", styler.getDatetimeCellStyle(), notNullValue());
      assertThat(stylerName + " HeaderCellStyle not null", styler.getHeaderCellStyle(), notNullValue());
      assertThat(stylerName + " NumberCellStyle not null", styler.getNumberCellStyle(), notNullValue());
      assertThat(stylerName + " SubFooterCellStyle not null", styler.getSubFooterCellStyle(), notNullValue());
   }
}