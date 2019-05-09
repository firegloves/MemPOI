package it.firegloves.mempoi.unit;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
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
      assertThat("MemPOI file not null", memPOI.getFile(), notNullValue());
      assertEquals("MemPOI adjustColumnWidth true ", memPOI.isAdjustColumnWidth(), true);
      assertThat("MemPOI workbook not null", memPOI.getWorkbook(), notNullValue());
      assertThat("MemPOI mempoiSheetList not null", memPOI.getMempoiSheetList(), notNullValue());
      assertEquals("MemPOI mempoiSheetList size 1", memPOI.getMempoiSheetList().size(), 1);
      assertThat("MemPOI first mempoiSheet not null", memPOI.getMempoiSheetList().get(0), notNullValue());
      assertThat("MemPOI first mempoiSheet sheet name not null", memPOI.getMempoiSheetList().get(0).getSheetName(), notNullValue());
      assertNotEquals("MemPOI first mempoiSheet sheet name not empty", memPOI.getMempoiSheetList().get(0).getSheetName(), "");
      assertThat("MemPOI first mempoiSheet prepStmt not null", memPOI.getMempoiSheetList().get(0).getPrepStmt(), notNullValue());
      assertThat("MemPOI mempoiReportStyler not null", memPOI.getMempoiReportStyler(), notNullValue());
      assertThat("MemPOI mempoiReportStyler CommonDataCellStyle not null", memPOI.getMempoiReportStyler().getCommonDataCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler DateCellStyle not null", memPOI.getMempoiReportStyler().getDateCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler DatetimeCellStyle not null", memPOI.getMempoiReportStyler().getDatetimeCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler HeaderCellStyle not null", memPOI.getMempoiReportStyler().getHeaderCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler NumberCellStyle not null", memPOI.getMempoiReportStyler().getNumberCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler SubFooterCellStyle not null", memPOI.getMempoiReportStyler().getSubFooterCellStyle(), notNullValue());
      assertEquals("MemPOI evaluateCellFormulas true", memPOI.isEvaluateCellFormulas(), true);
   }




   @Test
   public void mempoi_builder_minimum_populated() {

      MemPOI memPOI = new MempoiBuilder()
              .addMempoiSheet(new MempoiSheet(prepStmt))
              .build();

      assertThat("MemPOIBuilder returns a not null MemPOI", memPOI, notNullValue());
      assertThat("MemPOI file null", memPOI.getFile(), nullValue());
      assertEquals("MemPOI adjustColumnWidth false", memPOI.isAdjustColumnWidth(), false);
      assertThat("MemPOI workbook not null", memPOI.getWorkbook(), notNullValue());
      assertThat("MemPOI mempoiSheetList not null", memPOI.getMempoiSheetList(), notNullValue());
      assertEquals("MemPOI mempoiSheetList size 1", memPOI.getMempoiSheetList().size(), 1);
      assertThat("MemPOI first mempoiSheet not null", memPOI.getMempoiSheetList().get(0), notNullValue());
      assertThat("MemPOI first mempoiSheet sheet name null", memPOI.getMempoiSheetList().get(0).getSheetName(), nullValue());
      assertThat("MemPOI first mempoiSheet prepStmt not null", memPOI.getMempoiSheetList().get(0).getPrepStmt(), notNullValue());
      assertThat("MemPOI mempoiReportStyler not null", memPOI.getMempoiReportStyler(), notNullValue());
      assertThat("MemPOI mempoiReportStyler CommonDataCellStyle not null", memPOI.getMempoiReportStyler().getCommonDataCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler DateCellStyle not null", memPOI.getMempoiReportStyler().getDateCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler DatetimeCellStyle not null", memPOI.getMempoiReportStyler().getDatetimeCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler HeaderCellStyle not null", memPOI.getMempoiReportStyler().getHeaderCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler NumberCellStyle not null", memPOI.getMempoiReportStyler().getNumberCellStyle(), notNullValue());
      assertThat("MemPOI mempoiReportStyler SubFooterCellStyle not null", memPOI.getMempoiReportStyler().getSubFooterCellStyle(), notNullValue());
      assertEquals("MemPOI evaluateCellFormulas false", memPOI.isEvaluateCellFormulas(), false);
   }

}
