package it.firegloves.mempoi.styles;

import it.firegloves.mempoi.styles.template.*;
import it.firegloves.mempoi.testutil.AssertionHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

public class MempoiStylerTest {

   @Test
   public void constructorTest() {

      Workbook wb = new SXSSFWorkbook();
      ForestStyleTemplate template = new ForestStyleTemplate();

      MempoiStyler styler = new MempoiStyler(
              template.getHeaderCellStyle(wb),
              template.getCommonDataCellStyle(wb),
              template.getDateCellStyle(wb),
              template.getDatetimeCellStyle(wb),
              template.getNumberCellStyle(wb),
              template.getSubfooterCellStyle(wb)
      );

      AssertionHelper.validateCellStyle(styler.getHeaderCellStyle(), styler.getHeaderCellStyle());
      AssertionHelper.validateCellStyle(styler.getCommonDataCellStyle(), styler.getCommonDataCellStyle());
      AssertionHelper.validateCellStyle(styler.getDateCellStyle(), styler.getDateCellStyle());
      AssertionHelper.validateCellStyle(styler.getDatetimeCellStyle(), styler.getDatetimeCellStyle());
      AssertionHelper.validateCellStyle(styler.getNumberCellStyle(), styler.getNumberCellStyle());
      AssertionHelper.validateCellStyle(styler.getSubFooterCellStyle(), styler.getSubFooterCellStyle());
   }
}