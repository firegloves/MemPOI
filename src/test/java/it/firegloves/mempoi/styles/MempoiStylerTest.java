package it.firegloves.mempoi.styles;

import it.firegloves.mempoi.styles.template.*;
import it.firegloves.mempoi.testutil.AssertHelper;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

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

      AssertHelper.validateCellStyle(styler.getHeaderCellStyle(), styler.getHeaderCellStyle());
      AssertHelper.validateCellStyle(styler.getCommonDataCellStyle(), styler.getCommonDataCellStyle());
      AssertHelper.validateCellStyle(styler.getDateCellStyle(), styler.getDateCellStyle());
      AssertHelper.validateCellStyle(styler.getDatetimeCellStyle(), styler.getDatetimeCellStyle());
      AssertHelper.validateCellStyle(styler.getNumberCellStyle(), styler.getNumberCellStyle());
      AssertHelper.validateCellStyle(styler.getSubFooterCellStyle(), styler.getSubFooterCellStyle());
   }
}