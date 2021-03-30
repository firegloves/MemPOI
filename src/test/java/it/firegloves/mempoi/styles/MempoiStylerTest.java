package it.firegloves.mempoi.styles;

import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

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
              template.getIntegerCellStyle(wb),
              template.getFloatingPointCellStyle(wb),
              template.getSubfooterCellStyle(wb)
      );

      AssertionHelper.assertOnCellStyle(styler.getHeaderCellStyle(), styler.getHeaderCellStyle());
      AssertionHelper.assertOnCellStyle(styler.getCommonDataCellStyle(), styler.getCommonDataCellStyle());
      AssertionHelper.assertOnCellStyle(styler.getDateCellStyle(), styler.getDateCellStyle());
      AssertionHelper.assertOnCellStyle(styler.getDatetimeCellStyle(), styler.getDatetimeCellStyle());
      AssertionHelper.assertOnCellStyle(styler.getIntegerCellStyle(), styler.getIntegerCellStyle());
      AssertionHelper.assertOnCellStyle(styler.getFloatingPointCellStyle(), styler.getFloatingPointCellStyle());
      AssertionHelper.assertOnCellStyle(styler.getSubFooterCellStyle(), styler.getSubFooterCellStyle());
   }
}
