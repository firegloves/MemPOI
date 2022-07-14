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
              template.getSimpleTextHeaderCellStyle(wb),
              template.getColsHeaderCellStyle(wb),
              template.getCommonDataCellStyle(wb),
              template.getDateCellStyle(wb),
              template.getDatetimeCellStyle(wb),
              template.getIntegerCellStyle(wb),
              template.getFloatingPointCellStyle(wb),
              template.getSimpleTextFooterCellStyle(wb),
              template.getSubfooterCellStyle(wb)
      );

      AssertionHelper.assertOnCellStyle(styler.getSimpleTextHeaderCellStyle(), template.getSimpleTextHeaderCellStyle(wb));
      AssertionHelper.assertOnCellStyle(styler.getColsHeaderCellStyle(), template.getColsHeaderCellStyle(wb));
      AssertionHelper.assertOnCellStyle(styler.getCommonDataCellStyle(), template.getCommonDataCellStyle(wb));
      AssertionHelper.assertOnCellStyle(styler.getDateCellStyle(), template.getDateCellStyle(wb));
      AssertionHelper.assertOnCellStyle(styler.getDatetimeCellStyle(), template.getDatetimeCellStyle(wb));
      AssertionHelper.assertOnCellStyle(styler.getIntegerCellStyle(), template.getIntegerCellStyle(wb));
      AssertionHelper.assertOnCellStyle(styler.getFloatingPointCellStyle(), template.getFloatingPointCellStyle(wb));
      AssertionHelper.assertOnCellStyle(styler.getSimpleTextFooterCellStyle(), template.getSimpleTextFooterCellStyle(wb));
      AssertionHelper.assertOnCellStyle(styler.getSubFooterCellStyle(), template.getSubfooterCellStyle(wb));
   }
}
