package it.firegloves.mempoi.unit;

import it.firegloves.mempoi.styles.template.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import static org.hamcrest.CoreMatchers.notNullValue;
import static org.junit.Assert.assertThat;

public class StyleTemplateTest {

   @Mock
   Workbook workbook;

   @Before
   public void prepare() {
      MockitoAnnotations.initMocks(this);
   }

   @Test
   public void standard_template_test() {
      this.genericTemplateTest(new StandardStyleTemplate(), new SXSSFWorkbook());
   }

   @Test
   public void aqua_template_test() {
      this.genericTemplateTest(new AquaStyleTemplate(), new SXSSFWorkbook());
   }

   @Test
   public void panegiricon_template_test() {
      this.genericTemplateTest(new PanegiriconStyleTemplate(), new SXSSFWorkbook());
   }

   @Test
   public void forest_template_test() {
      this.genericTemplateTest(new ForestStyleTemplate(), new SXSSFWorkbook());
   }

   @Test
   public void purple_template_test() {
      this.genericTemplateTest(new PurpleStyleTemplate(), new SXSSFWorkbook());
   }

   @Test
   public void rose_template_test() {
      this.genericTemplateTest(new RoseStyleTemplate(), new SXSSFWorkbook());
   }

   @Test
   public void stone_template_test() {
      this.genericTemplateTest(new StoneStyleTemplate(), new SXSSFWorkbook());
   }

   @Test
   public void summer_template_test() {
      this.genericTemplateTest(new SummerStyleTemplate(), new SXSSFWorkbook());
   }


   /**
    * runs generic style template tests
    * @param template
    * @param workbook
    */
   private void genericTemplateTest(StyleTemplate template, Workbook workbook) {

      assertThat("template " + template.getClass().getName() + " common data cell style not null", template.getCommonDataCellStyle(workbook), notNullValue());
      assertThat("template " + template.getClass().getName() + " date cell style not null", template.getDateCellStyle(workbook), notNullValue());
      assertThat("template " + template.getClass().getName() + " datetime cell style not null", template.getDatetimeCellStyle(workbook), notNullValue());
      assertThat("template " + template.getClass().getName() + " header cell style not null", template.getHeaderCellStyle(workbook), notNullValue());
      assertThat("template " + template.getClass().getName() + " number cell style not null", template.getNumberCellStyle(workbook), notNullValue());
      assertThat("template " + template.getClass().getName() + " footer cell style not null", template.getSubfooterCellStyle(workbook), notNullValue());
   }
}