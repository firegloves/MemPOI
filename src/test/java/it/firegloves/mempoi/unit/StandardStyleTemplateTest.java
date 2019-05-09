package it.firegloves.mempoi.unit;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
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

public class StandardStyleTemplateTest {

   @Mock
   Workbook workbook;

   @Test
   public void mempoi_builder_full_populated() {

      StandardStyleTemplate template = new StandardStyleTemplate();

//      assertThat("template date cell style not null", template.getDateCellStyle(), notNullValue());
//      assertThat("template date cell style not null", template.getDateCellStyle(), notNullValue());
//      assertThat("template date cell style not null", template.getDateCellStyle(), notNullValue());
//      assertThat("template date cell style not null", template.getDateCellStyle(), notNullValue());
//      assertThat("template date cell style not null", template.getDateCellStyle(), notNullValue());
//      assertThat("template date cell style not null", template.getDateCellStyle(), notNullValue());
   }

}