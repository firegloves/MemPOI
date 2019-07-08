package it.firegloves.mempoi.unit;

import it.firegloves.mempoi.domain.EExportDataType;
import it.firegloves.mempoi.domain.MempoiColumn;
import org.apache.poi.ss.usermodel.Cell;
import org.junit.Test;

import java.sql.ResultSet;
import java.sql.Types;

import static org.hamcrest.CoreMatchers.nullValue;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

public class MempoiColumnTest {

   // TODO test MempoiColumn.getCellStyle and MempoiColumn.getSubFooterCell()

   @Test
   public void column_BIGINT() throws NoSuchMethodException {
      this.testMempoiColumn("column_BIGINT", Types.BIGINT, EExportDataType.DOUBLE);
   }

   @Test
   public void column_DOUBLE() throws NoSuchMethodException {
      this.testMempoiColumn("column_DOUBLE", Types.DOUBLE, EExportDataType.DOUBLE);
   }

   @Test
   public void column_DECIMAL() throws NoSuchMethodException {
      this.testMempoiColumn("column_DECIMAL", Types.DECIMAL, EExportDataType.FLOAT);
   }

   @Test
   public void column_FLOAT() throws NoSuchMethodException {
      this.testMempoiColumn("column_FLOAT", Types.FLOAT, EExportDataType.FLOAT);
   }

   @Test
   public void column_NUMERIC() throws NoSuchMethodException {
      this.testMempoiColumn("column_NUMERIC", Types.NUMERIC, EExportDataType.FLOAT);
   }

   @Test
   public void column_REAL() throws NoSuchMethodException {
      this.testMempoiColumn("column_REAL", Types.REAL, EExportDataType.FLOAT);
   }

   @Test
   public void column_INTEGER() throws NoSuchMethodException {
      this.testMempoiColumn("column_INTEGER", Types.INTEGER, EExportDataType.INT);
   }

   @Test
   public void column_SMALLINT() throws NoSuchMethodException {
      this.testMempoiColumn("column_SMALLINT", Types.SMALLINT, EExportDataType.INT);
   }

   @Test
   public void column_TINYINT() throws NoSuchMethodException {
      this.testMempoiColumn("column_TINYINT", Types.TINYINT, EExportDataType.INT);
   }

   @Test
   public void column_CHAR() throws NoSuchMethodException {
      this.testMempoiColumn("column_CHAR", Types.CHAR, EExportDataType.TEXT);
   }

   @Test
   public void column_NCHAR() throws NoSuchMethodException {
      this.testMempoiColumn("column_NCHAR", Types.NCHAR, EExportDataType.TEXT);
   }

   @Test
   public void column_VARCHAR() throws NoSuchMethodException {
      this.testMempoiColumn("column_VARCHAR", Types.VARCHAR, EExportDataType.TEXT);
   }

   @Test
   public void column_NVARCHAR() throws NoSuchMethodException {
      this.testMempoiColumn("column_NVARCHAR", Types.NVARCHAR, EExportDataType.TEXT);
   }

   @Test
   public void column_LONGVARCHAR() throws NoSuchMethodException {
      this.testMempoiColumn("column_LONGVARCHAR", Types.LONGVARCHAR, EExportDataType.TEXT);
   }

   @Test
   public void column_TIMESTAMP() throws NoSuchMethodException {
      this.testMempoiColumn("column_TIMESTAMP", Types.TIMESTAMP, EExportDataType.TIMESTAMP);
   }

   @Test
   public void column_DATE() throws NoSuchMethodException {
      this.testMempoiColumn("column_DATE", Types.DATE, EExportDataType.DATE);
   }

   @Test
   public void column_TIME() throws NoSuchMethodException {
      this.testMempoiColumn("column_TIME", Types.TIME, EExportDataType.TIME);
   }

   @Test
   public void column_BIT() throws NoSuchMethodException {
      this.testMempoiColumn("column_BIT", Types.BIT, EExportDataType.BOOLEAN);
   }

   @Test
   public void column_BOOLEAN() throws NoSuchMethodException {
      this.testMempoiColumn("column_BOOLEAN", Types.BOOLEAN, EExportDataType.BOOLEAN);
   }




   /**
    * parametrized method to test MempoiColumn
    * @param colName
    * @param sqlObjType
    * @param eExportDataType
    * @throws NoSuchMethodException
    */
   private void testMempoiColumn(String colName, int sqlObjType, EExportDataType eExportDataType) throws NoSuchMethodException {

      MempoiColumn mc = new MempoiColumn(sqlObjType, colName);

      assertEquals("mc " + colName + " EExportDataType", eExportDataType, mc.getType());
      assertEquals("mc " + colName + " column name", colName, mc.getColumnName());
      assertEquals("mc " + colName + " cellSetValueMethod", Cell.class.getMethod("setCellValue", eExportDataType.getRsReturnClass()), mc.getCellSetValueMethod());
      assertNull("mc " + colName + " cellStyle", mc.getCellStyle());
      assertEquals("mc " + colName + " rsAccessDataMethod", ResultSet.class.getMethod(eExportDataType.getRsAccessDataMethodName(), eExportDataType.getRsAccessParamClass()), mc.getRsAccessDataMethod());
   }

}
