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

import java.io.File;
import java.io.InputStream;
import java.io.Reader;
import java.math.BigDecimal;
import java.net.URL;
import java.sql.*;
import java.util.Calendar;

import static org.hamcrest.CoreMatchers.notNullValue;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.junit.Assert.*;

public class MempoiBuildertTest {

   @Mock
   private PreparedStatement prepStmt;

   @Before
   public void prepare() {
      this.prepStmt = this.getCustomPrepStmt();
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



   private PreparedStatement getCustomPrepStmt() {
      return new PreparedStatement() {
         @Override
         public ResultSet executeQuery() throws SQLException {
            return null;
         }

         @Override
         public int executeUpdate() throws SQLException {
            return 0;
         }

         @Override
         public void setNull(int parameterIndex, int sqlType) throws SQLException {

         }

         @Override
         public void setBoolean(int parameterIndex, boolean x) throws SQLException {

         }

         @Override
         public void setByte(int parameterIndex, byte x) throws SQLException {

         }

         @Override
         public void setShort(int parameterIndex, short x) throws SQLException {

         }

         @Override
         public void setInt(int parameterIndex, int x) throws SQLException {

         }

         @Override
         public void setLong(int parameterIndex, long x) throws SQLException {

         }

         @Override
         public void setFloat(int parameterIndex, float x) throws SQLException {

         }

         @Override
         public void setDouble(int parameterIndex, double x) throws SQLException {

         }

         @Override
         public void setBigDecimal(int parameterIndex, BigDecimal x) throws SQLException {

         }

         @Override
         public void setString(int parameterIndex, String x) throws SQLException {

         }

         @Override
         public void setBytes(int parameterIndex, byte[] x) throws SQLException {

         }

         @Override
         public void setDate(int parameterIndex, Date x) throws SQLException {

         }

         @Override
         public void setTime(int parameterIndex, Time x) throws SQLException {

         }

         @Override
         public void setTimestamp(int parameterIndex, Timestamp x) throws SQLException {

         }

         @Override
         public void setAsciiStream(int parameterIndex, InputStream x, int length) throws SQLException {

         }

         @Override
         public void setUnicodeStream(int parameterIndex, InputStream x, int length) throws SQLException {

         }

         @Override
         public void setBinaryStream(int parameterIndex, InputStream x, int length) throws SQLException {

         }

         @Override
         public void clearParameters() throws SQLException {

         }

         @Override
         public void setObject(int parameterIndex, Object x, int targetSqlType) throws SQLException {

         }

         @Override
         public void setObject(int parameterIndex, Object x) throws SQLException {

         }

         @Override
         public boolean execute() throws SQLException {
            return false;
         }

         @Override
         public void addBatch() throws SQLException {

         }

         @Override
         public void setCharacterStream(int parameterIndex, Reader reader, int length) throws SQLException {

         }

         @Override
         public void setRef(int parameterIndex, Ref x) throws SQLException {

         }

         @Override
         public void setBlob(int parameterIndex, Blob x) throws SQLException {

         }

         @Override
         public void setClob(int parameterIndex, Clob x) throws SQLException {

         }

         @Override
         public void setArray(int parameterIndex, Array x) throws SQLException {

         }

         @Override
         public ResultSetMetaData getMetaData() throws SQLException {
            return null;
         }

         @Override
         public void setDate(int parameterIndex, Date x, Calendar cal) throws SQLException {

         }

         @Override
         public void setTime(int parameterIndex, Time x, Calendar cal) throws SQLException {

         }

         @Override
         public void setTimestamp(int parameterIndex, Timestamp x, Calendar cal) throws SQLException {

         }

         @Override
         public void setNull(int parameterIndex, int sqlType, String typeName) throws SQLException {

         }

         @Override
         public void setURL(int parameterIndex, URL x) throws SQLException {

         }

         @Override
         public ParameterMetaData getParameterMetaData() throws SQLException {
            return null;
         }

         @Override
         public void setRowId(int parameterIndex, RowId x) throws SQLException {

         }

         @Override
         public void setNString(int parameterIndex, String value) throws SQLException {

         }

         @Override
         public void setNCharacterStream(int parameterIndex, Reader value, long length) throws SQLException {

         }

         @Override
         public void setNClob(int parameterIndex, NClob value) throws SQLException {

         }

         @Override
         public void setClob(int parameterIndex, Reader reader, long length) throws SQLException {

         }

         @Override
         public void setBlob(int parameterIndex, InputStream inputStream, long length) throws SQLException {

         }

         @Override
         public void setNClob(int parameterIndex, Reader reader, long length) throws SQLException {

         }

         @Override
         public void setSQLXML(int parameterIndex, SQLXML xmlObject) throws SQLException {

         }

         @Override
         public void setObject(int parameterIndex, Object x, int targetSqlType, int scaleOrLength) throws SQLException {

         }

         @Override
         public void setAsciiStream(int parameterIndex, InputStream x, long length) throws SQLException {

         }

         @Override
         public void setBinaryStream(int parameterIndex, InputStream x, long length) throws SQLException {

         }

         @Override
         public void setCharacterStream(int parameterIndex, Reader reader, long length) throws SQLException {

         }

         @Override
         public void setAsciiStream(int parameterIndex, InputStream x) throws SQLException {

         }

         @Override
         public void setBinaryStream(int parameterIndex, InputStream x) throws SQLException {

         }

         @Override
         public void setCharacterStream(int parameterIndex, Reader reader) throws SQLException {

         }

         @Override
         public void setNCharacterStream(int parameterIndex, Reader value) throws SQLException {

         }

         @Override
         public void setClob(int parameterIndex, Reader reader) throws SQLException {

         }

         @Override
         public void setBlob(int parameterIndex, InputStream inputStream) throws SQLException {

         }

         @Override
         public void setNClob(int parameterIndex, Reader reader) throws SQLException {

         }

         @Override
         public ResultSet executeQuery(String sql) throws SQLException {
            return null;
         }

         @Override
         public int executeUpdate(String sql) throws SQLException {
            return 0;
         }

         @Override
         public void close() throws SQLException {

         }

         @Override
         public int getMaxFieldSize() throws SQLException {
            return 0;
         }

         @Override
         public void setMaxFieldSize(int max) throws SQLException {

         }

         @Override
         public int getMaxRows() throws SQLException {
            return 0;
         }

         @Override
         public void setMaxRows(int max) throws SQLException {

         }

         @Override
         public void setEscapeProcessing(boolean enable) throws SQLException {

         }

         @Override
         public int getQueryTimeout() throws SQLException {
            return 0;
         }

         @Override
         public void setQueryTimeout(int seconds) throws SQLException {

         }

         @Override
         public void cancel() throws SQLException {

         }

         @Override
         public SQLWarning getWarnings() throws SQLException {
            return null;
         }

         @Override
         public void clearWarnings() throws SQLException {

         }

         @Override
         public void setCursorName(String name) throws SQLException {

         }

         @Override
         public boolean execute(String sql) throws SQLException {
            return false;
         }

         @Override
         public ResultSet getResultSet() throws SQLException {
            return null;
         }

         @Override
         public int getUpdateCount() throws SQLException {
            return 0;
         }

         @Override
         public boolean getMoreResults() throws SQLException {
            return false;
         }

         @Override
         public void setFetchDirection(int direction) throws SQLException {

         }

         @Override
         public int getFetchDirection() throws SQLException {
            return 0;
         }

         @Override
         public void setFetchSize(int rows) throws SQLException {

         }

         @Override
         public int getFetchSize() throws SQLException {
            return 0;
         }

         @Override
         public int getResultSetConcurrency() throws SQLException {
            return 0;
         }

         @Override
         public int getResultSetType() throws SQLException {
            return 0;
         }

         @Override
         public void addBatch(String sql) throws SQLException {

         }

         @Override
         public void clearBatch() throws SQLException {

         }

         @Override
         public int[] executeBatch() throws SQLException {
            return new int[0];
         }

         @Override
         public Connection getConnection() throws SQLException {
            return null;
         }

         @Override
         public boolean getMoreResults(int current) throws SQLException {
            return false;
         }

         @Override
         public ResultSet getGeneratedKeys() throws SQLException {
            return null;
         }

         @Override
         public int executeUpdate(String sql, int autoGeneratedKeys) throws SQLException {
            return 0;
         }

         @Override
         public int executeUpdate(String sql, int[] columnIndexes) throws SQLException {
            return 0;
         }

         @Override
         public int executeUpdate(String sql, String[] columnNames) throws SQLException {
            return 0;
         }

         @Override
         public boolean execute(String sql, int autoGeneratedKeys) throws SQLException {
            return false;
         }

         @Override
         public boolean execute(String sql, int[] columnIndexes) throws SQLException {
            return false;
         }

         @Override
         public boolean execute(String sql, String[] columnNames) throws SQLException {
            return false;
         }

         @Override
         public int getResultSetHoldability() throws SQLException {
            return 0;
         }

         @Override
         public boolean isClosed() throws SQLException {
            return false;
         }

         @Override
         public void setPoolable(boolean poolable) throws SQLException {

         }

         @Override
         public boolean isPoolable() throws SQLException {
            return false;
         }

         @Override
         public void closeOnCompletion() throws SQLException {

         }

         @Override
         public boolean isCloseOnCompletion() throws SQLException {
            return false;
         }

         @Override
         public <T> T unwrap(Class<T> iface) throws SQLException {
            return null;
         }

         @Override
         public boolean isWrapperFor(Class<?> iface) throws SQLException {
            return false;
         }
      };
   }
}
