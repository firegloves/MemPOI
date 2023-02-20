package it.firegloves.mempoi.strategos;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiColumnConfig.MempoiColumnConfigBuilder;
import java.util.Arrays;
import java.util.List;
import org.junit.Test;

public class MempoiColumnStrategosTest {

    private MempoiColumnStrategos mempoiColumnStrategos = new MempoiColumnStrategos();

    @Test
    public void shouldIgnoreTheDesiredColumns() {

        final MempoiColumnConfig nameColConf = MempoiColumnConfigBuilder.aMempoiColumnConfig().withColumnName("name")
                .withIgnoreColumn(true).build();
        final MempoiColumnConfig ageColConf = MempoiColumnConfigBuilder.aMempoiColumnConfig().withColumnName("age")
                .withIgnoreColumn(true).build();
        final MempoiColumnConfig marriedColConf = MempoiColumnConfigBuilder.aMempoiColumnConfig().withColumnName("married")
                .withIgnoreColumn(true).build();

        final List<MempoiColumn> mempoiColumns = Arrays.asList(
                new MempoiColumn("id"),
                new MempoiColumn("name").setMempoiColumnConfig(nameColConf),
                new MempoiColumn("surname"),
                new MempoiColumn("age").setMempoiColumnConfig(ageColConf),
                new MempoiColumn("married").setMempoiColumnConfig(marriedColConf),
                new MempoiColumn("date"));

        final List<MempoiColumn> filteredCols = mempoiColumnStrategos.setIgnoreAndPositionOrder(mempoiColumns);

        assertEquals(3, filteredCols.size());
        assertEquals("id", filteredCols.get(0).getColumnName());
        assertEquals("surname", filteredCols.get(1).getColumnName());
        assertEquals("date", filteredCols.get(2).getColumnName());
    }

    @Test
    public void shouldRearrangeTheDesiredColumns() {

        final MempoiColumnConfig nameColConf = MempoiColumnConfigBuilder.aMempoiColumnConfig().withColumnName("name")
                .withPositionOrder(0).build();
        final MempoiColumnConfig marriedColConf = MempoiColumnConfigBuilder.aMempoiColumnConfig().withColumnName("married")
                .withPositionOrder(1).build();
        final MempoiColumnConfig ageColConf = MempoiColumnConfigBuilder.aMempoiColumnConfig().withColumnName("age")
                .withPositionOrder(2).build();

        final List<MempoiColumn> mempoiColumns = Arrays.asList(
                new MempoiColumn("id"),
                new MempoiColumn("name").setMempoiColumnConfig(nameColConf),
                new MempoiColumn("surname"),
                new MempoiColumn("age").setMempoiColumnConfig(ageColConf),
                new MempoiColumn("married").setMempoiColumnConfig(marriedColConf),
                new MempoiColumn("date"));

        final List<MempoiColumn> filteredCols = mempoiColumnStrategos.setIgnoreAndPositionOrder(mempoiColumns);

        assertEquals(6, filteredCols.size());
        assertEquals("name", filteredCols.get(0).getColumnName());
        assertEquals("married", filteredCols.get(1).getColumnName());
        assertEquals("age", filteredCols.get(2).getColumnName());
        assertEquals("id", filteredCols.get(3).getColumnName());
        assertEquals("surname", filteredCols.get(4).getColumnName());
        assertEquals("date", filteredCols.get(5).getColumnName());
    }

    @Test
    public void shouldIgnoreAndRearrangeTheDesiredColumns() {

        final MempoiColumnConfig nameColConf = MempoiColumnConfigBuilder.aMempoiColumnConfig().withColumnName("name")
                .withPositionOrder(0).build();
        final MempoiColumnConfig marriedColConf = MempoiColumnConfigBuilder.aMempoiColumnConfig().withColumnName("married")
                .withPositionOrder(1).build();
        final MempoiColumnConfig ageColConf = MempoiColumnConfigBuilder.aMempoiColumnConfig().withColumnName("age")
                .withPositionOrder(2).build();
        final MempoiColumnConfig surnameColConf = MempoiColumnConfigBuilder.aMempoiColumnConfig().withColumnName("surname")
                .withIgnoreColumn(true).build();

        final List<MempoiColumn> mempoiColumns = Arrays.asList(
                new MempoiColumn("id"),
                new MempoiColumn("name").setMempoiColumnConfig(nameColConf),
                new MempoiColumn("surname").setMempoiColumnConfig(surnameColConf),
                new MempoiColumn("age").setMempoiColumnConfig(ageColConf),
                new MempoiColumn("married").setMempoiColumnConfig(marriedColConf),
                new MempoiColumn("date"));

        final List<MempoiColumn> filteredCols = mempoiColumnStrategos.setIgnoreAndPositionOrder(mempoiColumns);

        assertEquals(5, filteredCols.size());
        assertEquals("name", filteredCols.get(0).getColumnName());
        assertEquals("married", filteredCols.get(1).getColumnName());
        assertEquals("age", filteredCols.get(2).getColumnName());
        assertEquals("id", filteredCols.get(3).getColumnName());
        assertEquals("date", filteredCols.get(4).getColumnName());
    }
}
