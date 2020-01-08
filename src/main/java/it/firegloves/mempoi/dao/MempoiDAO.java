package it.firegloves.mempoi.dao;

import it.firegloves.mempoi.domain.MempoiColumn;

import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.List;

public interface MempoiDAO {

    /**
     * executes the received PreparedStatement needed to generate poi export
     *
     * IT IS RETURNING A ResultSet - YOU HAVE TO CLOSE IT OUTSIDE
     *
     * @param prepStmt the PreparedStatement to select export data ResultSet
     *
     * @return the resulting ResultSet
     */
    ResultSet executeExportQuery(PreparedStatement prepStmt);


    /**
     * read metadatas from the received ResultSet and return the correspongin List of MempoiColumn
     *
     * @param rs the ResultSet from which to read metadata
     * @return a List of MempoiColumn corresponding to the export relevant columns metadata
     */
    List<MempoiColumn> readMetadata(ResultSet rs);
}
