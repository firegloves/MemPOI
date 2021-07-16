package it.firegloves.mempoi.domain;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.exception.MempoiException;
import java.util.HashMap;
import java.util.Map;
import lombok.Data;
import lombok.experimental.Accessors;

@Data
@Accessors(chain = true)
public class MempoiReport {

    /**
     * contains the filename (absolute path) where the MempoiReport has been saved (if exporting to file)
     */
    private String file;
    /**
     * contains the byte array containing MempoiReport data (if exporting to byte array)
     */
    private byte[] bytes;
    /**
     * the map of sheets' metadata the key is the sheet index the value is the metadata object
     */
    private Map<Integer, MempoiSheetMetadata> mempoiSheetMetadataMap = new HashMap<>();
    /**
     * the WorkbookConfig used to generate the report
     */
    private WorkbookConfig usedWorkbookConfig;

    /**
     * set the file in which the report has been saved
     *
     * @param file the file with absolute path in which the report has been saved
     * @return the current instance
     */
    public MempoiReport setFile(String file) {
        if (null != this.bytes) {
            throw new MempoiException("The Mempoi report in generation has both a file and a byte array");
        }
        this.file = file;
        return this;
    }

    /**
     * set the byte array representing the generated report
     *
     * @param bytes the byte array representing the generated report
     * @return the current instance
     */
    public MempoiReport setBytes(byte[] bytes) {
        if (null != this.file) {
            throw new MempoiException("The Mempoi report in generation has both a file and a byte array");
        }
        this.bytes = bytes;
        return this;
    }

    /**
     * insert a MempoiSheetMetadata relative to the sheet at the index identified by the first param
     *
     * @param sheetIndex          the index of the sheet of which save the metadata
     * @param mempoiSheetMetadata the object containing the sheet metadata
     * @return the current instance
     */
    public MempoiReport addMempoiSheetMetadata(Integer sheetIndex, MempoiSheetMetadata mempoiSheetMetadata) {
        if (null == this.mempoiSheetMetadataMap) {
            this.mempoiSheetMetadataMap = new HashMap<>();
        }

        if (sheetIndex < 0) {
            throw new MempoiException(String.format("The index of a sheet can't be minor than 0. %d received", sheetIndex));
        }

        if (mempoiSheetMetadata == null) {
            throw new MempoiException("The received MempoiSheetMetadata is null");
        }

        this.mempoiSheetMetadataMap.put(sheetIndex, mempoiSheetMetadata);
        return this;
    }

    /**
     * return the MempoiSheetMetadata corresponding to the sheet at index i
     * @param i the index of the sheet of which return the MempoiSheetMetadata
     * @return the MempoiSheetMetadata corresponding to the sheet at index i
     */
    public MempoiSheetMetadata getMempoiSheetMetadata(int i) {
        return this.mempoiSheetMetadataMap.get(i);
    }
}
