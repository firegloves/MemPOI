package it.firegloves.mempoi.testutil;

import java.io.FileInputStream;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class EncryptionHelper {

    public static boolean isEncrypted(String path) throws Exception {
        try {
            new POIFSFileSystem(new FileInputStream(path));
            return true;
        } catch (OfficeXmlFileException e) {
            return false;
        }
    }
}
