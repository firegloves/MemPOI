package it.firegloves.mempoi.domain;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.exception.MempoiException;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.junit.Test;

public class MempoiEncryptionTest {

    @Test
    public void shouldReturnTheRightEntityWithFullBuilderPopulation() {

        String password = "mempaz";
        EncryptionInfo encryptionInfo = new EncryptionInfo(EncryptionMode.xor);

        MempoiEncryption mempoiEncryption = MempoiEncryption.MempoiEncryptionBuilder.aMempoiEncryption()
                .withPassword(password)
                .withEncryptionInfo(encryptionInfo)
                .build();

        assertEquals(password, mempoiEncryption.getPassword());
        assertEquals(encryptionInfo, mempoiEncryption.getEncryptionInfo());
    }

    @Test
    public void shouldCreateTheDefaultEncryptionInfoIfNotSupplied() {

        String password = "mempaz";

        MempoiEncryption mempoiEncryption = MempoiEncryption.MempoiEncryptionBuilder.aMempoiEncryption()
                .withPassword(password)
                .build();

        assertEquals(password, mempoiEncryption.getPassword());
        assertEquals(EncryptionMode.agile, mempoiEncryption.getEncryptionInfo().getEncryptionMode());
    }

    @Test
    public void shouldThrowExceptionIfNoPasswordHasBeenSet() {

        try {
            MempoiEncryption.MempoiEncryptionBuilder.aMempoiEncryption()
                    .build();
        } catch (MempoiException e) {
            assertEquals("MemPOI received a workbook encryption but no password has been set", e.getMessage());
        }
    }
}
