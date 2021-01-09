/**
 * contain the configuration of the encryption to apply to the mempoi report
 */
package it.firegloves.mempoi.domain;

import lombok.Builder;
import lombok.Builder.Default;
import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.poifs.crypt.ChainingMode;
import org.apache.poi.poifs.crypt.CipherAlgorithm;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.HashAlgorithm;

@Data
@Accessors(chain = true)
@Builder(setterPrefix = "with")
public class MempoiEncryption {

    /**
     * password to open the encrypted resulting file
     */
    private String password;
    /**
     * the encryption mode to use
     */
    @Default
    private EncryptionMode encryptionMode = EncryptionMode.agile;
    /**
     * the cipher algorithm to use
     */
    private CipherAlgorithm cipherAlgorithm;
    /**
     * the hash algorithm to use
     */
    private HashAlgorithm hashAlgorithm;
    private int keyBits;
    private int blockSize;
    /**
     * the chaining mode to use
     */
    private ChainingMode chainingMode;
}
