/**
 * contain the configuration of the encryption to apply to the mempoi report
 */
package it.firegloves.mempoi.domain;

import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import lombok.AccessLevel;
import lombok.Data;
import lombok.Setter;
import lombok.experimental.Accessors;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;

@Data
@Setter(AccessLevel.PRIVATE)
@Accessors(chain = true)
//@Builder(setterPrefix = "with")
public class MempoiEncryption {

    /**
     * password to open the encrypted resulting file
     */
    private String password;
    /**
     * the encryption info to use to encrypt the file
     */
    private EncryptionInfo encryptionInfo;

    private MempoiEncryption() {}

    /****************************************************
     * BUILDER CLASS
     ***************************************************/
    public static class MempoiEncryptionBuilder {

        private String password;
        private EncryptionInfo encryptionInfo;

        private MempoiEncryptionBuilder() {
        }

        public static MempoiEncryptionBuilder aMempoiEncryption() {
            return new MempoiEncryptionBuilder();
        }

        public MempoiEncryptionBuilder withPassword(String password) {
            this.password = password;
            return this;
        }

        public MempoiEncryptionBuilder withEncryptionInfo(EncryptionInfo encryptionInfo) {
            this.encryptionInfo = encryptionInfo;
            return this;
        }

        public MempoiEncryption build() {

            if (null == password) {
                throw new MempoiException(Errors.ERR_ENCRYPTION_WITH_NO_PWD);
            }

            if (null == encryptionInfo) {
                encryptionInfo = new EncryptionInfo(EncryptionMode.agile);
            }

            MempoiEncryption mempoiEncryption = new MempoiEncryption();
            mempoiEncryption.setPassword(password);
            mempoiEncryption.setEncryptionInfo(encryptionInfo);

            return mempoiEncryption;
        }
    }
}
