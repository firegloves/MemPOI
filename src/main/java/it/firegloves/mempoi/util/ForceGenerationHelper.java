/**
 * created by firegloves
 */

package it.firegloves.mempoi.util;


import it.firegloves.mempoi.config.MempoiConfig;
import it.firegloves.mempoi.exception.MempoiException;
import org.slf4j.Logger;

public class ForceGenerationHelper {

    /**
     * if force generation is active => logs the received error, otherwise throws the received exception
     * @param e
     * @param error
     * @param logger
     */
    public static void manageForceGeneration(MempoiException e, String error, Logger logger) {

        if (MempoiConfig.getInstance().isForceGeneration()) {
            logger.error(error);
        } else {
            throw e;
        }
    }
}
