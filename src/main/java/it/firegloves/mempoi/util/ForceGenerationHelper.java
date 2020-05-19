/**
 * created by firegloves
 */

package it.firegloves.mempoi.util;


import it.firegloves.mempoi.config.MempoiConfig;
import it.firegloves.mempoi.exception.MempoiException;
import lombok.experimental.UtilityClass;
import org.slf4j.Logger;

@UtilityClass
public class ForceGenerationHelper {

    /**
     * if force generation is active logs the received error, otherwise throws the received exception
     *
     * @param e       the exception to throw if force generation is not enabled
     * @param message the message to log
     * @param logger  the logger used to log
     */
    public static void manageForceGeneration(MempoiException e, String message, Logger logger) {
        manageForceGeneration(e, message, logger, null);
    }


    /**
     * if force generation is active logs the received error, otherwise throws the received exception
     *
     * @param e       the exception to throw if force generation is not enabled
     * @param message the message to log
     * @param logger  the logger used to log
     * @param handler a function ForceGenerationHandler to apply some logic to the force generation event
     */
    public static void manageForceGeneration(MempoiException e, String message, Logger logger, ForceGenerationHandler handler) {

        if (MempoiConfig.getInstance().isForceGeneration()) {

            logger.error(message);

            if (null != handler) {
                handler.handle();
            }
        } else if (e != null) {
            throw e;
        }
    }


    /**
     * if force generation is active logs the received error, otherwise throws the received exception
     *
     * @param message the message to log
     * @param logger  the logger used to log
     * @param handler a function ForceGenerationHandler to apply some logic to the force generation event
     */
    public static void manageForceGeneration(String message, Logger logger, ForceGenerationHandler handler) {
        manageForceGeneration(null, message, logger, handler);
    }
}
