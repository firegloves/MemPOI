package it.firegloves.mempoi.testutil;

import it.firegloves.mempoi.config.MempoiConfig;

public class ForceGenerationUtils {

    /**
     * executes a test enabling force execution for the single test
     * @param runnable the test logic to execute
     */
    public static void executeTestWithForceGeneration(Runnable runnable) {

        try {
            MempoiConfig.getInstance().setForceGeneration(true);
            runnable.run();
        } finally {
            MempoiConfig.getInstance().setForceGeneration(false);
        }
    }
}
