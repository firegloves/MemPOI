/**
 * created by firegloves
 */

package it.firegloves.mempoi.config;

import static junit.framework.TestCase.assertTrue;

import org.junit.Test;

public class MempoiConfigTest {


    @Test
    public void forceGenerationTestTrue() {
        MempoiConfig.getInstance().setForceGeneration(true);
        assertTrue(MempoiConfig.getInstance().isForceGeneration());
        MempoiConfig.getInstance().setForceGeneration(false);
    }
}
