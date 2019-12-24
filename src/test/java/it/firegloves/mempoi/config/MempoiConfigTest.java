/**
 * created by firegloves
 */

package it.firegloves.mempoi.config;

import org.junit.Test;

import static junit.framework.TestCase.assertFalse;
import static junit.framework.TestCase.assertTrue;

public class MempoiConfigTest {

    @Test
    public void debugTestTrue() {
        MempoiConfig.getInstance().setDebug(true);
        assertTrue(MempoiConfig.getInstance().isDebug());
    }

    @Test
    public void debugTestFlase() {
        assertFalse(MempoiConfig.getInstance().isDebug());
    }

    @Test
    public void forceGenerationTestTrue() {
        MempoiConfig.getInstance().setForceGeneration(true);
        assertTrue(MempoiConfig.getInstance().isForceGeneration());
    }
}
