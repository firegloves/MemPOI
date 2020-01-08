package it.firegloves.mempoi.config;

import ch.qos.logback.classic.Level;
import ch.qos.logback.classic.Logger;
import org.slf4j.LoggerFactory;

public class MempoiConfig {

    private static MempoiConfig instance = new MempoiConfig();

    private boolean debug = false;
    private boolean forceGeneration = false;

    private MempoiConfig() {
        this.setLogLevel();
    }

    public static MempoiConfig getInstance() {
        return instance;
    }

    public boolean isDebug() {
        return debug;
    }

    public MempoiConfig setDebug(boolean debug) {
        this.debug = debug;
        this.setLogLevel();
        return this;
    }

    public MempoiConfig setForceGeneration(boolean forceGeneration) {
        this.forceGeneration = forceGeneration;
        return this;
    }


    public boolean isForceGeneration() {
        return this.forceGeneration;
    }

    private void setLogLevel() {
        Logger rootLogger = (ch.qos.logback.classic.Logger) LoggerFactory.getLogger(org.slf4j.Logger.ROOT_LOGGER_NAME);
        rootLogger.setLevel(this.debug ? Level.DEBUG : Level.INFO);
    }


}
