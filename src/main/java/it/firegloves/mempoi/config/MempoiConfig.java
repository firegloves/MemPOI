package it.firegloves.mempoi.config;

import ch.qos.logback.classic.Level;
import ch.qos.logback.classic.Logger;
import lombok.Data;
import lombok.experimental.Accessors;
import org.slf4j.LoggerFactory;

@Data
@Accessors(chain = true)
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

    private void setLogLevel() {
        Logger rootLogger = (ch.qos.logback.classic.Logger) LoggerFactory.getLogger(org.slf4j.Logger.ROOT_LOGGER_NAME);
        rootLogger.setLevel(this.debug ? Level.DEBUG : Level.INFO);
    }


}
