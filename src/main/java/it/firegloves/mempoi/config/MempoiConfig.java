package it.firegloves.mempoi.config;

import java.util.Optional;

import lombok.Data;
import lombok.experimental.Accessors;

@Data
@Accessors(chain = true)
public final class MempoiConfig {

	private static MempoiConfig instance = new MempoiConfig();

	private boolean debug;
	private boolean forceGeneration;

	private MempoiConfig() {
		final String debugProperty = System.getProperty("mempoi.debug", System.getenv("MEMPOI_DEBUG"));
		this.debug = Optional.ofNullable(debugProperty).map(Boolean::valueOf).orElse(Boolean.FALSE);
	}

	public static MempoiConfig getInstance() {
		return instance;
	}
}
