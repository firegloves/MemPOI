package it.firegloves.mempoi.config;

import lombok.Data;
import lombok.experimental.Accessors;

@Data
@Accessors(chain = true)
public final class MempoiConfig {

	private static MempoiConfig instance = new MempoiConfig();

	private boolean debug;
	private boolean forceGeneration;

	private MempoiConfig() {
	}

	public static MempoiConfig getInstance() {
		return instance;
	}
}
