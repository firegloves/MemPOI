package it.firegloves.mempoi.domain;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.Getter;
import lombok.experimental.Accessors;

import java.util.Date;

@Getter
@AllArgsConstructor
public enum EExportDataType {

    TEXT("getString", String.class, String.class),
    INT("getInt", String.class, double.class),
    DOUBLE("getDouble", String.class, double.class),
    FLOAT("getDouble", String.class, double.class),
    CURRENCY("getBigDecimal", String.class, double.class),
    DATE("getDate", String.class, Date.class),
    TIME("getTime", String.class, Date.class),
    TIMESTAMP("getTimestamp", String.class, Date.class),
    BOOLEAN("getBoolean", String.class, boolean.class);

    private String rsAccessDataMethodName;
    private Class rsAccessParamClass;
    private Class rsReturnClass;
}
