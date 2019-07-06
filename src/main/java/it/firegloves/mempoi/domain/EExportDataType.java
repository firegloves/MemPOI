package it.firegloves.mempoi.domain;

import java.util.Date;

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

    EExportDataType(String rsAccessDataMethodName, Class rsAccessParamClass, Class rsReturnClass) {
        this.rsAccessDataMethodName = rsAccessDataMethodName;
        this.rsAccessParamClass = rsAccessParamClass;
        this.rsReturnClass = rsReturnClass;
    }

    public String getRsAccessDataMethodName() {
        return rsAccessDataMethodName;
    }

    public Class getRsAccessParamClass() {
        return rsAccessParamClass;
    }

    public Class getRsReturnClass() {
        return rsReturnClass;
    }
}
