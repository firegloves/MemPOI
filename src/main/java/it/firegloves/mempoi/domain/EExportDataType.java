package it.firegloves.mempoi.domain;

import java.util.Date;
import java.util.EnumSet;

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

    public static EnumSet<EExportDataType> DATE_STYLER_TYPES = EnumSet.of(DATE);
    public static EnumSet<EExportDataType> DATETIME_STYLER_TYPES = EnumSet.of(TIME, TIMESTAMP);
    public static EnumSet<EExportDataType> NUMBER_STYLER_TYPES = EnumSet.of(INT, DOUBLE, FLOAT);

    EExportDataType(String rsAccessDataMethodName, Class rsAccessParamClass, Class rsReturnClass) {
        this.rsAccessDataMethodName = rsAccessDataMethodName;
        this.rsAccessParamClass = rsAccessParamClass;
        this.rsReturnClass = rsReturnClass;
    }

    public String getRsAccessDataMethodName() {
        return rsAccessDataMethodName;
    }

    public void setRsAccessDataMethodName(String rsAccessDataMethodName) {
        this.rsAccessDataMethodName = rsAccessDataMethodName;
    }

    public Class getRsAccessParamClass() {
        return rsAccessParamClass;
    }

    public void setRsAccessParamClass(Class rsAccessParamClass) {
        this.rsAccessParamClass = rsAccessParamClass;
    }

    public Class getRsReturnClass() {
        return rsReturnClass;
    }

    public void setRsReturnClass(Class rsReturnClass) {
        this.rsReturnClass = rsReturnClass;
    }
}
