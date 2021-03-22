/**
 * mapping to collect ResultSet data useful to configure MempoiColumns for the reflection
 *
 * To better explain the properties of this class let's use an example. In the code:
 * Date dateValue = resultSet.getDate("date_col_name")
 * These are the relative properties values:
 * rsAccessDataMethodName = "getDate" - to invoke the right method using the reflection
 * rsAccessParamClass = "String" because the method "getDate" accept a String (in this case "date_col_name")
 * rsReturnClass = Date because the invoked ResultSet's method return a Date object
 */
package it.firegloves.mempoi.domain;

import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.Date;

@Getter
@AllArgsConstructor
public enum EExportDataType {

    TEXT("getString", String.class, String.class),
    INT("getInt", String.class, double.class),
    DOUBLE("getDouble", String.class, double.class),
    BIG_INTEGER("getDouble", String.class, double.class),
    FLOAT("getDouble", String.class, double.class),
    CURRENCY("getBigDecimal", String.class, double.class),
    DATE("getDate", String.class, Date.class),
    TIME("getTime", String.class, Date.class),
    TIMESTAMP("getTimestamp", String.class, Date.class),
    BOOLEAN("getBoolean", String.class, boolean.class);

    /**
     * the method to call on the ResultSet to extract the desired data
     */
    private String rsAccessDataMethodName;
    /**
     * the param type to pass to the method identified by the rsAccessDataMethodName property
     */
    private Class<?> rsAccessParamClass;
    /**
     * the return type class of the resultSet's method identified by the properies rsAccessDataMethodName and rsAccessParamClass
     */
    private Class<?> rsReturnClass;
}
