/**
 * enum containing the default data type formats
 */

package it.firegloves.mempoi.styles;

public enum StandardDataFormat {

    STANDARD_DATE_FORMAT("yyyy/mm/dd"),
    STANDARD_DATETIME_FORMAT("yyyy/mm/dd h:mm"),
    STANDARD_FLOATING_NUMBER_FORMAT("#,##0.00");

    private String format;

    StandardDataFormat(String format) {
        this.format = format;
    }

    public String getFormat() {
        return format;
    }
}
