/**
 * this is a transformation function to apply to the ResultSet record's value to obtain the final value to set in the
 * poi Cell
 */
package it.firegloves.mempoi.domain.datatransformation;

import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import lombok.Setter;

@Setter
public abstract class DataTransformationFunction<I, O> {

    public static final String TRANSFORM_METHOD_NAME = "transform";

    /**
     * the name of the column against which apply the value transformation
     */
    private String columnName;

    public abstract O execute(final Object value);

    protected O applyTransformation(final Object value, Class<I> clazz) {
        I castValue = this.<I>cast(value, clazz);
        return this.transform(castValue);
    }

    protected final <T> T cast(Object obj, Class<T> toType) throws MempoiException {
        try {
            return toType.cast(obj);
        } catch (Exception e) {
            throw new MempoiException(
                    String.format(Errors.ERR_DATA_TRANSFORMATION_FUNCTION_CAST_EXCEPTION, columnName, toType,
                            obj.getClass()));
        }
    }

    protected abstract O transform(final I value) throws MempoiException;
}
