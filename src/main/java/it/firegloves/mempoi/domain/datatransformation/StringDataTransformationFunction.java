/**
 * this is a transformation function to apply to the ResultSet record's value to obtain the final value to set in the poi Cell
 * the value received by the db is assumed to be a String
 */
package it.firegloves.mempoi.domain.datatransformation;

import it.firegloves.mempoi.exception.MempoiException;
import lombok.Setter;


@Setter
public abstract class StringDataTransformationFunction<O> extends DataTransformationFunction<String, O> {

    @Override
    public final O applyTransformation(Object value) {
        String castValue = super.<String>cast(value, String.class);
        return this.transform(castValue);
    }

    public abstract O transform(final String value) throws MempoiException;
}
