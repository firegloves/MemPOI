/**
 * this is a transformation function to apply to the ResultSet record's value to obtain the final value to set in the poi Cell
 * the value received by the db is assumed to be a Boolean
 */
package it.firegloves.mempoi.domain.datatransformation;

import it.firegloves.mempoi.exception.MempoiException;
import lombok.Setter;


@Setter
public abstract class BooleanDataTransformationFunction<O> extends DataTransformationFunction<Boolean, O> {

    @Override
    public final O applyTransformation(Object value) {
        Boolean castValue = super.<Boolean>cast(value, Boolean.class);
        return this.transform(castValue);
    }

    public abstract O transform(final Boolean value) throws MempoiException;
}
