/**
 * this is a transformation function to apply to the ResultSet record's value to obtain the final value to set in the poi Cell
 * the value received by the db is assumed to be a Double
 */
package it.firegloves.mempoi.domain.datatransformation;

import it.firegloves.mempoi.exception.MempoiException;
import lombok.Setter;


@Setter
public abstract class DoubleDataTransformationFunction<O> extends DataTransformationFunction<Double, O> {

    @Override
    public final O applyTransformation(Object value) {
        Double castValue = super.<Double>cast(value, Double.class);
        return this.transform(castValue);
    }

    public abstract O transform(final Double value) throws MempoiException;
}
