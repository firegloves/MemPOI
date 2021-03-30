/**
 * this is a transformation function to apply to the ResultSet record's value to obtain the final value to set in the
 * poi Cell
 * the value received by the db is assumed to be a Date
 */

package it.firegloves.mempoi.domain.datatransformation;

import it.firegloves.mempoi.exception.MempoiException;
import java.util.Date;
import lombok.Setter;

@Setter
public abstract class DateDataTransformationFunction<O> extends DataTransformationFunction<Date, O> {

    @Override
    public O execute(Object value) {
        return super.applyTransformation(value, Date.class);
    }

    public abstract O transform(final Date value) throws MempoiException;
}
