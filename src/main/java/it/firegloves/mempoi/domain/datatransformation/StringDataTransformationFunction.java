/**
 * this is a transformation function to apply to the ResultSet record's value to obtain the final value to set in the
 * poi Cell
 * the value received by the db is assumed to be a String
 */
package it.firegloves.mempoi.domain.datatransformation;

import java.sql.ResultSet;
import lombok.Setter;

@Setter
public abstract class StringDataTransformationFunction<O> extends DataTransformationFunction<String, O> {

    @Override
    public final O execute(final ResultSet rs, Object value) {
        return super.applyTransformation(rs, value, String.class);
    }

}
