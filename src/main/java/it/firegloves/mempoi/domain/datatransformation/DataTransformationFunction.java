/**
 * this is a transformation function to apply to the ResultSet record's value to obtain the final value to set in the poi Cell
 */
package it.firegloves.mempoi.domain.datatransformation;

import it.firegloves.mempoi.exception.MempoiException;

@FunctionalInterface
public interface DataTransformationFunction<I, O> {

    String TRANSFORM_METHOD_NAME = "transform";

    O transform(final I value) throws MempoiException;
}
