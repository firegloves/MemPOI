/**
 * this is a transformation function to apply to the ResultSet record's value to obtain the final value to set in the poi Cell
 */
package it.firegloves.mempoi.domain;

@FunctionalInterface
public interface DataTransformationFunction<T> {

    T apply(Object value);
}
