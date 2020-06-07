/**
 * created by firegloves
 */

package it.firegloves.mempoi.testutil;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

public class PrivateAccessHelper {


    /**
     * sets as accessible a field of the received object and returns it
     * @param o the object of which set accessible a field
     * @param fieldName the name of the field to make accessible
     * @return the Field now accessible
     * @throws Exception
     */
    public static Field getPrivateField(Object o, String fieldName) throws Exception {
        Field f = o.getClass().getDeclaredField(fieldName);
        f.setAccessible(true);
        return f;
    }


    /**
     * sets a value into an inaccessible field of the received object
     * @param o the object of which set accessible a field
     * @param fieldName the name of the field to make accessible
     * @param value the value to set into the field
     * @throws Exception
     */
    public static void setPrivateField(Object o, String fieldName, Object value) throws Exception {
        Field f = o.getClass().getDeclaredField(fieldName);
        f.setAccessible(true);
        f.set(o, value);
    }

//    /**
//     * inject a value into the desired private field of the received object
//     * @param obj the object from which get the field to set
//     * @param fieldName the name of the private field to set
//     * @param value the value to inject
//     * @throws Exception
//     */
//    public static void setPrivateFieldValue(Object obj, String fieldName, Object value) throws Exception {
//        Field f = setAccessibleField(obj, fieldName);
//        f.set(obj, value);
//    }


    /**
     * sets as accessible a method of the received object and returns it
     * @param o the object of which set accessible a method
     * @param methodName the name of the method to make accessible
     * @param parameterTypes the list of the param classes received by the method to set accessible
     * @return the Method now accessible
     * @throws Exception
     */
    public static Method getAccessibleMethod(Object o, String methodName, Class<?>... parameterTypes) throws Exception {
        return getAccessibleMethod(o.getClass(), methodName, parameterTypes);
    }

    /**
     * sets as accessible a method of the received object and returns it
     * @param clazz the Class of the object of which set accessible a method
     * @param methodName the name of the method to make accessible
     *                   @param parameterTypes the list of the param classes received by the method to set accessible
     * @return the Method now accessible
     * @throws Exception
     */
    public static Method getAccessibleMethod(Class clazz, String methodName, Class<?>... parameterTypes) throws Exception {
        Method m = clazz.getDeclaredMethod(methodName, parameterTypes);
        m.setAccessible(true);
        return m;
    }
}
