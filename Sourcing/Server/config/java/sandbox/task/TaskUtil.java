package sandbox.task;

import java.util.List;
import java.util.Map;

import ariba.util.core.Fmt;
import ariba.util.core.ListUtil;
import ariba.util.core.StringUtil;

/**
 *
 * @author fhenri
 */
@SuppressWarnings("unchecked")
public class TaskUtil {

    /**
     * Return a list parameter
     *
     * @param parameter
     * @param parameters
     * @param isRequired
     * @return
     */
    public static List getListParameter(String parameter, Map parameters, boolean isRequired) {
        Object value = parameters.get(parameter);
        if (value == null) {
            if (isRequired)
                throw new IllegalArgumentException(Fmt.S("%s is required", parameter));
            else
                return ListUtil.list();
        }
        if (!(value instanceof List))
            throw new IllegalArgumentException(Fmt.S("Expected list for %s",parameter));
        return (List) value;
    }

    /**
     * Utility method to retrieve a string parameter from the scheduled
     * tasks table file. An {@link IllegalArgumentException} is thrown if the
     * parameters is missing and required or empty.
     *
     * @param scheduledTaskName
     * @param parameter
     * @param parameters
     * @param isRequired is this parameter required for the task
     * @return
     */
    public static String getStringParameter (
            String scheduledTaskName,
            String parameter,
            Map parameters,
            boolean isRequired) {

        Object value = parameters.get(parameter);
        // The parameter must be present
        if (value == null)
            if (isRequired) {
                throw new IllegalArgumentException(
                        Fmt.S("%s expects the parameter %s", scheduledTaskName, parameter));
            } else {
                return null;
            }
        // The parameter must be a String
        if (!(value instanceof String))
            throw new IllegalArgumentException(
                    Fmt.S("%s: %s must be a string", scheduledTaskName, parameter));
        // The parameter must be non-empty
        if (isRequired && StringUtil.nullOrEmptyOrBlankString((String) value))
            throw new IllegalArgumentException(
                    Fmt.S("%s: %s must be a non-empty string", scheduledTaskName, parameter));
        return (String) value;
    }

    /**
     * Utility method to retrieve a boolean parameter from the scheduled
     * tasks table file. An {@link IllegalArgumentException} is thrown if the
     * parameters is missing and required or empty.
     *
     * @param scheduledTaskName
     * @param parameter
     * @param parameters
     * @param isRequired is this parameter required for the task
     * @return
     */
    public static boolean getBooleanParameter (
            String scheduledTaskName,
            String parameter,
            Map parameters,
            boolean isRequired) {

        String value = getStringParameter(scheduledTaskName, parameter, parameters, isRequired);
        return Boolean.valueOf(value).booleanValue();
    }

}
