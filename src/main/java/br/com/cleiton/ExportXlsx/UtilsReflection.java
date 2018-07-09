package br.com.cleiton.ExportXlsx;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.StringTokenizer;


public class UtilsReflection {
	/**
	 * retorna o resultado de getValueOfKeyOnTarget
	 * 
	 * @param element
	 * @param object
	 * @return Object
	 */
	public static Object getMethodCallResult(Element element, Object object) throws IllegalArgumentException {
		StringTokenizer fieldPath = new StringTokenizer(element.getField(), ".");
		Object value = getValueOfKeyOnTarget(fieldPath, object);
		if (value != null)
			return value;
		return null;
	}

	private static Object getValueOfKeyOnTarget(StringTokenizer iterator, Object target)
			throws IllegalArgumentException {
		if (iterator.hasMoreElements()) {
			String nextField = iterator.nextElement().toString();
			if (target == null)
				return null;

			try {
				Method getMethod = target.getClass().getMethod(String.format("get%s", WordUtils.capitalize(nextField)));
				return getValueOfKeyOnTarget(iterator, getMethod.invoke(target));
			} catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException | SecurityException e) {
				e.printStackTrace();
				throw new IllegalArgumentException(String.format(" %s's getter not Found on object of class %s!",
						nextField, target.getClass().getName()));
			}
		} else {
			return target;
		}
	}
}
