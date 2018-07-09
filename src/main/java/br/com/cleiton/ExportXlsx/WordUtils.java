package br.com.cleiton.ExportXlsx;

public class WordUtils {

	public static String capitalize(String str) {
	return capitalize(str, null);
	}
	private static String capitalize(String str, char... delimiters) {
		int delimLen = delimiters == null ? -1 : delimiters.length;
		if (isEmpty(str) || delimLen == 0) {
			return str;
		}
		char[] buffer = str.toCharArray();
		boolean capitalizeNext = true;
		for (int i = 0; i < buffer.length; i++) {
			char ch = buffer[i];
			if (isDelimiter(ch, delimiters)) {
				capitalizeNext = true;
			} else if (capitalizeNext) {
				buffer[i] = Character.toTitleCase(ch);
				capitalizeNext = false;
			}
		}
		return new String(buffer);
	}

	private static boolean isDelimiter(char ch, char[] delimiters) {
		if (delimiters == null) {
			return Character.isWhitespace(ch);
		}
		for (char delimiter : delimiters) {
			if (ch == delimiter) {
				return true;
			}
		}
		return false;
	}

	public static boolean isEmpty(String str) {
		return str == null || str.length() == 0;
	}
}
