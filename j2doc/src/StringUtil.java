public class StringUtil {
	public static final boolean isEmptyString(String s) {
		return s == null || s.trim().equals("") || "NULL".equalsIgnoreCase(s);
	}

	public static String removeNull(Object o) {
		return removeNull(o, "");
	}

	public static String removeNull(Object o, String s) {
		if (o == null) {
			return s;
		}
		return o.toString().trim();
	}
}