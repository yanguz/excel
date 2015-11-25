package excel.utils;


public class BeanUtil {
	/**
	 * 获取数据流编码格式
	 * 
	 * @param bytes
	 *            数据流
	 * @return
	 */
	public static String getEncoding(byte[] bytes) {
		CharsetDetector detector = new CharsetDetector();
		detector.setText(bytes);
		CharsetMatch cm = detector.detect();
		String encoding = cm.getName();
		return encoding;
	}

}
