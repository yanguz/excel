package writer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import reader.Excel2007Reader;

public class Excel2007Writer {

	/**
	 * 对excel的sheet1进行重新排列
	 * 
	 * @param sourceExcel
	 * @param destExcel
	 * @param newTitleList
	 * @throws Exception
	 */
	public static void sortSheet1(String sourceExcel, String destExcel, List<String> newTitleList) throws Exception {

		Excel2007Reader reader = new Excel2007Reader();
		File tmp = reader.converSheet1NewXml(sourceExcel, newTitleList);
		ZipFile zip = new ZipFile(sourceExcel);

		FileOutputStream out = new FileOutputStream(new File(destExcel));
		ZipOutputStream zos = new ZipOutputStream(out);

		@SuppressWarnings("unchecked")
		Enumeration<ZipEntry> en = (Enumeration<ZipEntry>) zip.entries();
		while (en.hasMoreElements()) {
			ZipEntry ze = en.nextElement();
			zos.putNextEntry(new ZipEntry(ze.getName()));
			InputStream is = null;
			is = zip.getInputStream(ze);
			// sheet1.xml文件为新的
			if (ze.getName().endsWith("sheet1.xml")) {
				is = new FileInputStream(tmp);
			} else {
				is = zip.getInputStream(ze);
			}
			copyStream(is, zos);
			is.close();
		}

		zip.close();
		zos.close();
	}

	private static void copyStream(InputStream in, OutputStream out) throws IOException {
		byte[] chunk = new byte[1024];
		int count;
		while ((count = in.read(chunk)) >= 0) {
			out.write(chunk, 0, count);
		}
	}

}