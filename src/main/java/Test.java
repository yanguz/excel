import java.util.ArrayList;
import java.util.List;

import writer.Excel2007Writer;

public class Test {
	private static List<String> newTiltes = new ArrayList<String>();
	static {
		newTiltes.add("余额类别");
		newTiltes.add("到期日");
		newTiltes.add("客户号");
		newTiltes.add("或有余额-折美元");
		newTiltes.add("信用证最迟装期");
		newTiltes.add("或有余额-原币");
		newTiltes.add("币种");
		newTiltes.add("生效日");
		newTiltes.add("保证金-折美元");
		newTiltes.add("业务编号");
		newTiltes.add("子编号");
		newTiltes.add("兑付方式");
		newTiltes.add("保证金-原币");
		newTiltes.add("客户名称");
		newTiltes.add("远期天数");
		newTiltes.add("保证金比例");
		newTiltes.add("相关人");
	}

	/**
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception {
		long start = System.currentTimeMillis();
		// 生成xml文件
		String sourceExcel = "D:\\张杰.xlsx";
		String destExcel = "D:\\aaa.xlsx";
		Excel2007Writer.sortSheet1(sourceExcel, destExcel, newTiltes);

		System.out.println((System.currentTimeMillis() - start) / 1000.0 + "s");

	}
}
