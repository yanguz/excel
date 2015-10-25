package excel.read;


public class Test {
	public static void main(String[] args) {
		try {
			new ExcelUtil().readOneSheet("D:\\张杰.xlsx");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
