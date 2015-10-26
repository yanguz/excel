package reader;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.Writer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class Excel2007Reader extends DefaultHandler {

	/**
	 * 共享字符串表
	 */
	private SharedStringsTable sst;
	/**
	 * 上一次的内容
	 */
	private String lastContents;
	/**
	 * 字符串标识
	 */
	private boolean nextIsString;

	/**
	 * 行集合
	 */
	private List<Map<String, String>> rowlist = new ArrayList<Map<String, String>>();
	/**
	 * 标题序号
	 */
	private Map<String, Integer> titleMap = new HashMap<String, Integer>();
	/**
	 * 新的标题序号：外部传入
	 */
	private List<String> newTiltes = new ArrayList<String>();
	/**
	 * 当前单元格
	 */
	private Map<String, String> curCell;
	/**
	 * 当前行
	 */
	private int curRow = 0;
	/**
	 * 当前列
	 */
	private int curCol = 0;
	/**
	 * 第一行标志
	 */
	private boolean firstRow = true;

	/**
	 * 写文件
	 */
	private Writer fw;
	/**
	 * sheet1.xml
	 */
	private File sheet1Xml;

	/**
	 * 读取第一个工作簿的sheet1.xml文件
	 * 
	 * @param path
	 * @param newTiltes
	 * @return
	 * @throws Exception
	 */
	public File converSheet1NewXml(String path, List<String> newTiltes) throws Exception {
		// 新的表头顺序
		this.newTiltes = newTiltes;

		// 生成xml文件
		sheet1Xml = File.createTempFile("sheet1", "xml");
		fw = new FileWriter(sheet1Xml);
		fw.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");

		// 读取excel文件
		OPCPackage pkg = OPCPackage.open(path);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();
		XMLReader parser = fetchSheetParser(sst);
		InputStream sheet = r.getSheet("rId1");
		InputSource sheetSource = new InputSource(sheet);
		parser.parse(sheetSource);
		sheet.close();

		fw.close();

		return sheet1Xml;
	}

	/**
	 * <row>中的<c>数据排序
	 */
	public String getNewRowOrder() {
		StringBuilder sb = new StringBuilder();
		for (int i = 0, size = newTiltes.size(); i < size; i++) {
			// 找新的标题
			String key = newTiltes.get(i);
			// 找新标题的序号
			Integer index = titleMap.get(key);
			// 找到老的单元格
			Map<String, String> oldCell = rowlist.get(i);
			// 找到新的单元格
			Map<String, String> newCell = rowlist.get(index);
			// 拼接新的<row>
			if (newCell != null && oldCell != null) {
				sb.append("<c ");
				for (String k : newCell.keySet()) {
					// v不在attribute里面，而是下一个子元素
					if (!"v".equals(k)) {
						sb.append(k).append("=\"");
						// r的内容需要从老单元格获取
						if ("r".equals(k)) {
							sb.append(oldCell.get(k));
						} else {
							sb.append(newCell.get(k));
						}
						sb.append("\" ");
					}
				}
				sb.append(">");
				String v = newCell.get("v");
				// 添加v子元素
				if (v != null) {
					sb.append("<v>").append(v).append("</v>");
				}
				sb.append("</c>");
			}
		}
		rowlist.clear();
		return sb.toString();
	}

	public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
		XMLReader parser = XMLReaderFactory.createXMLReader();
		this.sst = sst;
		parser.setContentHandler(this);
		return parser;
	}

	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		// 需要写的xml内容
		String xmlContent = null;
		// 解析c单元格
		if ("c".equals(name)) {
			// 属性全部添加进去
			curCell = new HashMap<String, String>();
			if (attributes != null && attributes.getLength() > 0) {
				for (int i = 0, size = attributes.getLength(); i < size; i++) {
					curCell.put(attributes.getQName(i), attributes.getValue(i));
				}
			}
			rowlist.add(curCell);
			// 如果下一个元素是 SST 的索引，则将nextIsString标记为true
			String cellType = attributes.getValue("t");
			if (cellType != null && cellType.equals("s")) {
				nextIsString = true;
			} else {
				nextIsString = false;
			}
		} else if ("v".equals(name)) {
			// 什么也不做
		} else if ("worksheet".equals(name)) {
			xmlContent = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">";
		} else {
			StringBuilder sb = new StringBuilder();
			sb.append("<").append(name);
			if (attributes != null && attributes.getLength() > 0) {
				for (int i = 0, size = attributes.getLength(); i < size; i++) {
					sb.append(" ").append(attributes.getQName(i)).append("=\"").append(attributes.getValue(i)).append("\" ");
				}
			}
			sb.append(">");
			xmlContent = sb.toString();
		}

		if (xmlContent != null) {
			try {
				fw.write(xmlContent);
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		// 置空
		lastContents = "";
	}

	public void endElement(String uri, String localName, String name) throws SAXException {
		String xmlContent = null;
		if ("v".equals(name)) {
			curCell.put("v", lastContents);
			if (firstRow) {
				String readVal = null;
				if (nextIsString) {
					int idx = Integer.parseInt(lastContents);
					readVal = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
				} else {
					readVal = lastContents;
				}
				titleMap.put(readVal, curCol);
			}
			curCol++;
		} else if ("c".equals(name)) {

		} else if ("row".equals(name)) {
			if (firstRow) {
				firstRow = false;
			}
			curRow++;
			curCol = 0;

			String newRow = getNewRowOrder();
			xmlContent = newRow + "</" + name + ">";
		} else {
			xmlContent = "</" + name + ">";
		}

		if (xmlContent != null) {
			try {
				fw.write(xmlContent);
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

	}

	public void characters(char[] ch, int start, int length) throws SAXException {
		// 得到单元格内容的值
		lastContents += new String(ch, start, length);
	}

}