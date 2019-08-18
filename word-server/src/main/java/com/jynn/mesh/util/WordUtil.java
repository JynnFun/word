package com.jynn.mesh.util;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.springframework.beans.BeanUtils;

/**
 * 功能描述:word工具类
 *
 * @author jynn
 * @created 2019年8月15日
 * @version 1.0.0
 */
public class WordUtil {

	/**
	 * 功能描述:word下载
	 *
	 * @param response
	 * @param patientMap
	 * @param list
	 * @param itemList
	 * @param file
	 * @see [相关类/方法](可选)
	 * @since [产品/模块版本](可选)
	 */
	public static final void DownloadWord(HttpServletResponse response, Map<String, Object> patientMap,
			List<Map<String, Object>> list, List<List<String>> itemList, String file) {
		CustomXWPFDocument document = null;
		ServletOutputStream servletOS = null;
		ByteArrayOutputStream ostream = null;
		// 添加表格
		try {
			servletOS = response.getOutputStream();
			ostream = new ByteArrayOutputStream();
			document = new CustomXWPFDocument(POIXMLDocument.openPackage(file));// 生成word文档并读取模板
			// 病人信息
			XWPFTable patientTable = document.getTables().get(0);
			eachTable(document, patientTable.getRows(), patientMap);
			// 病历信息
			// 根据病历数量复制表格
			for (int i = 0; i < list.size(); i++) {
				// 标题
				XWPFParagraph paragraph = document.createParagraph();
				XWPFRun paragraphRun = paragraph.createRun();
				paragraphRun.setText(list.get(i).get("${title}").toString());
				paragraph.setAlignment(ParagraphAlignment.CENTER);

				CTTbl ctTbl = CTTbl.Factory.newInstance(); // 创建新的 CTTbl ， table
				ctTbl.set(document.getTables().get(1).getCTTbl()); // 复制原来的CTTbl
				IBody iBody = document.getTables().get(1).getBody();
				BeanUtils.copyProperties(document.getTables().get(1).getBody(), iBody);
				XWPFTable newTable = new XWPFTable(ctTbl, iBody); // 新增一个table，使用复制好的Cttbl

				List<String> iList = itemList.get(i);
				Integer itemIndex = 0;
				// 新建西医检查信息
				for (String item : iList) {
					// 复制行，主要用于复制样式和重设图片文本
					// 直接新增行还需要手动改样式，比较繁琐
					XWPFTableRow titleRow = newTable.createRow();
					// 注意setText方式是在原来文本的后面添加，若不需要原先的文本在需要删除原先的run，新增一个run
					copyTableRow(titleRow, newTable.getRows().get(7), null);
					titleRow.getTableCells().get(0).setText(item);
					XWPFTableRow imageRow = newTable.createRow();
					// 带入序号重设文本
					copyTableRow(imageRow, newTable.getRows().get(8), itemIndex);
					itemIndex++;
				}

				// 删除作为模板的检查项标题和图片行
				newTable.removeRow(7);
				newTable.removeRow(7);
				// 遍历表格,并替换模板

				eachTable(document, newTable.getRows(), list.get(i));
				document.createTable(); // 创建一个空的Table
				// 设置table值
				document.setTable(i + 2, newTable); // 将table设置到word中
			}
			List<XWPFTable> tables = document.getTables();
			// 删除作为模板的第一个表格
			for (int i = tables.get(1).getRows().size(); i >= 0; i--) {
				tables.get(1).removeRow(i);
			}
			// 输出word内容文件流，提供下载
			response.setContentType("application/x-msdownload");
			String name = java.net.URLEncoder.encode("病历.docx", "UTF8");
			name = new String((name).getBytes("UTF-8"), "ISO-8859-1");
			response.addHeader("Content-Disposition", "attachment; filename*=utf-8'zh_cn'" + name);
			document.write(ostream);
			servletOS.write(ostream.toByteArray());
		} catch (Exception e) {
			System.out.print(e.getMessage());
		} finally {
			try {
				if (ostream != null) {
					ostream.close();
				}
				if (servletOS != null) {
					servletOS.close();
				}
			} catch (IOException e) {

			}
		}
	}

	/**
	 * 功能描述:复制单元格，从source到target
	 *
	 * @param target
	 * @param source
	 * @param index
	 * @see [相关类/方法](可选)
	 * @since [产品/模块版本](可选)
	 */
	public static void copyTableCell(XWPFTableCell target, XWPFTableCell source, Integer index) {
		// 列属性
		if (source.getCTTc() != null) {
			target.getCTTc().setTcPr(source.getCTTc().getTcPr());
		}
		// 删除段落
		for (int pos = 0; pos < target.getParagraphs().size(); pos++) {
			target.removeParagraph(pos);
		}
		// 添加段落
		for (XWPFParagraph sp : source.getParagraphs()) {
			XWPFParagraph targetP = target.addParagraph();
			copyParagraph(targetP, sp, index);
		}
	}

	/**
	 * 功能描述:复制段落，从source到target
	 *
	 * @param target
	 * @param source
	 * @param index
	 * @see [相关类/方法](可选)
	 * @since [产品/模块版本](可选)
	 */
	public static void copyParagraph(XWPFParagraph target, XWPFParagraph source, Integer index) {

		// 设置段落样式
		target.getCTP().setPPr(source.getCTP().getPPr());

		// 移除所有的run
		for (int pos = target.getRuns().size() - 1; pos >= 0; pos--) {
			target.removeRun(pos);
		}

		// copy 新的run
		for (XWPFRun s : source.getRuns()) {
			XWPFRun targetrun = target.createRun();
			copyRun(targetrun, s, index);
		}

	}

	/**
	 * 功能描述:复制RUN，从source到target
	 *
	 * @param target
	 * @param source
	 * @param index
	 * @see [相关类/方法](可选)
	 * @since [产品/模块版本](可选)
	 */
	public static void copyRun(XWPFRun target, XWPFRun source, Integer index) {
		// 设置run属性
		target.getCTR().setRPr(source.getCTR().getRPr());
		// 设置文本
		String tail = "";
		if (index != null) {
			tail = index.toString();
		}
		target.setText(source.text().replace("}", "") + tail + "}");
	}

	/**
	 * 功能描述:复制行，从source到target
	 *
	 * @param target
	 * @param source
	 * @param index
	 * @see [相关类/方法](可选)
	 * @since [产品/模块版本](可选)
	 */
	public static void copyTableRow(XWPFTableRow target, XWPFTableRow source, Integer index) {
		// 复制样式
		if (source.getCtRow() != null) {
			target.getCtRow().setTrPr(source.getCtRow().getTrPr());
		}
		// 复制单元格
		for (int i = 0; i < source.getTableCells().size(); i++) {
			XWPFTableCell cell1 = target.getCell(i);
			XWPFTableCell cell2 = source.getCell(i);
			if (cell1 == null) {
				cell1 = target.addNewTableCell();
			}
			copyTableCell(cell1, cell2, index);
		}
	}

	/**
	 * 功能描述:遍历表格，替换信息
	 *
	 * @param document
	 * @param rows
	 * @param textMap
	 * @see [相关类/方法](可选)
	 * @since [产品/模块版本](可选)
	 */
	public static void eachTable(CustomXWPFDocument document, List<XWPFTableRow> rows, Map<String, Object> textMap) {
		for (XWPFTableRow row : rows) {
			List<XWPFTableCell> cells = row.getTableCells();
			for (XWPFTableCell cell : cells) {
				// 判断单元格是否需要替换
				if (checkText(cell.getText())) {
					List<XWPFParagraph> paragraphs = cell.getParagraphs();
					for (XWPFParagraph paragraph : paragraphs) {
						List<XWPFRun> runs = paragraph.getRuns();
						for (XWPFRun run : runs) {
							Object ob = changeValue(run.toString(), textMap);
							if (ob instanceof String) {
								run.setText((String) ob, 0);
							} else if (ob instanceof Map) {
								run.setText("", 0);
								Map pic = (Map) ob;
								int width = Integer.parseInt(pic.get("width").toString());
								int height = Integer.parseInt(pic.get("height").toString());
								int picType = getPictureType(pic.get("type").toString());
								String urls = pic.get("content").toString();

								String[] urlList = urls.split(";");
								for (String url : urlList) {
									ByteArrayInputStream byteInputStream;
									try {
										//网络图片取文件数据
										byteInputStream = new ByteArrayInputStream(getImageData(url));
										document.addPictureData(byteInputStream, picType);
										int id2 = document.getAllPackagePictures().size() - 1;
										document.createPicture(id2, width, height, paragraph);
									} catch (Exception e) {
										e.printStackTrace();
									}
								}
							}
							break;
						}
					}
				}
			}
		}
	}

	/**
	 * 功能描述:读取线上图片文件流
	 *
	 * @param strUrl
	 * @return
	 * @see [相关类/方法](可选)
	 * @since [产品/模块版本](可选)
	 */
	public static byte[] getImageData(String strUrl) {
		InputStream inStream = null;
		try {
			// new一个URL对象
			URL url = new URL(strUrl);
			// 打开链接
			HttpURLConnection conn = (HttpURLConnection) url.openConnection();
			// 设置请求方式为"GET"
			conn.setRequestMethod("GET");
			// 超时响应时间为5秒
			conn.setConnectTimeout(10 * 1000);
			// 通过输入流获取图片数据
			inStream = conn.getInputStream();
			byte[] data = readInputStream(inStream);
			return data;
		} catch (Exception e) {
			return null;
		} finally {
			if (inStream != null) {
				try {
					inStream.close();
				} catch (Exception e2) {
					System.out.println("关闭流失败");
				}
			}
		}

	}

	/**
	 * 功能描述:读取文件流
	 *
	 * @param inStream
	 * @return
	 * @throws Exception
	 * @see [相关类/方法](可选)
	 * @since [产品/模块版本](可选)
	 */
	public static byte[] readInputStream(InputStream inStream) throws Exception {
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		// 创建一个Buffer字符串
		byte[] buffer = new byte[1024];
		// 每次读取的字符串长度，如果为-1，代表全部读取完毕
		int len = 0;
		// 使用一个输入流从buffer里把数据读取出来
		while ((len = inStream.read(buffer)) != -1) {
			// 用输出流往buffer里写入数据，中间参数代表从哪个位置开始读，len代表读取的长度
			outStream.write(buffer, 0, len);
		}
		// 关闭输入流
		inStream.close();
		// 把outStream里的数据写入内存
		return outStream.toByteArray();
	}

	/**
	 * 功能描述:为表格插入数据，行数不够添加新行
	 *
	 * @param table
	 * @param tableList
	 * @param daList
	 * @param type
	 * @see [相关类/方法](可选)
	 * @since [产品/模块版本](可选)
	 */
	public static void insertTable(XWPFTable table, List<String> tableList, List<String[]> daList, Integer type) {
		if (2 == type) {
			// 创建行和创建需要的列
			for (int i = 1; i < daList.size(); i++) {
				// 添加一个新行
				XWPFTableRow row = table.insertNewTableRow(1);
				for (int k = 0; k < daList.get(0).length; k++) {
					row.createCell();// 根据String数组第一条数据的长度动态创建列
				}
			}

			// 创建行,根据需要插入的数据添加新行，不处理表头
			for (int i = 0; i < daList.size(); i++) {
				List<XWPFTableCell> cells = table.getRow(i + 1).getTableCells();
				for (int j = 0; j < cells.size(); j++) {
					XWPFTableCell cell02 = cells.get(j);
					cell02.setText(daList.get(i)[j]);
				}
			}
		} else if (4 == type) {
			// 插入表头下面第一行的数据
			for (int i = 0; i < tableList.size(); i++) {
				XWPFTableRow row = table.createRow();
				List<XWPFTableCell> cells = row.getTableCells();
				cells.get(0).setText(tableList.get(i));
			}
		}
	}

	/**
	 * 功能描述:判断文本中时候包含$
	 *
	 * @param text
	 * @return
	 * @see [相关类/方法](可选)
	 * @since [产品/模块版本](可选)
	 */
	public static boolean checkText(String text) {
		boolean check = false;
		if (text.indexOf("$") != -1) {
			check = true;
		}
		return check;
	}

	/**
	 * 功能描述:匹配传入信息集合与模板
	 *
	 * @param value
	 * @param textMap
	 * @return
	 * @see [相关类/方法](可选)
	 * @since [产品/模块版本](可选)
	 */
	public static Object changeValue(String value, Map<String, Object> textMap) {
		Set<Map.Entry<String, Object>> textSets = textMap.entrySet();
		Object valu = "";
		for (Map.Entry<String, Object> textSet : textSets) {
			// 匹配模板与替换值 格式${key}
			String key = textSet.getKey();
			if (value.indexOf(key) != -1) {
				valu = textSet.getValue();
			}
		}
		return valu;
	}

	/**
	 * 功能描述:根据图片类型，取得对应的图片类型代码
	 *
	 * @param picType
	 * @return
	 * @see [相关类/方法](可选)
	 * @since [产品/模块版本](可选)
	 */
	private static int getPictureType(String picType) {
		int res = CustomXWPFDocument.PICTURE_TYPE_PICT;
		if (picType != null) {
			if (picType.equalsIgnoreCase("png")) {
				res = CustomXWPFDocument.PICTURE_TYPE_PNG;
			} else if (picType.equalsIgnoreCase("dib")) {
				res = CustomXWPFDocument.PICTURE_TYPE_DIB;
			} else if (picType.equalsIgnoreCase("emf")) {
				res = CustomXWPFDocument.PICTURE_TYPE_EMF;
			} else if (picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")) {
				res = CustomXWPFDocument.PICTURE_TYPE_JPEG;
			} else if (picType.equalsIgnoreCase("wmf")) {
				res = CustomXWPFDocument.PICTURE_TYPE_WMF;
			}
		}
		return res;
	}

}