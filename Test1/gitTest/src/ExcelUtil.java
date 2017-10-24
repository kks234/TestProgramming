

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Random;
import java.util.UUID;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.MessageSource;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.BasicQuery;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.stereotype.Service;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import com.mongodb.BasicDBObject;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;

import pdms.common.info.SessionInfo;
import pdms.common.runtime.SessionHandler;

/**
 * @author kks
 * @date 2017. 9. 29.
 */
@Service
public class ExcelUtil {

	@Value("#{config['main.filePath']}")
	private static String filePath;

	@Autowired
	private static MongoTemplate mongoTemplate;

	public MongoTemplate getMongoTemplate() {
		return mongoTemplate;
	}

	@Autowired
	public void setMongoTemplate(MongoTemplate mongoTemplate) {
		this.mongoTemplate = mongoTemplate;
	}

	@Autowired
	static MessageSource message;

	/**
	 * @param cursor
	 *            : MongoCursor
	 * @param keyList
	 *            : Db에서 가져올 실제 컬럼명
	 * @param keyNameList
	 *            : 엑셀에 표시 되어질 컬럼명
	 * @param usr_id
	 *            : session 사용 중인 사용자 아이디
	 * @return 정상 : 서버에 저장되어진 File 객체, 비정상 : null;
	 * @throws IOException
	 */
	public static File makeExcelFile(DBCursor cursor, List<String> keyList, List<String> keyNameList, String usr_id) throws IOException {

		BufferedWriter bw = null;
		FileWriter bos = null;

		// 파일 명 만들기
		Random ran = new Random();
		int ranInt = ran.nextInt(65536);
		File tempxls = new File(filePath+"/excel", DateUtil.getMMDDHHMISSMS() + "_" + (Integer.toHexString(ranInt)) + "_" + usr_id + ".xls");
		if (!tempxls.getParentFile().exists()) {
			tempxls.getParentFile().mkdirs();
		}
		tempxls.createNewFile();
		try {

			bos = new FileWriter(tempxls);
			try {
				bw = new BufferedWriter(bos);
				StringBuffer sb = new StringBuffer();
				sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n");
				sb.append("<?mso-application progid=\"Excel.Sheet\"?>\r\n");
				sb.append(
						"<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">\r\n");
				// 스타일추가
				sb.append("<Styles>");
				sb.append("  <Style ss:ID=\"Default\" ss:Name=\"Normal\">");
				sb.append("   <Alignment ss:Vertical=\"Center\"/>");
				sb.append("   <Borders/>");
				sb.append("   <Font ss:FontName=\"맑은 고딕\" x:CharSet=\"129\" x:Family=\"Modern\" ss:Size=\"11\"");
				sb.append("    ss:Color=\"#000000\"/>");
				sb.append("   <Interior/>");
				sb.append("   <NumberFormat/>");
				sb.append("   <Protection/>");
				sb.append("  </Style>");
				sb.append("  <Style ss:ID=\"s40\" ss:Name=\"40% - 강조색1\">");
				sb.append("   <Font ss:FontName=\"맑은 고딕\" x:CharSet=\"129\" x:Family=\"Modern\" ss:Size=\"11\"");
				sb.append("    ss:Color=\"#000000\"/>");
				sb.append("   <Interior ss:Color=\"#B8CCE4\" ss:Pattern=\"Solid\"/>");
				sb.append("  </Style>");
				sb.append("  <Style ss:ID=\"s62\">");
				sb.append("   <Borders>");
				sb.append("    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
				sb.append("    <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
				sb.append("    <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
				sb.append("    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
				sb.append("   </Borders>");
				sb.append("  </Style>");
				sb.append("  <Style ss:ID=\"s65\" ss:Parent=\"s40\">");
				sb.append("   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/>");
				sb.append("   <Borders>");
				sb.append("    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"");
				sb.append("     ss:Color=\"#7F7F7F\"/>");
				sb.append("    <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"");
				sb.append("     ss:Color=\"#7F7F7F\"/>");
				sb.append("    <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"");
				sb.append("     ss:Color=\"#7F7F7F\"/>");
				sb.append("    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"");
				sb.append("     ss:Color=\"#7F7F7F\"/>");
				sb.append("   </Borders>");
				sb.append("  </Style>");
				sb.append(" </Styles>");

				int sheetNum = 1;
				long colNum = 1;
				for (; cursor.hasNext(); colNum++) {
					DBObject row = cursor.next();

					try {
   
						if (colNum % 600000 == 1) {
							sb.append("<Worksheet ss:Name=\"Sheet" + (sheetNum++) + "\">\r\n");
							sb.append("<Table>\r\n");
							sb.append("<Column ss:Index=\"" + sheetNum + "\" ss:AutoFitWidth=\"0\" ss:Width=\"110\"/>\r\n");
							sb.append("<Row>\r\n");
							sb.append("<Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">순번</Data></Cell>\r\n");

							// 해더 넣기
							if (keyNameList == null) {

								Iterator<String> keyIter = row.keySet().iterator();

								while (keyIter.hasNext()) {
									String key = keyIter.next();
									sb.append("<Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">" + key + "</Data></Cell>\r\n");
								}

							} else {
								for (String keyName : keyNameList) {
									sb.append("<Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">" + keyName + "</Data></Cell>\r\n");
								}
							}
							sb.append("</Row>\r\n");
						}
						sb.append("<Row>\r\n");
						sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"Number\">" + colNum + "</Data></Cell>\r\n");
						if (keyList == null) {
							Iterator<String> keyIter = row.keySet().iterator();
							while (keyIter.hasNext()) {
								String key = keyIter.next();

								Object val = row.get(key);
								if (val == null) {
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">-</Data></Cell>\r\n");
								} else if (val instanceof Long || val instanceof Integer || val instanceof Double) {
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"Number\">" + val + "</Data></Cell>\r\n");
									// } else if(val instanceof ObjectId){
									// continue;

								} else {
									String valreplace = (val + "").replace("&", "&amp;").replace(">", "&gt;").replace("<", "&lt;").replace("\"", "&quot;").replace("\'", "&apos;");
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + valreplace + "</Data></Cell>\r\n");
								}
							}
						} else {
							for (String key : keyList) {
								Object val = row.get(key);

								if (val == null) {
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">-</Data></Cell>\r\n");
								} else if (val instanceof Long || val instanceof Integer || val instanceof Double) {
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"Number\">" + val + "</Data></Cell>\r\n");
								} else {
									String valreplace = (val + "").replace("&", "&amp;").replace(">", "&gt;").replace("<", "&lt;").replace("\"", "&quot;").replace("\'", "&apos;");
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + valreplace + "</Data></Cell>\r\n");
								}
							}
						}
						sb.append("</Row>\r\n");
						bw.write(sb.toString());

						sb = new StringBuffer();
						// 다음 시트로
						if (colNum % 600000 == 0) {
							sb.append("</Table>\r\n");
							sb.append("</Worksheet>\r\n");
						}

					} catch (Exception e) {
						e.printStackTrace();
					}
				}

				// 마지막 닫기
				if (colNum % 600000 != 1) {
					sb.append("</Table>\r\n");
					sb.append("</Worksheet>\r\n");
				}

				sb.append("</Workbook>\r\n");
				bw.write(sb.toString());
				return tempxls;
			} finally {
				if (bw != null) {
					bw.close();
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
			return null;
		} finally {

			if (bos != null) {
				bos.close();
			}

		}

	}

	// List<String> keyList, List<String> keyNameList, String usr_id
	public static File makeExcelFileForList(List<DBObject> data, String usr_id) throws IOException {
		BufferedWriter bw = null;
		FileWriter bos = null;

		// 파일 명 만들기
		Random ran = new Random();
		int ranInt = ran.nextInt(65536);
		File tempxls = new File("/excel", DateUtil.getMMDDHHMISSMS() + "_" + (Integer.toHexString(ranInt)) + "_" + usr_id + ".xls");
		if (!tempxls.getParentFile().exists()) {
			tempxls.getParentFile().mkdirs();
		}
		tempxls.createNewFile();
		try {

			bos = new FileWriter(tempxls);
			try {
				bw = new BufferedWriter(bos);
				StringBuffer sb = new StringBuffer();
				sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n");
				sb.append("<?mso-application progid=\"Excel.Sheet\"?>\r\n");
				sb.append(
						"<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">\r\n");
				// 스타일추가
				sb.append("<Styles>");
				sb.append("  <Style ss:ID=\"Default\" ss:Name=\"Normal\">");
				sb.append("   <Alignment ss:Vertical=\"Center\"/>");
				sb.append("   <Borders/>");
				sb.append("   <Font ss:FontName=\"맑은 고딕\" x:CharSet=\"129\" x:Family=\"Modern\" ss:Size=\"11\"");
				sb.append("    ss:Color=\"#000000\"/>");
				sb.append("   <Interior/>");
				sb.append("   <NumberFormat/>");
				sb.append("   <Protection/>");
				sb.append("  </Style>");
				sb.append("  <Style ss:ID=\"s40\" ss:Name=\"40% - 강조색1\">");
				sb.append("   <Font ss:FontName=\"맑은 고딕\" x:CharSet=\"129\" x:Family=\"Modern\" ss:Size=\"11\"");
				sb.append("    ss:Color=\"#000000\"/>");
				sb.append("   <Interior ss:Color=\"#B8CCE4\" ss:Pattern=\"Solid\"/>");
				sb.append("  </Style>");
				sb.append("  <Style ss:ID=\"s62\">");
				sb.append("   <Borders>");
				sb.append("    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
				sb.append("    <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
				sb.append("    <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
				sb.append("    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
				sb.append("   </Borders>");
				sb.append("  </Style>");
				sb.append("  <Style ss:ID=\"s65\" ss:Parent=\"s40\">");
				sb.append("   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/>");
				sb.append("   <Borders>");
				sb.append("    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"");
				sb.append("     ss:Color=\"#7F7F7F\"/>");
				sb.append("    <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"");
				sb.append("     ss:Color=\"#7F7F7F\"/>");
				sb.append("    <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"");
				sb.append("     ss:Color=\"#7F7F7F\"/>");
				sb.append("    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"");
				sb.append("     ss:Color=\"#7F7F7F\"/>");
				sb.append("   </Borders>");
				sb.append("  </Style>");
				sb.append(" </Styles>");

				int sheetNum = 1;
				long colNum = 1;

				for (DBObject row : data) {
					// DBObject row = cursor.next();

					try {

						if (colNum % 600000 == 1) {
							sb.append("<Worksheet ss:Name=\"Sheet" + (sheetNum++) + "\">\r\n");
							sb.append("<Table>\r\n");
							sb.append("<Column ss:Index=\"" + sheetNum + "\" ss:AutoFitWidth=\"0\" ss:Width=\"110\"/>\r\n");
							sb.append("<Row>\r\n");
							sb.append("<Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">순번</Data></Cell>\r\n");

							// 해더 넣기
							{

								Iterator<String> keyIter = row.keySet().iterator();
								while (keyIter.hasNext()) {
									String key = keyIter.next();
									sb.append("<Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">" + key + "</Data></Cell>\r\n");
								}
							}
							sb.append("</Row>\r\n");
						}
						sb.append("<Row>\r\n");
						sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"Number\">" + colNum + "</Data></Cell>\r\n");
						{
							Iterator<String> keyIter = row.keySet().iterator();
							while (keyIter.hasNext()) {
								String key = keyIter.next();

								Object val = row.get(key);

								if (val == null) {
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">-</Data></Cell>\r\n");
								} else if (val instanceof Long || val instanceof Integer || val instanceof Double) {
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"Number\">" + val + "</Data></Cell>\r\n");
									// } else if(val instanceof ObjectId){
									// continue;

								} else {
									String valreplace = (val + "").replace("&", "&amp;").replace(">", "&gt;").replace("<", "&lt;").replace("\"", "&quot;").replace("\'", "&apos;");
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + valreplace + "</Data></Cell>\r\n");
								}
							}
						}
						sb.append("</Row>\r\n");
						bw.write(sb.toString());

						sb = new StringBuffer();
						// 다음 시트로
						if (colNum % 600000 == 0) {
							sb.append("</Table>\r\n");
							sb.append("</Worksheet>\r\n");
						}

					} catch (Exception e) {
						e.printStackTrace();
					}
					colNum++;
				}

				// 마지막 닫기
				if (colNum % 600000 != 1) {
					sb.append("</Table>\r\n");
					sb.append("</Worksheet>\r\n");
				}

				sb.append("</Workbook>\r\n");
				bw.write(sb.toString());
				return tempxls;
			} finally {
				if (bw != null) {
					bw.close();
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
			return null;
		} finally {

			if (bos != null) {
				bos.close();
			}

		}

	}

	public static File makeExcelFile(Iterator<DBObject> cursor, List<String> keyList, List<String> keyNameList, String usr_id) throws IOException {
		BufferedWriter bw = null;
		FileWriter bos = null;

		// 파일 명 만들기
		Random ran = new Random();
		int ranInt = ran.nextInt(65536);
		File tempxls = new File("/excel", DateUtil.getMMDDHHMISSMS() + "_" + (Integer.toHexString(ranInt)) + "_" + usr_id + ".xls");
		tempxls.createNewFile();
		try {

			bos = new FileWriter(tempxls);
			try {
				bw = new BufferedWriter(bos);
				StringBuffer sb = new StringBuffer();
				sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n");
				sb.append("<?mso-application progid=\"Excel.Sheet\"?>\r\n");
				sb.append(
						"<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">\r\n");
				// 스타일추가
				sb.append("<Styles>");
				sb.append("  <Style ss:ID=\"Default\" ss:Name=\"Normal\">");
				sb.append("   <Alignment ss:Vertical=\"Center\"/>");
				sb.append("   <Borders/>");
				sb.append("   <Font ss:FontName=\"맑은 고딕\" x:CharSet=\"129\" x:Family=\"Modern\" ss:Size=\"11\"");
				sb.append("    ss:Color=\"#000000\"/>");
				sb.append("   <Interior/>");
				sb.append("   <NumberFormat/>");
				sb.append("   <Protection/>");
				sb.append("  </Style>");
				sb.append("  <Style ss:ID=\"s40\" ss:Name=\"40% - 강조색1\">");
				sb.append("   <Font ss:FontName=\"맑은 고딕\" x:CharSet=\"129\" x:Family=\"Modern\" ss:Size=\"11\"");
				sb.append("    ss:Color=\"#000000\"/>");
				sb.append("   <Interior ss:Color=\"#B8CCE4\" ss:Pattern=\"Solid\"/>");
				sb.append("  </Style>");
				sb.append("  <Style ss:ID=\"s62\">");
				sb.append("   <Borders>");
				sb.append("    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
				sb.append("    <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
				sb.append("    <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
				sb.append("    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>");
				sb.append("   </Borders>");
				sb.append("  </Style>");
				sb.append("  <Style ss:ID=\"s65\" ss:Parent=\"s40\">");
				sb.append("   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/>");
				sb.append("   <Borders>");
				sb.append("    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"");
				sb.append("     ss:Color=\"#7F7F7F\"/>");
				sb.append("    <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"");
				sb.append("     ss:Color=\"#7F7F7F\"/>");
				sb.append("    <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"");
				sb.append("     ss:Color=\"#7F7F7F\"/>");
				sb.append("    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"");
				sb.append("     ss:Color=\"#7F7F7F\"/>");
				sb.append("   </Borders>");
				sb.append("  </Style>");
				sb.append(" </Styles>");

				int sheetNum = 1;
				long colNum = 1;
				for (; cursor.hasNext(); colNum++) {
					DBObject row = cursor.next();

					try {

						if (colNum % 600000 == 1) {
							sb.append("<Worksheet ss:Name=\"Sheet" + (sheetNum++) + "\">\r\n");
							sb.append("<Table>\r\n");
							sb.append("<Column ss:Index=\"" + sheetNum + "\" ss:AutoFitWidth=\"0\" ss:Width=\"110\"/>\r\n");
							sb.append("<Row>\r\n");
							sb.append("<Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">순번</Data></Cell>\r\n");

							// 해더 넣기
							for (String keyName : keyNameList) {
								sb.append("<Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">" + keyName + "</Data></Cell>\r\n");
							}
							sb.append("</Row>\r\n");
						}
						sb.append("<Row>\r\n");
						sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"Number\">" + colNum + "</Data></Cell>\r\n");
						if (keyList == null) {
							Iterator<String> keyIter = row.keySet().iterator();
							while (keyIter.hasNext()) {
								String key = keyIter.next();

								Object val = row.get(key);

								if (val == null) {
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">-</Data></Cell>\r\n");
								} else if (val instanceof Long || val instanceof Integer || val instanceof Double) {
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"Number\">" + val + "</Data></Cell>\r\n");
									// } else if(val instanceof ObjectId){
									// continue;

								} else {
									String valreplace = (val + "").replace("&", "&amp;").replace(">", "&gt;").replace("<", "&lt;").replace("\"", "&quot;").replace("\'", "&apos;");
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + valreplace + "</Data></Cell>\r\n");
								}
							}
						} else {
							for (String key : keyList) {
								Object val = row.get(key);

								if (val == null) {
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">-</Data></Cell>\r\n");
								} else if (val instanceof Long || val instanceof Integer || val instanceof Double) {
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"Number\">" + val + "</Data></Cell>\r\n");
								} else {
									String valreplace = (val + "").replace("&", "&amp;").replace(">", "&gt;").replace("<", "&lt;").replace("\"", "&quot;").replace("\'", "&apos;");
									sb.append("<Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + valreplace + "</Data></Cell>\r\n");
								}
							}
						}
						sb.append("</Row>\r\n");
						bw.write(sb.toString());

						sb = new StringBuffer();
						// 다음 시트로
						if (colNum % 600000 == 0) {
							sb.append("</Table>\r\n");
							sb.append("</Worksheet>\r\n");
						}

					} catch (Exception e) {
						e.printStackTrace();
					}
				}

				// 마지막 닫기
				if (colNum % 600000 != 1) {
					sb.append("</Table>\r\n");
					sb.append("</Worksheet>\r\n");
				}

				sb.append("</Workbook>\r\n");
				bw.write(sb.toString());
				return tempxls;
			} finally {
				if (bw != null) {
					bw.close();
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
			return null;
		} finally {

			if (bos != null) {
				bos.close();
			}

		}

	}

	public static void download(HttpServletRequest request, HttpServletResponse response, Map<String, Object> data, String path) throws IOException {

		String fileUploadPath = path;
		String fileLogicName = "" + data.get("log_file");
		String filePhysicName = "" + data.get("phy_file");

		BufferedInputStream fin = null;
		try {
			BufferedOutputStream outs = null;
			try {

				String ans_log_file = fileLogicName;
				String ans_phy_file = filePhysicName;

				String upDir = fileUploadPath;
				File theFile = new File(upDir + "/" + ans_phy_file);
//				String strClient = request.getHeader("user-agent");

				if (theFile.exists()) {
					response.setContentType("application/octet-stream;");
					String testtxt = java.net.URLEncoder.encode(ans_log_file, "UTF-8");
					response.setHeader("Content-Disposition", "attachment; filename=\"" + testtxt + "\";");
					response.setContentType("text/html;charset=utf-8");
					response.setCharacterEncoding("UTF-8");
					response.setContentLength((int) theFile.length());
					response.setHeader("Pragma", "no-cache;");
					response.setHeader("Expires", "-1;");

					byte b[] = new byte[1024];
					fin = new BufferedInputStream(new FileInputStream(theFile));
					outs = new BufferedOutputStream(response.getOutputStream());
					int read = 0;
					while ((read = fin.read(b)) != -1) {
						outs.write(b, 0, read);
					}

				} else {
					try {
						response.setContentType("text/html;charset=utf-8");
						response.getWriter().write("<script>alert(\"파일이 존재하지 않습니다.\"); history.go(-2)</script>");
					} catch (IOException ie) {
						ie.getMessage();
					}
				}
			} finally {
				if (outs != null)
					outs.close();
			}
		} finally {
			if (fin != null)
				fin.close();
		}

	}

	public static void download(HttpServletRequest request, HttpServletResponse response, File file, String downloadFileName) throws IOException {

		BufferedInputStream fin = null;
		BufferedOutputStream outs = null;
		try {
			if (file.exists()) {
				response.setContentType("application/octet-stream;");
				String testtxt = java.net.URLEncoder.encode(downloadFileName, "UTF-8");
				response.setHeader("Content-Disposition", "attachment; filename=\"" + testtxt + "\";");
				response.setContentType("text/html;charset=utf-8");
				response.setCharacterEncoding("UTF-8");
				response.setContentLength((int) file.length());
				response.setHeader("Pragma", "no-cache;");
				response.setHeader("Expires", "-1;");

				byte b[] = new byte[1024];
				fin = new BufferedInputStream(new FileInputStream(file));
				outs = new BufferedOutputStream(response.getOutputStream());
				int read = 0;
				while ((read = fin.read(b)) != -1) {
					outs.write(b, 0, read);
				}

			} else {
				try {
					response.setContentType("text/html;charset=utf-8");
					response.getWriter().write("<script>alert(\"파일이 존재하지 않습니다.\"); history.go(-2)</script>");
				} catch (IOException e) {
					e.getMessage();
				}
			}
		} finally {
			if (outs != null)
				outs.close();
			if (fin != null)
				fin.close();
		}
	}

	public static Map<String, Object> uploadFile(Map<String, Object> map, HttpServletRequest request, String path) throws Exception {

		String filePath = path;

		MultipartHttpServletRequest multipartHttpServletRequest = (MultipartHttpServletRequest) request;
		Iterator<String> iterator = multipartHttpServletRequest.getFileNames();

		MultipartFile multipartFile = null;
		String originalFileName = null;
		String originalFileExtension = null;
		String storedFileName = null;

		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		Map<String, Object> listMap = null;

		File file = new File(filePath);
		if (file.exists() == false) {
			file.mkdirs();
		}

		while (iterator.hasNext()) {
			multipartFile = multipartHttpServletRequest.getFile(iterator.next());
			if (multipartFile.isEmpty() == false) {
				originalFileName = multipartFile.getOriginalFilename();
				originalFileExtension = originalFileName.substring(originalFileName.lastIndexOf("."));
				storedFileName = getRandomString() + originalFileExtension;

				file = new File(filePath + storedFileName);
				multipartFile.transferTo(file);

				listMap = new HashMap<String, Object>();
				listMap.put("filePath", filePath);
				listMap.put("orginal_filename", originalFileName);
				listMap.put("stored_filename", storedFileName);
				listMap.put("filesize", fileSize(multipartFile.getSize()));
			}
		}
		return listMap;
	}

	public static String fileSize(long file_size) {
		String[] gubn = { "B", "KB", "MB", "GB" };

		String returnSize = new String();

		int gubnKey = 0;
		long changeSize = 0;

		long fileSize = file_size;

		for (int x = 0; (fileSize / (double) 1024) > 0; x++, fileSize /= (double) 1024) {
			gubnKey = x;
			changeSize = fileSize;
		}

		returnSize = changeSize + gubn[gubnKey];

		return returnSize;
	}

	public static String getRandomString() {
		return UUID.randomUUID().toString().replaceAll("-", "");
	}

	public static Map<String, Object> assetExcelImport(Map<String, Object> data, String path) {

		String collection = "" + data.get("collection");

		Map<String, Object> result = new HashMap<String, Object>();

		HttpServletRequest request = ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getRequest();
		SessionInfo sessionInfo = (SessionInfo) SessionHandler.getInstance().getLoginInfo(request);
		String userId = sessionInfo.getUserId();
		String userNm = sessionInfo.getUserNm();

		XSSFRow row;
		XSSFCell cell;

		int wrk_no;
		int up_wrk_no;

		if (("" + data.get("wrk_no_excel")) == null || ("" + data.get("wrk_no_excel")) == "") {
			wrk_no = 1;
			up_wrk_no = -1;
		} else {
			wrk_no = Integer.parseInt("" + data.get("wrk_no_excel"));
			up_wrk_no = Integer.parseInt("" + data.get("up_wrk_no_excel"));
		}

		int i_cnt = 0;
		BasicDBObject bdo = null;
		try {
			FileInputStream inputStream = new FileInputStream(path);
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			XSSFSheet sheet = null;

			// sheet수 취득
			int sheetCn = workbook.getNumberOfSheets();

			for (int k = 0; k < sheetCn; k++) {

				// 0번째 sheet 정보 취득
				sheet = workbook.getSheetAt(k);

				// 취득된 sheet에서 rows수 취득
				int rows = sheet.getPhysicalNumberOfRows();

				for (int r = 0; r < rows; r++) {
					System.out.println("rows : " + rows);
					bdo = new BasicDBObject();
					row = sheet.getRow(r); // row 가져오기

					// 취득된 row에서 취득대상 cell수 취득
					int cells = row.getPhysicalNumberOfCells();

					if (row != null) {

						String key = "";
						String value = "";
						for (int c = 0; c < cells; c++) {

							cell = row.getCell(c);
							if (cell != null) {

								if (c == 0)
									key = "wrk_ip_s";
								else if (c == 1)
									key = "wrk_ip_e";
								else if (c == 2)
									key = "wrk_nm";
								else if (c == 3)
									key = "wrk_nm_en";
								else if (c == 4)
									key = "etc";

								switch (cell.getCellType()) {
								case XSSFCell.CELL_TYPE_FORMULA:
									value = cell.getCellFormula();
									break;
								case XSSFCell.CELL_TYPE_NUMERIC:
									value = "" + cell.getNumericCellValue();
									break;
								case XSSFCell.CELL_TYPE_STRING:
									value = "" + cell.getStringCellValue();
									break;
								case XSSFCell.CELL_TYPE_BLANK:
									value = "[null 아닌 공백]";
									break;
								case XSSFCell.CELL_TYPE_ERROR:
									value = "" + cell.getErrorCellValue();
									break;
								default:
								}

								Date date = new Date();

								SimpleDateFormat dayTime = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
								String realTime = dayTime.format(date);

								bdo.put("wrk_no", wrk_no);
								bdo.put("up_wrk_no", up_wrk_no);
								bdo.put("reg_dt", realTime);
								bdo.put("reg_id", userId);
								bdo.put("reg_nm", userNm);
								if (key.equals("wrk_ip_s")) {
									bdo.put("wrk_ip_s_long", IPUtil.getStringIpDecIp(value));
								} else if (key.equals("wrk_ip_e")) {
									bdo.put("wrk_ip_e_long", IPUtil.getStringIpDecIp(value));
								}
								bdo.put("key", SequenceUtil.sequence("wrk_ip_info", "key"));
								bdo.put(key, value);

							}

						} // for(c) 문

						BasicDBObject s_bdo = new BasicDBObject();
						s_bdo.put("wrk_ip_s_long", bdo.get("wrk_ip_s_long"));
						s_bdo.put("wrk_ip_e_long", bdo.get("wrk_ip_e_long"));

						Query query = new BasicQuery(s_bdo);
						List<Map> list = mongoTemplate.find(query, Map.class, collection);

						if (list.size() == 0) {
							mongoTemplate.insert(bdo, collection);
							i_cnt++;
						}

					}
				} // for(r) 문
			}

			result.put("ex_cnt", "" + i_cnt);
			result.put("success", true);

		} catch (FileNotFoundException e) {
			LogUtil.error("Error : ", e);
			result.put("success", false);
			result.put("msg", message.getMessage("fail.common.msg", null, Locale.KOREA));
		} catch (IOException e) {
			LogUtil.error("Error : ", e);
			result.put("success", false);
			result.put("msg", message.getMessage("fail.common.msg", null, Locale.KOREA));
		}

		return result;
	}

	public static Map<String, Object> whiteExcelImport(Map<String, Object> data, String path) {

		String collection = "" + data.get("collection");

		Map<String, Object> result = new HashMap<String, Object>();

		HttpServletRequest request = ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getRequest();
		SessionInfo sessionInfo = (SessionInfo) SessionHandler.getInstance().getLoginInfo(request);
		String userId = sessionInfo.getUserId();
		String userNm = sessionInfo.getUserNm();

		XSSFRow row;
		XSSFCell cell;

		int i_cnt = 0;
		BasicDBObject bdo = null;
		try {
			FileInputStream inputStream = new FileInputStream(path);
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			XSSFSheet sheet = null;

			// sheet수 취득
			int sheetCn = workbook.getNumberOfSheets();

			for (int k = 0; k < sheetCn; k++) {

				// 0번째 sheet 정보 취득
				sheet = workbook.getSheetAt(k);

				// 취득된 sheet에서 rows수 취득
				int rows = sheet.getPhysicalNumberOfRows();

				for (int r = 0; r < rows; r++) {
					System.out.println("rows : " + rows);
					bdo = new BasicDBObject();
					row = sheet.getRow(r); // row 가져오기

					// 취득된 row에서 취득대상 cell수 취득
					int cells = row.getPhysicalNumberOfCells();

					if (row != null) {

						String key = "";
						String value = "";
						for (int c = 0; c < cells; c++) {

							cell = row.getCell(c);
							if (cell != null) {

								if (c == 0)
									key = "part";
								else if (c == 1)
									key = "object";
								else if (c == 2)
									key = "expln";

								switch (cell.getCellType()) {
								case XSSFCell.CELL_TYPE_FORMULA:
									value = cell.getCellFormula();
									break;
								case XSSFCell.CELL_TYPE_NUMERIC:
									value = "" + cell.getNumericCellValue();
									break;
								case XSSFCell.CELL_TYPE_STRING:
									value = "" + cell.getStringCellValue();
									break;
								case XSSFCell.CELL_TYPE_BLANK:
									value = "[null 아닌 공백]";
									break;
								case XSSFCell.CELL_TYPE_ERROR:
									value = "" + cell.getErrorCellValue();
									break;
								default:
								}

								Date date = new Date();

								SimpleDateFormat dayTime = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
								String realTime = dayTime.format(date);

								bdo.put("time", realTime);
								bdo.put("user_id", userId);
								bdo.put("user_nm", userNm);

								if (key.equals("part")) {
									if (value.equals("도메인")) {
										value = "url";
									} else if (value.equals("파일해시") || value.equals("해시")) {
										value = "file_hash";
									} else if (value.equals("아이피")) {
										value = "ip";
									}

									bdo.put(key, value);
								} else {
									bdo.put(key, value);
								}

							}

						} // for(c) 문

						BasicDBObject s_bdo = new BasicDBObject();
						s_bdo.put("object", bdo.get("object"));

						Query query = new BasicQuery(s_bdo);
						List<Map> list = mongoTemplate.find(query, Map.class, collection);

						if (list.size() == 0) {
							mongoTemplate.insert(bdo, collection);
							i_cnt++;
						}

					}
				} // for(r) 문
			}

			result.put("ex_cnt", "" + i_cnt);
			result.put("success", true);

		} catch (FileNotFoundException e) {
			LogUtil.error("Error : ", e);
			result.put("success", false);
			result.put("msg", message.getMessage("fail.common.msg", null, Locale.KOREA));
		} catch (IOException e) {
			LogUtil.error("Error : ", e);
			result.put("success", false);
			result.put("msg", message.getMessage("fail.common.msg", null, Locale.KOREA));
		}

		return result;
	}

	public static File getExcel(String collName, List<String> fields, List<String> fieldNames, String userId, DBObject match) {
		DBCollection coll = mongoTemplate.getCollection(collName);
		DBCursor cursor = null;
		if (match == null) {
			cursor = coll.find();
		} else {
			cursor = coll.find(match);
		}
		try {
			File excel = ExcelUtil.makeExcelFile(cursor, fields, fieldNames, userId);
			return excel;
		} catch (Throwable e) {
			e.printStackTrace();
			return null;
		}
	}

}
