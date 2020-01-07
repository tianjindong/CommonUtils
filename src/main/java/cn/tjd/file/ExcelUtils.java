package cn.tjd.file;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.*;


/**
 * Apache POI操作Excel对象 HSSF：操作Excel 2007之前版本(.xls)格式,生成的EXCEL不经过压缩直接导出
 *  XSSF：操作Excel 2007及之后版本(.xlsx)格式,内存占用高于HSSF SXSSF:从POI3.8
 *  beta3开始支持,基于XSSF,低内存占用,专门处理大数据量(建议)。
 *  <p>
 *  注意: 值得注意的是SXSSFWorkbook只能写(导出)不能读(导入)
 *  <p>
 *  说明: .xls格式的excel(最大行数65536行,最大列数256列) .xlsx格式的excel(最大行数1048576行,最大列数16384列)
 * @Auther: TJD
 * @Date: 2019-12-23
 * @DESCRIPTION:
 **/
public class ExcelUtils {

    public static final String DEFAULT_DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";// 默认日期格式
    public static final int DEFAULT_COLUMN_WIDTH = 17;// 默认列宽

    public static enum ExcelType {
        XLS, XLSX;
    }

    /**
     * 导出Excel(.xlsx)格式
     * <p>
     * 表格头信息集合
     *
     * @param dataArray 数据数组
     * @param os        文件输出流
     */
    @SuppressWarnings({"rawtypes", "deprecation", "resource"})
    public static void createWorkBook(ExcelType excelType, List<Map<String, Object>> dataArray, String fileName, List<LinkedHashMap> headMap, OutputStream os) {
        String datePattern = DEFAULT_DATE_PATTERN;
        int minBytes = DEFAULT_COLUMN_WIDTH;
        //声明一个工作簿
        Workbook workbook = null;
        if (excelType == ExcelType.XLS) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new SXSSFWorkbook();
        }
        // 大于1000行时会把之前的行写入硬盘
//		workbook.setCompressTempFiles(true);

        // 表头1样式
        CellStyle title1Style = workbook.createCellStyle();
        title1Style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 水平居中
        title1Style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 垂直居中
        Font titleFont = workbook.createFont();// 字体
        titleFont.setFontHeightInPoints((short) 20);
        titleFont.setBoldweight((short) 700);
        title1Style.setFont(titleFont);

        // 表头2样式
        CellStyle title2Style = workbook.createCellStyle();
        title2Style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        title2Style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        title2Style.setBorderTop(HSSFCellStyle.BORDER_THIN);// 上边框
        title2Style.setBorderRight(HSSFCellStyle.BORDER_THIN);// 右
        title2Style.setBorderBottom(HSSFCellStyle.BORDER_THIN);// 下
        title2Style.setBorderLeft(HSSFCellStyle.BORDER_THIN);// 左
        Font title2Font = workbook.createFont();
        title2Font.setUnderline((byte) 1);
        title2Font.setColor(HSSFColor.BLUE.index);
        title2Style.setFont(title2Font);

        // head样式
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        headerStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        headerStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);// 设置颜色
        headerStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);// 前景色纯色填充
        headerStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        Font headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        headerStyle.setFont(headerFont);

        // 单元格样式
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        Font cellFont = workbook.createFont();
        cellFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        cellStyle.setFont(cellFont);


        /**
         * 生成一个(带名称)表格
         */
        Sheet sheet = workbook.createSheet(fileName);
        /**
         * 生成head相关信息+设置每列宽度
         */
        @SuppressWarnings("unchecked")
        LinkedHashMap<String, String> headMaps = headMap.get(0);
        int[] colWidthArr = new int[headMaps.size()];// 列宽数组
        String[] headKeyArr = new String[headMaps.size()];// headKey数组
        String[] headValArr = new String[headMaps.size()];// headVal数组
        int i = 0;
        for (Map.Entry<String, String> entry : headMaps.entrySet()) {
            headKeyArr[i] = entry.getKey();
            headValArr[i] = entry.getValue();

            int bytes = headKeyArr[i].getBytes().length;
            colWidthArr[i] = bytes < minBytes ? minBytes : bytes;
            sheet.setColumnWidth(i, colWidthArr[i] * 256);// 设置列宽
            i++;
        }
        /**
         * 遍历数据集合，产生Excel行数据
         */
        int rowIndex = 0;
        for (Map<String, Object> obj : dataArray) {
            // 生成title+head信息
            if (rowIndex == 0) {
                Row headerRow = sheet.createRow(0);// head行
                for (int j = 0; j < headValArr.length; j++) {
                    headerRow.createCell(j).setCellValue(headValArr[j]);
                    headerRow.getCell(j).setCellStyle(headerStyle);
                }
                rowIndex = 1;
            }
            // 生成数据
            Row dataRow = sheet.createRow(rowIndex);// 创建行
            for (int k = 0; k < headKeyArr.length; k++) {
                Cell cell = dataRow.createCell(k);// 创建单元格
                Object o = obj.get(headKeyArr[k]);
                String cellValue = "";
                if (o == null) {
                    cellValue = "";
                } else if (o instanceof Date) {
                    cellValue = new SimpleDateFormat(datePattern).format(o);
                } else if (o instanceof Float || o instanceof Double) {
                    cellValue = new BigDecimal(o.toString()).setScale(2,
                            BigDecimal.ROUND_HALF_UP).toString();
                } else {
                    cellValue = o.toString();
                }

                cell.setCellValue(cellValue);
                cell.setCellStyle(cellStyle);
            }
            rowIndex++;
        }
        try {
            workbook.write(os);
            os.flush();// 刷新此输出流并强制将所有缓冲的输出字节写出
            os.close();// 关闭流
            workbook.close();// 释放workbook所占用的所有windows资源
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 通过输出流导出Excel文件
	 * <p>headMap用于指定表头数据，其中headMap的LinkedHashMap的key与dataArray中Map的Key相对应，value为生成的Excel文件的表头：</p>
	 * <p>List<LinkedHashMap> titles = new ArrayList<>();</p>
	 * <p>LinkedHashMap<String, String> title = new LinkedHashMap<>();</p>
	 * <p>titles.add(title);</p>
	 * <p>title.put("name", "姓名");</p>
	 * <p>title.put("age", "年龄");</p>
	 * <p>title.put("errtxt", "错误详情");</p>
	 * <p>ExcelUtils.exportExcel(ExcelUtils.ExcelType.XLS, titles, data, "dealErrorDetails", "sheet1", response);</p>
     *
     * @param excelType 导出的文件类型（通过ExcelUtils.ExcelType.XLS或ExcelUtils.ExcelType.XLSX指定文件类型）
	 * @param headMap   表头名称
	 * @param dataArray 数据集合
	 * @param filename  文件名称
	 * @param sheetname 工作簿名称
     * @param response  响应对象
     * @throws IOException
     */
    public static void exportExcel(ExcelType excelType, List<LinkedHashMap> headMap, List<Map<String, Object>> dataArray, String filename, String sheetname, HttpServletResponse response) throws IOException {
        if (excelType == ExcelType.XLS) {
            response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode(filename + ".xls", "UTF-8"));
        } else {
            response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode(filename + ".xlsx", "UTF-8"));
        }
        createWorkBook(excelType, dataArray, sheetname, headMap, response.getOutputStream());
    }

    /**
     * 将文件生成在本地后，然后将下载路径返回给前端
     *
     * @param dataArray 数据集合
     * @param fileName  sheet名称
     * @param headMap   列集合名称
     * @param execlPath 服务器存放地址
     * @param request
     * @return
     */
    @SuppressWarnings("rawtypes")
    public static String exportExcel(ExcelType excelType, List<Map<String, Object>> dataArray, String fileName, ArrayList<LinkedHashMap> headMap, String execlPath, HttpServletRequest request) {
        String result = null;

        String path = request.getSession().getServletContext().getRealPath("");
        File file = new File(path + "/download");
        if (!file.exists()) file.mkdirs();// 创建该文件夹目录
        OutputStream os = null;
        try {
            long start = System.currentTimeMillis();
            // .xlsx格式
            if (excelType == ExcelType.XLS) {
                os = new FileOutputStream(file.getAbsolutePath() + File.separator + start + ".xls");
            } else {
                os = new FileOutputStream(file.getAbsolutePath() + File.separator + start + ".xlsx");
            }
            createWorkBook(excelType, dataArray, fileName, headMap, os);
            if (excelType == ExcelType.XLS) {
                result = execlPath + File.separator + start + ".xls";
            } else {
                result = execlPath + File.separator + start + ".xlsx";
            }
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result.replace('\\', '/');
    }

    /**
     * 判断字符串首字母是否为大写，
     * 是则转小写，不是直接返回
     *
     * @param str
     * @return
     */
    public static String getString(String str) {
        char[] chars = new char[1];
        chars[0] = str.charAt(0);
        String temp = new String(chars);
        if (chars[0] >= 'A' && chars[0] <= 'Z') {//当为字母时，则转换为小写
            temp = str.replaceFirst(temp, temp.toLowerCase());
        } else {
            temp = str;
        }
        return temp;
    }
}
