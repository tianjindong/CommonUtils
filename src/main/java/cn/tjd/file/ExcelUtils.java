package cn.tjd.file;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;


/**
 * Apache POI操作Excel对象 HSSF：操作Excel 2007之前版本(.xls)格式,生成的EXCEL不经过压缩直接导出
 * XSSF：操作Excel 2007及之后版本(.xlsx)格式,内存占用高于HSSF SXSSF:从POI3.8
 * beta3开始支持,基于XSSF,低内存占用,专门处理大数据量(建议)。
 * <p>
 * 注意: 值得注意的是SXSSFWorkbook只能写(导出)不能读(导入)
 * <p>
 * 说明: .xls格式的excel(最大行数65536行,最大列数256列) .xlsx格式的excel(最大行数1048576行,最大列数16384列)
 *
 * @Auther: TJD
 * @Date: 2019-12-23
 * @DESCRIPTION:
 **/
public class ExcelUtils {

    public static final String DEFAULT_DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";// 默认日期格式
    public static final int DEFAULT_COLUMN_WIDTH = 17;// 默认列宽

    public enum ExcelType {
        XLS(".xls"), XLSX(".xlsx");

        private String suffix;

        ExcelType(String suffix) {
            this.suffix = suffix;
        }

        public String getSuffix() {
            return this.suffix;
        }
    }

    /**
     * 根据Map类型的数据集，导出默认文件名的Excel文件。文件通过获取HttpResponse中输出流，将文件响应给前端<br><br/>
     * headMap的key与数据对象（Map）的key相对应；headMap的value用于指定Excel表头显示的文字<br><br/>
     *
     * @param excelType 用于指定导出的Excel文件格式
     * @param headMap   表头名称
     * @param dataArray 数据集
     * @param response
     * @throws IOException
     */
    public static void exportExcelByMap(ExcelType excelType, LinkedHashMap<String, String> headMap, List<Map<String, Object>> dataArray, HttpServletResponse response) throws IOException {
        exportExcelByMap(excelType, headMap, dataArray, "default", "sheet1", response);
    }

    /**
     * 根据Map类型的数据集，文件通过获取HttpResponse中输出流，将文件响应给前端<br><br/>
     * headMap用于指定表头数据，其中LinkedHashMap的key与dataArray中Map的Key相对应，value为生成的Excel文件的表头：<br><br/>
     * LinkedHashMap<String, String> title = new LinkedHashMap<>();<br/>
     * titles.add(title);<br/>
     * title.put("name", "姓名");<br/>
     * title.put("age", "年龄");<br/>
     * title.put("errtxt", "错误详情");<br/>
     * ExcelUtils.exportExcelByMap(ExcelUtils.ExcelType.XLS, title, data,"xx","xx" response);<br/>
     *
     * @param excelType 导出的文件类型（通过ExcelUtils.ExcelType.XLS或ExcelUtils.ExcelType.XLSX指定文件类型）
     * @param headMap   表头名称
     * @param dataArray 数据集合
     * @param filename  文件名称
     * @param sheetname 工作簿名称
     * @param response  响应对象
     * @throws IOException
     */
    public static void exportExcelByMap(ExcelType excelType, LinkedHashMap<String, String> headMap, List<Map<String, Object>> dataArray, String filename, String sheetname, HttpServletResponse response) throws IOException {
        response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode(filename + excelType.getSuffix(), "UTF-8"));
        createWorkBookByMap(excelType, headMap, dataArray, sheetname, response.getOutputStream());
    }

    /**
     * 根据Java Bean对象集，导出默认文件名的Excel文件。文件通过获取HttpResponse中输出流，将文件响应给前端<br><br/>
     * headMap的key与JavaBean对象的属性名相对应；headMap的value用于指定Excel表头显示的文字<br><br/>
     *
     * @param excelType 用于指定导出的Excel文件格式
     * @param headMap   表头名称
     * @param dataList  数据集
     * @param response
     * @throws IOException
     */
    public static void exportExcelByObject(ExcelType excelType, LinkedHashMap<String, String> headMap, List dataList, HttpServletResponse response) throws IOException {
        exportExcelByObject(excelType, headMap, dataList, "default", "sheet", response);
    }

    /**
     * 通过JavaBean数据集指定的数据，生成对应的Excel文件，并通过获取HttpResponse中的输出流，将文件响应给前端
     *
     * @param excelType 导出的文件类型（通过ExcelUtils.ExcelType.XLS或ExcelUtils.ExcelType.XLSX指定文件类型）
     * @param headMap   表头名称
     * @param dataList  数据集合
     * @param filename  文件名称
     * @param sheetname 工作簿名称
     * @param response  响应对象
     * @throws IOException
     */
    public static void exportExcelByObject(ExcelType excelType, LinkedHashMap<String, String> headMap, List dataList, String filename, String sheetname, HttpServletResponse response) throws IOException {
        response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode(filename + excelType.getSuffix(), "UTF-8"));
        createWorkBookByObject(excelType, headMap, dataList, sheetname, response.getOutputStream());
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
    @Deprecated
    public static String exportExcelByMap(ExcelType excelType, LinkedHashMap<String, String> headMap, List<Map<String, Object>> dataArray, String fileName, String execlPath, HttpServletRequest request) {
        String result = null;

        String path = request.getSession().getServletContext().getRealPath("");
        File file = new File(path + "/download");
        if (!file.exists()) file.mkdirs();// 创建该文件夹目录
        OutputStream os = null;
        try {
            long start = System.currentTimeMillis();
            os = new FileOutputStream(file.getAbsolutePath() + File.separator + start + excelType.getSuffix());
            createWorkBookByMap(excelType, headMap, dataArray, fileName, os);
            result = execlPath + File.separator + start + excelType.getSuffix();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result.replace('\\', '/');
    }


    /**
     * 根据excelType指定的文件格式，生成对应的WorkBoot，并通过输出流导出文件，其中表格数据由Map类型的集合作为载体
     *
     * @param excelType    Excel文件的格式（XLS、XLSX）
     * @param headerMap    用于指定表头信息，其中key对应dataList中的key，value对表表头显示的文字
     * @param dataList     数据集
     * @param sheetName    工作簿名称
     * @param outputStream 输出流
     */
    public static void createWorkBookByMap(ExcelType excelType, LinkedHashMap<String, String> headerMap, List<Map<String, Object>> dataList, String sheetName, OutputStream outputStream) {
        String datePattern = DEFAULT_DATE_PATTERN;
        int minBytes = DEFAULT_COLUMN_WIDTH;
        //声明一个工作簿
        Workbook workbook = generateWorkBook(excelType);
        //生成表头样式
        CellStyle headerStyle = generateHeaderStyle(workbook);
        //生成数据样式
        CellStyle cellStyle = generateCellStyle(workbook);
        //生成一个(带名称)表格
        Sheet sheet = workbook.createSheet(sheetName);
        int[] colWidthArr = new int[headerMap.size()];// 列宽数组
        String[] headKeyArr = new String[headerMap.size()];// headKey数组
        String[] headValArr = new String[headerMap.size()];// headVal数组
        int i = 0;
        for (Map.Entry<String, String> entry : headerMap.entrySet()) {
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
        for (Map<String, Object> obj : dataList) {
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
            workbook.write(outputStream);
            outputStream.flush();// 刷新此输出流并强制将所有缓冲的输出字节写出
            outputStream.close();// 关闭流
            workbook.close();// 释放workbook所占用的所有windows资源
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * 根据excelType指定的文件格式，生成对应的WorkBoot，并通过输出流导出文件，其中表格数据由Java Bean对象作为载体
     *
     * @param excelType    Excel文件的格式（XLS、XLSX）
     * @param headerMap    用于指定表头信息，其中key对应dataList中Java Bean的属性，value对表表头显示的文字
     * @param dataList     Java Bean集合作为数据载体
     * @param sheetName    工作簿名称
     * @param outputStream 输出流
     */
    private static void createWorkBookByObject(ExcelType excelType, LinkedHashMap<String, String> headerMap, List dataList, String sheetName, OutputStream outputStream) {
        String datePattern = DEFAULT_DATE_PATTERN;
        int minBytes = DEFAULT_COLUMN_WIDTH;
        //声明一个工作簿
        Workbook workbook = generateWorkBook(excelType);
        //生成表头样式
        CellStyle headerStyle = generateHeaderStyle(workbook);
        //生成数据样式
        CellStyle cellStyle = generateCellStyle(workbook);
        //生成一个(带名称)表格
        Sheet sheet = workbook.createSheet(sheetName);
        int[] colWidthArr = new int[headerMap.size()];// 列宽数组
        String[] headKeyArr = new String[headerMap.size()];// headKey数组
        String[] headValArr = new String[headerMap.size()];// headVal数组
        int i = 0;
        for (Map.Entry<String, String> entry : headerMap.entrySet()) {
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
        boolean isMapTag = false;
        for (Object obj : dataList) {
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
            Class<?> clazz = obj.getClass();
            for (int k = 0; k < headKeyArr.length; k++) {
                Cell cell = dataRow.createCell(k);// 创建单元格
                String filedName = headKeyArr[k];
                String cellValue = "";
                try {
                    Method method = clazz.getMethod("get" + filedName.substring(0, 1).toUpperCase() + filedName.substring(1));
                    Object result = method.invoke(obj);
                    if (result == null) {
                        cellValue = "";
                    } else if (result instanceof Date) {
                        cellValue = new SimpleDateFormat(datePattern).format(result);
                    } else if (result instanceof Float || result instanceof Double) {
                        cellValue = new BigDecimal(result.toString()).setScale(2,
                                BigDecimal.ROUND_HALF_UP).toString();
                    } else {
                        cellValue = result.toString();
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
                cell.setCellValue(cellValue);
                cell.setCellStyle(cellStyle);
            }
            rowIndex++;
        }
        try {
            workbook.write(outputStream);
            outputStream.flush();// 刷新此输出流并强制将所有缓冲的输出字节写出
            outputStream.close();// 关闭流
            workbook.close();// 释放workbook所占用的所有windows资源
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 根据 ExcelType 动态生成WorkBook（XLS和XLSX两种）
     *
     * @param excelType
     * @return
     */
    private static Workbook generateWorkBook(ExcelType excelType) {
        Workbook workbook;
        if (excelType == ExcelType.XLS) {
            workbook = new HSSFWorkbook();
        } else {
            SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();
            // 大于1000行时会把之前的行写入硬盘
            sxssfWorkbook.setCompressTempFiles(true);
            workbook = sxssfWorkbook;
        }
        return workbook;
    }

    /**
     * 生成Excel表头单元格样式
     *
     * @param workbook
     * @return
     */
    private static CellStyle generateHeaderStyle(Workbook workbook) {
        // head样式
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);// 设置颜色
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);// 前景色纯色填充
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        Font headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        return headerStyle;
    }

    /**
     * 生成Excel数据单元格样式
     *
     * @param workbook
     * @return
     */
    private static CellStyle generateCellStyle(Workbook workbook) {
        // 单元格样式
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        Font cellFont = workbook.createFont();
        cellFont.setBold(false);
        cellStyle.setFont(cellFont);
        return cellStyle;
    }

}
