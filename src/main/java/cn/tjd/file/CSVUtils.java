package cn.tjd.file;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.CSVRecord;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * CSVUtils是用于CSV文件快速导入导出的工具类，依赖于<br/>
 * <pre>
 * &lt;dependency&gt;
 *  &lt;groupId&gt;org.apache.commons&lt;/groupId&gt;
 *  &lt;artifactId&gt;commons-csv&lt;/artifactId&gt;
 *  &lt;version&gt;1.6&lt;/version&gt;
 * &lt;/dependency&gt;
 * <pre/>
 *
 *
 *
 * @Auther: TJD
 * @Date: 2020-03-13
 * @DESCRIPTION:
 **/
public class CSVUtils {

    public static final String DEFAULT_DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";// 默认日期格式

    private CSVUtils() {
    }

    /**
     * --------------------------导出-----------------------------------------------
     */

    /**
     * 根据Map类型的数据集，导出默认文件名的CSV文件。文件通过获取HttpResponse中输出流，将文件响应给前端<br><br/>
     * headMap的key与数据对象（Map）的key相对应；headMap的value用于指定CSV表头显示的文字<br><br/>
     *
     * @param headMap   表头名称
     * @param dataArray 数据集
     * @param encoding  文件的编码格式，由于CSV本质上是一个存文本格式，这里需要设置文件的编码格式
     * @param response
     * @throws IOException
     */
    public static void exportCSVByMap(LinkedHashMap<String, String> headMap, List<Map<String, Object>> dataArray,
                                      String encoding, HttpServletResponse response) throws IOException {
        exportCSVByMap(headMap, dataArray, "default", encoding, response);
    }


    /**
     * 根据Map类型的数据集，文件通过获取HttpResponse中输出流，将文件响应给前端<br><br/>
     * headMap用于指定表头数据，它的key与dataArray中Map的Key相对应，value为生成的CSV文件的表头：<br><br/>
     * LinkedHashMap<String, String> title = new LinkedHashMap<>();<br/>
     * titles.add(title);<br/>
     * title.put("name", "姓名");<br/>
     * title.put("age", "年龄");<br/>
     * title.put("errtxt", "错误详情");<br/>
     * CSVUtils.exportCSVByMap(title, data,"xx","xx" response);<br/>
     *
     * @param headMap   表头名称
     * @param dataArray 数据集合
     * @param filename  文件名称
     * @param encoding  文件的编码格式，由于CSV本质上是一个存文本格式，这里需要设置文件的编码格式
     * @param response  响应对象
     * @throws IOException
     */
    public static void exportCSVByMap(LinkedHashMap<String, String> headMap, List<Map<String, Object>> dataArray,
                                      String filename, String encoding, HttpServletResponse response) throws IOException {
        response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode(filename + ".csv", "UTF-8"));
        generateCsv(headMap, dataArray, encoding, response.getOutputStream());
    }


    /**
     * 根据Java Bean对象集，导出默认文件名的CSV文件。文件通过获取HttpResponse中输出流，将文件响应给前端<br><br/>
     * headMap的key与JavaBean对象的属性名相对应；headMap的value用于指定CSV表头显示的文字<br><br/>
     *
     * @param headMap   表头名称
     * @param dataArray 数据集
     * @param encoding  字符编码，CSV文件本质上是文本文件，需要指定字符编码
     * @param response
     * @throws IOException
     */
    public static void exportCSVByObject(LinkedHashMap<String, String> headMap, List<Map<String, Object>> dataArray, String encoding, HttpServletResponse response) throws IOException {
        response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode("defualt" + ".csv", "UTF-8"));
        generateCsv(headMap, dataArray, encoding, response.getOutputStream());
    }

    /**
     * 通过JavaBean数据集指定的数据，生成对应的CSV文件，并通过获取HttpResponse中的输出流，将文件响应给前端
     *
     * @param headMap   表头名称
     * @param dataArray 数据集合
     * @param filename  文件名称
     * @param encoding  文件编码，CSV文件本质上是文本文件，需要指定字符编码
     * @param response  响应对象
     * @throws IOException
     */
    public static void exportCSVByObject(LinkedHashMap<String, String> headMap, List<Map<String, Object>> dataArray,
                                         String filename, String encoding, HttpServletResponse response) throws IOException {
        response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode(filename + ".csv", "UTF-8"));
        generateCsv(headMap, dataArray, encoding, response.getOutputStream());
    }

    /**
     * 通过标准输出流，导出CSV文件
     *
     * @param headMap      表头约束；headMap的key与数据对象（Map）的key相对应；headMap的value用于指定CSV表头显示的文字
     * @param dataArray    数据行,数据行具体的类型可以是Map也可以是JavaBean
     * @param encoding     文件的编码格式，由于CSV本质上是一个存文本格式，这里需要设置文件的编码格式
     * @param outputStream 输出流
     * @throws IOException
     */
    public static void generateCsv(LinkedHashMap<String, String> headMap, List dataArray,
                                   String encoding, OutputStream outputStream) throws IOException {
        OutputStreamWriter osw = new OutputStreamWriter(outputStream, encoding);
        CSVFormat csvFormat = null;
        CSVPrinter csvPrinter = null;

        if (headMap != null) {
            Set<String> keySet = headMap.keySet();
            //生成文件头
            String[] titles = keySet.toArray(new String[keySet.size()]);
            csvFormat = CSVFormat.DEFAULT.withHeader(titles);
            csvPrinter = new CSVPrinter(osw, csvFormat);
            //生成数据行
            for (Object row : dataArray) {
                csvPrinter.printRecord(generateRowArray(headMap, row));
            }
        }
        if (csvPrinter != null) {
            try {
                csvPrinter.flush();
                csvPrinter.close();
            } catch (IOException e) {
                csvPrinter.flush();
                csvPrinter.close();
            }
        }
    }

    /**
     * 按照表头指定的顺序，生成数据行
     *
     * @param headMap
     * @param dataRow dataRow为数据行对象，可能是一个Map对象，也可能是一个JavaBean
     * @return
     */
    private static Object[] generateRowArray(LinkedHashMap<String, String> headMap, Object dataRow) {
        Object[] rowObjects = new Object[headMap.size()];
        int i = 0;
        if (dataRow instanceof Map) {
            for (Map.Entry<String, String> entry : headMap.entrySet()) {
                Object o = ((Map) dataRow).get(entry.getKey());
                rowObjects[i] = objToString(o);
                i++;
            }
        } else {//如果是JavaBean
            Class<?> clazz = dataRow.getClass();
            for (Map.Entry<String, String> entry : headMap.entrySet()) {
                String filedName = entry.getKey();
                try {
                    Method method = clazz.getMethod("get" + filedName.substring(0, 1).toUpperCase() + filedName.substring(1));
                    Object cell = method.invoke(dataRow);
                    rowObjects[i] = objToString(cell);
                } catch (Exception e) {
                    rowObjects[i] = "";
                }
                i++;
            }

        }
        return rowObjects;
    }

    private static String objToString(Object obj) {
        String result = null;
        if (obj == null) {
            result = "";
        } else if (obj instanceof Date) {
            result = new SimpleDateFormat(DEFAULT_DATE_PATTERN).format(obj);
        } else if (obj instanceof Float || obj instanceof Double) {
            result = new BigDecimal(obj.toString()).setScale(2, BigDecimal.ROUND_HALF_UP).toString();
        } else {
            result = obj.toString();
        }
        return result;
    }

    /**
     * --------------------------导入-----------------------------------------------
     */

    public static List<Map<String, String>> readAll(InputStream inputStream, String encoding) throws IOException {
        Reader reader = new BufferedReader(new InputStreamReader(inputStream, encoding));
        CSVParser parser = CSVFormat.EXCEL.withHeader().parse(reader);
        List<CSVRecord> records = parser.getRecords();
        //存在一行数据
        if (records.size() >= 1) {
            List<Map<String, String>> result = new ArrayList<>(records.size());
            Map<String, Integer> headerMap = parser.getHeaderMap();
            for (int i = 1; i < records.size(); i++) {
                Map<String, String> row = convertToMap(headerMap, records.get(i));
                result.add(row);
            }
            return result;
        }
        return Collections.emptyList();
    }

    private static Map<String, String> convertToMap(Map<String, Integer> title, CSVRecord row) {
        if (title.size() > 0) {
            Map<String, String> result = new LinkedHashMap<>(title.size());//指定大小，避免扩容造性能损耗
            for (Map.Entry<String, Integer> titleEntry : title.entrySet()) {
                Integer index = titleEntry.getValue();
                String value = row.size() > index ? row.get(index) : "";
                result.put(titleEntry.getKey(), value);
            }
            return result;
        }
        return Collections.emptyMap();
    }
}
