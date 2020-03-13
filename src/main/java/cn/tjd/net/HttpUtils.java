package cn.tjd.net;

import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLConnection;
import java.util.List;
import java.util.Map;

/**
 * Http请求工具类，该工具类仅依赖于Jackson（Java的JSON转换工具），该工具类使用JDK自带的类库实现，市面上有很多优秀的第三方库（OkHttp）
 *
 * @Auther: TJD
 * @Date: 2020-01-14
 * @DESCRIPTION:
 **/
public class HttpUtils {

    private final static ObjectMapper mapper = new ObjectMapper();

    private HttpUtils() {
    }

    /**
     * 向指定URL发送GET方法的请求
     *
     * @param url    发送请求的URL
     * @param params 请求参数
     * @return URL 所代表远程资源的响应结果
     */
    public static String sendGet(String url, Map<String, String> params) throws IOException {
        String result = "";
        BufferedReader in = null;
        try {
            String param = generateParamStr(params);
            String urlNameString = url + "?" + param;
            URL realUrl = new URL(urlNameString);
            // 打开和URL之间的连接
            URLConnection connection = realUrl.openConnection();
            // 设置通用的请求属性
            connection.setRequestProperty("accept", "*/*");
            connection.setRequestProperty("connection", "Keep-Alive");
            connection.setRequestProperty("user-agent",
                    "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1;SV1)");
            // 建立实际的连接
            connection.connect();
            // 获取所有响应头字段
            Map<String, List<String>> map = connection.getHeaderFields();
            // 定义 BufferedReader输入流来读取URL的响应
            in = new BufferedReader(new InputStreamReader(
                    connection.getInputStream()));
            String line;
            while ((line = in.readLine()) != null) {
                result += line;
            }
        }
        // 使用finally块来关闭输入流
        finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return result;
    }

    /**
     * 向指定 URL 发送POST方法的请求，请求参数最终通过Body传输
     *
     * @param urlStr 发送请求的 URL
     * @param params 请求参数
     * @return 所代表远程资源的响应结果
     */
    public static String sendPost(String urlStr, Map<String, String> params) throws IOException {
        OutputStreamWriter out = null;
        InputStream is = null;
        try {
            URL url = new URL(urlStr);// 创建连接
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setDoOutput(true);
            connection.setDoInput(true);
            connection.setUseCaches(false);
            connection.setInstanceFollowRedirects(true);
            connection.setRequestMethod("POST"); // 设置请求方式
            connection.setRequestProperty("Accept", "application/json"); // 设置接收数据的格式
            connection.setRequestProperty("Content-Type", "application/json"); // 设置发送数据的格式
            connection.connect();
            out = new OutputStreamWriter(connection.getOutputStream(), "UTF-8"); // utf-8编码
            //将Map类型的实例，转换为JSON格式字符串
            String jsonParams = mapper.writeValueAsString(params);
            out.append(jsonParams);
            out.flush();
            out.close();
            // 读取响应
            is = connection.getInputStream();
            int length = (int) connection.getContentLength();// 获取长度
            String result = null;
            if (length != -1) {
                byte[] data = new byte[length];
                byte[] temp = new byte[512];
                int readLen = 0;
                int destPos = 0;
                while ((readLen = is.read(temp)) > 0) {
                    System.arraycopy(temp, 0, data, destPos, readLen);
                    destPos += readLen;
                }
                result = new String(data, "UTF-8"); // utf-8编码
            }
            return result;
        } finally {
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 根据Map类型的参数生成URL格式参数 aaa=1&bbb=2&ccc=3
     *
     * @param params 如果params为null或者size=0，则返回""（空字符串）
     * @return
     */
    private static String generateParamStr(Map<String, String> params) {
        StringBuilder paramStr = new StringBuilder();
        if (params == null || params.size() == 0) {
            return "";
        }
        for (Map.Entry<String, String> entry : params.entrySet()) {
            paramStr.append(entry.getKey() + "=" + entry.getValue() + "&");
        }
        return paramStr.toString();
    }
}
