package cn.tjd.security;

import java.io.UnsupportedEncodingException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

/**
 * MD5签名算法工具类
 *
 * @Auther: TJD
 * @Date: 2020-03-06
 * @DESCRIPTION:
 **/
public class MD5Utils {

    // 十六进制下数字到字符的映射数组
    private final static String[] hexDigits = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d",
            "e", "f"};

    private MD5Utils() {
    }

    /**
     * 使用默认字符集对字符串签名3次，由于默认字符集与平台相关，在不同环境中签名出来的结果可能会不一样，强烈不建议使用
     *
     * @param txt
     * @return
     */
    @Deprecated
    public static String digest3(String txt) {
        try {
            return digest(digest(digest(txt) + txt) + txt);
        } catch (Exception ex) {
            return null;
        }
    }

    /**
     * 对字符串签名三次，使用UTF-8对字符串进行解码
     *
     * @param txt
     * @param encoding 字符编码
     * @return
     * @throws UnsupportedEncodingException 不支持该字符编码
     */
    public static String digest3(String txt, String encoding) throws UnsupportedEncodingException {
        try {
            return digest(digest(digest(txt, encoding) + txt, encoding) + txt, encoding);
        } catch (NoSuchAlgorithmException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 使用默认字符集对字符串进行MD5签名，由于默认字符集与平台相关，在不同环境中签名出来的结果可能会不一样，强烈不建议使用
     *
     * @throws NoSuchAlgorithmException
     */
    @Deprecated
    public static String digest(String originString) throws NoSuchAlgorithmException {
        if (originString != null) {
            // 创建具有指定算法名称的信息摘要
            MessageDigest md = MessageDigest.getInstance("MD5");
            // 使用指定的字节数组对摘要进行最后更新，然后完成摘要计算
            byte[] results = md.digest(originString.getBytes());
            // 将得到的字节数组变成字符串返回
            String resultString = byteArrayToHexString(results);
            return resultString;
        }
        return null;
    }

    /**
     * 对字符串进行MD5签名(签名一次)
     *
     * @throws NoSuchAlgorithmException
     */
    public static String digest(String originString, String charset) throws NoSuchAlgorithmException, UnsupportedEncodingException {
        if (originString != null) {
            // 创建具有指定算法名称的信息摘要
            MessageDigest md = MessageDigest.getInstance("MD5");
            // 使用指定的字节数组对摘要进行最后更新，然后完成摘要计算
            byte[] results = md.digest(originString.getBytes(charset));
            // 将得到的字节数组变成字符串返回
            String resultString = byteArrayToHexString(results);
            return resultString;
        }
        return null;
    }

    /**
     * 转换字节数组为十六进制字符串
     *
     * @param b 字节数组
     * @return 十六进制字符串
     */
    private static String byteArrayToHexString(byte[] b) {
        StringBuffer resultSb = new StringBuffer();
        for (int i = 0; i < b.length; i++) {
            resultSb.append(byteToHexString(b[i]));
        }
        return resultSb.toString();
    }

    /**
     * 判断两个签名字符串是否相等
     *
     * @param md5Str
     * @param md5Str2
     * @return
     */
    public static boolean isEquals(String md5Str, String md5Str2) {
        if (md5Str == md5Str2) {
            return true;
        }
        if (md5Str != null) {
            md5Str = md5Str.replaceAll("-", "").toLowerCase();
            md5Str2 = md5Str2.replaceAll("-", "").toLowerCase();
            if (md5Str.equals(md5Str2)) {
                return true;
            }
            return false;
        }
        return false;
    }

    /**
     * 将一个字节转化成十六进制形式的字符串
     */
    private static String byteToHexString(byte b) {
        int n = b;
        if (n < 0) {
            n = 256 + n;
        }
        int d1 = n / 16;
        int d2 = n % 16;
        return hexDigits[d1] + hexDigits[d2];
    }

}
