package cn.tjd.string;

/**
 * @Auther: TJD
 * @Date: 2020-01-05
 * @DESCRIPTION:
 **/
public class StringUtils {

    private StringUtils() {
    }

    /**
     * 由于数据库中varchar数据类型，一个中文字符占两个字节，也就是说varchar(100)，只能插入50个中文字符，
     * 该方法用于将字符串动态裁切成指定字节长度的字符串
     *
     * @param str
     * @param byteLength
     * @return
     */
    public static String trimStringToByteLength(String str, int byteLength) {
        char[] chars = str.toCharArray();
        int byteNum = 0;
        for (int i = 0; i < chars.length; i++) {
            char c = chars[i];
            //如果是中文字符字节数“+2”，不是则“+1”
            if (isChinese(c)) {
                byteNum += 2;
            } else {
                byteNum++;
            }
            if (byteNum > byteLength) {
                return new String(chars, 0, i);
            }
        }
        return str;
    }

    /**
     * 判断指定字符是否是一个中文字符
     *
     * @param c
     * @return
     */
    public static boolean isChinese(char c) {
        Character.UnicodeBlock ub = Character.UnicodeBlock.of(c);
        if (ub == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS
                || ub == Character.UnicodeBlock.CJK_COMPATIBILITY_IDEOGRAPHS
                || ub == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_A
                || ub == Character.UnicodeBlock.GENERAL_PUNCTUATION
                || ub == Character.UnicodeBlock.CJK_SYMBOLS_AND_PUNCTUATION
                || ub == Character.UnicodeBlock.HALFWIDTH_AND_FULLWIDTH_FORMS) {
            return true;
        }
        return false;
    }

    /**
     * 判断字符串是否为null或者是空字符串，以及空格组成的字符串
     * @param str
     * @return
     */
    public static boolean isBlank(String str) {
        if (str == null || str.replaceAll(" ", "").length() == 0) {
            return true;
        }
        return false;
    }
}
