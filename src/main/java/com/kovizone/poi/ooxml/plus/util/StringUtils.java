package com.kovizone.poi.ooxml.plus.util;

/**
 * <p>字符串通用工具类</p>
 *
 * @author KoviCHen
 * @version 1.0
 */
public class StringUtils {

    /**
     * <p>是否为空(空字符串或null)</p>
     *
     * <p><pre class="code">
     * StringUtils.isEmpty("") = true;
     * StringUtils.isEmpty("  ") = true;
     * StringUtils.isEmpty(null) = true;
     * StringUtils.isEmpty("null") = false;
     * StringUtils.isEmpty("value") = false;
     * </pre></p>
     *
     * @param arg 字符串
     * @return {@code true}则表示字符串非{@code null}且非空字符串
     */
    public static boolean isEmpty(String arg) {
        return arg == null || "".equals(arg.trim());
    }

    /**
     * <p>首字母大写</p>
     *
     * <p><pre class="code">
     *  StringUtils.upperFirstCase("abc") = "Abc";
     *  StringUtils.upperFirstCase("Abc") = "Abc";
     *  StringUtils.upperFirstCase(null) = null;
     *  StringUtils.upperFirstCase("") = "";
     *  StringUtils.upperFirstCase("#") = "#";
     *  StringUtils.lowerFirstCase("#abc") = "#abc";
     * </pre></p>
     *
     * @param arg 字符串
     * @return {@code arg}的首字母大写{@code String}
     */
    public static String upperFirstCase(String arg) {
        if (isEmpty(arg)) {
            return arg;
        }
        char[] cs = arg.toCharArray();
        final char lowerCaseLetterA = 'a';
        final char lowerCaseLetterZ = 'z';
        if (cs[0] >= lowerCaseLetterA && cs[0] <= lowerCaseLetterZ) {
            cs[0] -= 32;
        }
        return String.valueOf(cs);
    }

    /**
     * <p>首字母小写</p>
     *
     * <p><pre class="code">
     *  StringUtils.lowerFirstCase("abc") = "abc";
     *  StringUtils.lowerFirstCase("Abc") = "abc";
     *  StringUtils.lowerFirstCase(null) = null;
     *  StringUtils.lowerFirstCase("") = "";
     *  StringUtils.lowerFirstCase("#") = "#";
     *  StringUtils.lowerFirstCase("#Abc") = "#Abc";
     * </pre></p>
     *
     * @param arg 字符串
     * @return {@code arg}的首字母小写{@code String}
     */
    public static String lowerFirstCase(String arg) {
        if (isEmpty(arg)) {
            return arg;
        }
        char[] cs = arg.toCharArray();
        final char upperCaseLetterA = 'A';
        final char upperCaseLetterZ = 'Z';
        if (cs[0] >= upperCaseLetterA && cs[0] <= upperCaseLetterZ) {
            cs[0] += 32;
        }
        return String.valueOf(cs);
    }

}
