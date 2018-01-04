package com.tornado.zy;

/**
 * @author
 * @create 2018-01-04 15:13
 **/

public class StringUtils {

    public static boolean isEmpty(String s) {
        return s == null || s.length() == 0;
    }

    public static void main(String[] args) {
        System.out.println("".length());
    }
}
