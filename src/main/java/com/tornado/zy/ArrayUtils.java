package com.tornado.zy;

/**
 * @author
 * @create 2018-01-04 15:11
 **/

public class ArrayUtils {
    public static boolean isNotEmpty(Object[] tar) {
        return tar != null && tar.length > 0;
    }

    public static boolean isEmpty(Object[] tar) {
        return !isNotEmpty(tar);
    }

}
