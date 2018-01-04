package com.tornado.zy;

import java.util.List;

/**
 * @author
 * @create 2018-01-04 15:12
 **/

public class CollectionUtils {
    public static boolean isNotEmpty(List s) {
        return s != null && s.size() > 0;
    }

    public static boolean isEmpty(List s) {
        return !isNotEmpty(s);
    }
}
