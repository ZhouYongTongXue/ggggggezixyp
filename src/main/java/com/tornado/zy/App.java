package com.tornado.zy;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Hello world!
 *
 */
public class App 
{
    private static String FILE = System.getProperty("user.dir");
    public static void main( String[] args )
    {
         List<Item> data = getData();
        try {
            ExcelExportUtils.c().contentColumns("one", "two", "three").contentData(data)
                    .export(new FileOutputStream(FILE + "/export_data.xls"));
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("OK");
    }

    public static List<Item> getData() {
        List<Item> data = new ArrayList<>();
        try {
            BufferedReader in = new BufferedReader(new FileReader(new File(FILE + "/export_data.txt")));
            String line;
            while ((line = in.readLine()) != null) {
                String onetwo = line.replaceFirst("^\\s+`(\\S+)` (\\S+).*","$1,$2");
                if (onetwo.split(",").length != 2) {
                    continue;
                }
                Item item = new Item();
                item.setOne(onetwo.split(",")[0]);
                item.setTwo(onetwo.split(",")[1].toUpperCase());
                item.setThree(line.contains("NOT NULL") ? "否" : "是");
                data.add(item);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return data;
    }
}
