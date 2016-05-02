package com.veblen;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by dingyunxiang on 16/5/2.
 */
public class Test {
    public static void main(String[] args) {
        ExcelHandle<testClass> handle = new ExcelHandle<testClass>(testClass.class);


        List<testClass> list1 = new ArrayList<testClass>();

        for(int i=0;i<10;i++){
            list1.add(new testClass(i,i));
        }

        String[][] a = {{"name","姓名"},{"age","年龄"}};

        Map<String,String> map = new HashMap<String,String>();
        for(int i=0;i<a.length;i++){
            map.put(a[i][1],a[i][0]);
        }
       // System.out.println(map.get("姓名"));

        File f = new File("workbook.xls");
        System.out.println("file:"+f);
        List<testClass> list = handle.excelToList(f,a,true);

        for(testClass t:list){
            System.out.println(t.getName());
        }
    }
}
