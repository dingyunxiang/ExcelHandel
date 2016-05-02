package com.veblen;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

import java.beans.PropertyDescriptor;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Array;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by dingyunxiang on 16/5/2.
 */

//用于接受需要处理的对象
public class ExcelHandle<T> {

    private Class<T> type;

    public ExcelHandle(Class<T> type){
        this.type = type;
    }

    /**
     *
     * @param f 传入的文件
     * @param arr 两列，以及对应的值
     * @param isNum 是否包含列号
     * @return 返回一个list
     *
     */
    public List<T> excelToList(File f,String[][] arr,boolean isNum){
        //定义输出流，读出文件
        List<T> rs = new ArrayList<T>();
        Map<String,String> map = new HashMap<String,String>();
        for(int i=0;i<arr.length;i++){
            map.put(arr[i][1],arr[i][0]);
        }



        FileInputStream input = null;
        HSSFWorkbook wb = null;
        try {
            input = new FileInputStream(f);
            //得到Excel工作簿对象
            wb= new HSSFWorkbook(input);
        } catch (Exception e) {
            e.printStackTrace();
        }

        //得到Excel工作表对象
        HSSFSheet sheet = wb.getSheetAt(0);
        //总行数
        int height = sheet.getLastRowNum();
        //得到Excel工作表的行
        HSSFRow row = sheet.getRow(0);
        //总列数
        int width = row.getLastCellNum();
        //5.得到Excel工作表指定行的单元格
        HSSFCell cell = row.getCell((short)1);
        //6.得到单元格样式
        CellStyle cellStyle = cell.getCellStyle();

        //得到所有的Cell格
        HSSFCell[][] cellArr = new HSSFCell[height+1][width];

        //定义x个T格式的对象
        T[] tArr = (T[]) Array.newInstance(type,height);
        for(int i=0;i<height;i++){
            try {
                tArr[i] = (T) type.newInstance();
            }catch(Exception e){
                e.printStackTrace();
            }
        }

//        T t = new T();
//        //test
//        System.out.println();
        System.out.println("height:"+height);
        System.out.println("tArr[0]:"+tArr[0]);

        //获取到每个cell格
        for(int i=0;i<height+1;i++){
            row = sheet.getRow(i);
            for(int j=0;j<width;j++){
                cell = row.getCell(j);
                if(cell!=null){
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                }
                cellArr[i][j] = cell;
            }
        }

        //所有数据按列处理

        //处理编号列，如果含有编号列，则从第一列开始处理
        int index;
        if(isNum){
            index = 1;
        }else{
            index = 0;
        }

        for(int i=index;i<width;i++){

            String key = cellArr[0][i].getStringCellValue();
           // System.out.println(key);
            String str = map.get(key);
            System.out.println(i+" ___"+str);
            //处理以下列
            for(int j=1;j<=height;j++){
                try {
                    Class clazz = tArr[j - 1].getClass();
                    PropertyDescriptor pd = new PropertyDescriptor(str, clazz);
                    Method getMethod = pd.getWriteMethod();
                    String ls = cellArr[j][i].getStringCellValue();
                    System.out.println(i+":"+j+" "+ls);
                    getMethod.invoke(tArr[j-1],ls);//执行set方法返回一个Object

                }catch (Exception e){
                    e.printStackTrace();
                }
            }
        }

        //将所有值加到List中
        for(int i=0;i<height;i++){
            rs.add(tArr[i]);
        }
        return rs;
    }


    /**
     *
     * @param list 传递进来的对象数组 T中需要携带get和set方法
     * @param arr  传进来一个二维数组，每个一位数组的长度为2，分别对应T中的属性以及在excel中显示的列值 Ex:arr[0][0] = "name" arr[0][1] = "姓名"
     * @param isNum  是否需要带第一列编号列（true-yes false-no）
     * @return 返回临时创建的File格式的文件，创建的文件存在当前目录
     */
    public File listToEcxel(List<T> list,String[][] arr,boolean isNum){

        /*
            完成对Excel的初始工作
         */
        // 创建Excel的工作书册 Workbook,对应到一个excel文档
        HSSFWorkbook wb = new HSSFWorkbook();

        // 创建Excel的工作sheet,对应到一个excel文档的tab
        HSSFSheet sheet = wb.createSheet("sheet1");

        // 设置excel每列宽度
        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 3500);

        // 创建字体样式
        HSSFFont font = wb.createFont();
        font.setFontName("Verdana");
        font.setBoldweight((short) 100);
        font.setFontHeight((short) 300);
        font.setColor(HSSFColor.BLUE.index);

        // 创建单元格样式
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        // 设置边框
        style.setBottomBorderColor(HSSFColor.RED.index);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);

        style.setFont(font);// 设置字体


        //获取list的大小，以及需要显示在list中列的个数，以此构建数组
        int listSize = list.size();
        int arrLength = arr.length;
        HSSFCell[][] cellArr = new HSSFCell[listSize+1][arrLength+1];


        //创建listSize+1行，将数组行与每个cell绑定
        for(int i=0;i<listSize+1;i++){
            HSSFRow row = sheet.createRow(i);
            for(int j=0;j<arrLength+1;j++){
                HSSFCell cell = row.createCell(j);
                cellArr[i][j] = cell;
            }
        }

        //写第一列，编号列
        if(isNum) {
            cellArr[0][0].setCellValue("编号");
            for (int i = 1; i < listSize + 1; i++) {
                cellArr[i][0].setCellValue(i);
            }
        }


        //依次写出list中剩下的数据，每次写出一列
        int index = 1;
        for(int i=0;i<arr.length;i++){
            index  = 1;
            cellArr[0][i+1].setCellValue(arr[i][1]);
            for(T t:list){
                try {
                    //使用反射获取每一列的值
                    Class clazz = t.getClass();
                    PropertyDescriptor pd = new PropertyDescriptor(arr[i][0],clazz);
                    Method getMethod = pd.getReadMethod();//获得get方法
                    Object o = getMethod.invoke(t);//执行get方法返回一个Object
                    cellArr[index++][i+1].setCellValue(o + "");
                }catch(Exception e){
                    e.printStackTrace();
                }
            }
        }

        try (FileOutputStream os = new FileOutputStream("workbook.xls")) {
            wb.write(os);
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        File f = new File("workbook.xls");
        return f;

    }

}
