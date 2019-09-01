package com.zh.oukele.read;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.metadata.Sheet;
import com.zh.oukele.listener.ExcelModelListener;
import com.zh.oukele.model.ExcelMode;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

/**
 * 读取 excel 表格 内容
 */
public class EasyExcelRead {

    public static void main(String[] args) {
        simpleRead1();
    }

    // 简单读取 (同步读取)
    public static void simpleRead() {

        // 读取 excel 表格的路径
        String readPath = "C:\\Users\\oukele\\Desktop\\模拟数据.xlsx";

        try {
            // sheetNo --> 读取哪一个 表单
            // headLineMun --> 从哪一行开始读取( 不包括定义的这一行，比如 headLineMun为2 ，那么取出来的数据是从 第三行的数据开始读取 )
            // clazz --> 将读取的数据，转化成对应的实体，需要 extends BaseRowModel
            Sheet sheet = new Sheet(1, 1, ExcelMode.class);

            // 这里 取出来的是 ExcelModel实体 的集合
            List<Object> readList = EasyExcelFactory.read(new FileInputStream(readPath), sheet);
            // 存 ExcelMode 实体的 集合
            List<ExcelMode> list = new ArrayList<ExcelMode>();
            for (Object obj : readList) {
                list.add((ExcelMode) obj);
            }

            // 取出数据
            StringBuilder str = new StringBuilder();
            str.append("{");
            String link = "";
            for (ExcelMode mode : list) {
                str.append(link).append("\""+mode.getColumn1()+"\":").append("\""+mode.getColumn2()+"\"");
                link= ",";
            }
            str.append("};");
            System.out.println(str);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }


    }

    // 异步读取
    public static void simpleRead1(){

        // 读取 excel 表格的路径
        String readPath = "C:\\Users\\oukele\\Desktop\\模拟数据.xlsx";

        try {
            Sheet sheet = new Sheet(1,1,ExcelMode.class);
            EasyExcelFactory.readBySax(new FileInputStream(readPath),sheet,new ExcelModelListener());

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

    }


}
