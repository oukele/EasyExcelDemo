package com.zh.oukele.write;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.zh.oukele.model.ExcelMode;
import org.apache.poi.ss.formula.functions.T;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * 内容 写入 excel 表格
 */
public class EasyExcelWrite {

    public static void main(String[] args) {
        simpleWrite();
    }

    public static void simpleWrite(){

        // 文件输出位置
        String outPath = "C:\\Users\\oukele\\Desktop\\test.xlsx";

        try {
            // 所有行的集合
            List<List<Object>> list = new ArrayList<List<Object>>();

            for (int i = 1; i <= 10; i++) {
                // 第 n 行的数据
                List<Object> row = new ArrayList<Object>();
                row.add("第" + i + "单元格");
                row.add("第" + i + "单元格");
                list.add(row);
            }

            ExcelWriter excelWriter = EasyExcelFactory.getWriter(new FileOutputStream(outPath));
            // 表单
            Sheet sheet = new Sheet(1,0);
            sheet.setSheetName("第一个Sheet");
            // 创建一个表格
            Table table = new Table(1);
            // 动态添加 表头 headList --> 所有表头行集合
            List<List<String>> headList = new ArrayList<List<String>>();
            // 第 n 行 的表头
            List<String> headTitle0 = new ArrayList<String>();
            List<String> headTitle1 = new ArrayList<String>();
            List<String> headTitle2 = new ArrayList<String>();
            headTitle0.add("最顶部-1");
            headTitle0.add("标题1");
            headTitle1.add("最顶部-1");
            headTitle1.add("标题2");
            headTitle2.add("最顶部-1");
            headTitle2.add("标题3");

            headList.add(headTitle0);
            headList.add(headTitle1);
            headList.add(headTitle2);
            table.setHead(headList);

            excelWriter.write1(list,sheet,table);
            // 记得 释放资源
            excelWriter.finish();
            System.out.println("ok");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

    }


}
