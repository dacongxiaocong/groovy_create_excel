package com.xjc.createexcel.Util

import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow

import java.lang.reflect.Field
class FillObjectIntoRow<T> {
    //向行中填入数据，数据顺序由list决定
    void fill(T t,HSSFRow row,List<String> class_titles){
        //获取行数据类对象的class对象
        Class<T> tClass = t.class
        //通过反射获取所有全局属性
        Field[] fields = tClass.declaredFields
        int cellNum = 0
        //遍历list，通过list的顺序决定行数据的顺序
        for (class_title in class_titles){
            //遍历类中的属性
            fields.each {field ->
                if (field.name.equals(class_title)){
                    //groovy对抽象的简化，可以获得对象同名属性的值
                    String value = t."$field.name"
                    //获取单元格对象
                    HSSFCell cell = row.createCell(cellNum)
                    //将值填入单元格中
                    cell.setCellValue(value)
                }
            }
            cellNum++
        }
    }
}
