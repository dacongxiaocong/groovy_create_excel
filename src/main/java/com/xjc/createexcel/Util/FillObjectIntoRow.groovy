package com.xjc.createexcel.Util

import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow

import java.lang.reflect.Field

class FillObjectIntoRow<T> {
    void fill(T t,HSSFRow row,List<String> class_titles){
        Class<T> tClass = t.class
        Field[] fields = tClass.declaredFields
        int cellNum = 0
        for (class_title in class_titles){
            fields.each {field ->
                if (field.name.equals(class_title)){
                    String value = t."$field.name"
                    HSSFCell cell = row.createCell(cellNum)
                    cell.setCellValue(value)
                }
            }
            cellNum++
        }
    }
}
