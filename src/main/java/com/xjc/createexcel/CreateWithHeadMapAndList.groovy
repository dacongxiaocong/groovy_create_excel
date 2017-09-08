package com.xjc.createexcel

import com.xjc.createexcel.Util.FillObjectIntoRow
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook

class CreateWithHeadMapAndList<T> {
    HSSFWorkbook create(LinkedHashMap<String,String> heads,List<T> objects,String sheetName = "default"){
        HSSFWorkbook workbook = new HSSFWorkbook()
        HSSFSheet sheet = workbook.createSheet(sheetName)
        HSSFRow headRow = sheet.createRow(0)
        FillObjectIntoRow<T> fillObject = new FillObjectIntoRow<>()
        int headCellNum = 0
        int bodyRowNum = 1
        List<String> class_titles= new ArrayList<>()

        for (head in heads){
            String title = head.key
            String class_title = head.value
            HSSFCell cell = headRow.createCell(headCellNum)
            cell.setCellValue(title)
            class_titles<<class_title
            headCellNum++
        }
        for (object in objects){
            HSSFRow bodyRow = sheet.createRow(bodyRowNum)
            fillObject.fill(object,bodyRow,class_titles)
            bodyRowNum++
        }
        return workbook
    }
}
