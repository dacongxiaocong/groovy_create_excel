package com.xjc.createexcel

import com.xjc.createexcel.Util.FillObjectIntoRow
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook

class CreateWithHeadMapAndList<T> {
    //groovy支持在方法定义时设置默认属性值，LinkedHashMap可以保证放入数据的顺序，通过map设置标题与类中全局变量的对应关系和数据列顺序，List为需要插入的行对象，T为行对象类
    HSSFWorkbook create(LinkedHashMap<String,String> heads,List<T> objects,String sheetName = "default"){
        //获取workbook对象
        HSSFWorkbook workbook = new HSSFWorkbook()
        //获取sheet对象
        HSSFSheet sheet = workbook.createSheet(sheetName)
        //获取标题行对象
        HSSFRow headRow = sheet.createRow(0)
        //获取行对象插入工具类
        FillObjectIntoRow<T> fillObject = new FillObjectIntoRow<>()
        int headCellNum = 0
        int bodyRowNum = 1
        //通过map获取这个list，这个list的作用是保证列数据顺序是用户在map中设置的顺序
        List<String> class_titles= new ArrayList<>()

        //插入标题
        for (head in heads){
            String title = head.key
            String class_title = head.value
            HSSFCell cell = headRow.createCell(headCellNum)
            cell.setCellValue(title)
            class_titles<<class_title
            headCellNum++
        }
        //填入行数据
        for (object in objects){
            HSSFRow bodyRow = sheet.createRow(bodyRowNum)
            fillObject.fill(object,bodyRow,class_titles)
            bodyRowNum++
        }
        return workbook
    }
}
