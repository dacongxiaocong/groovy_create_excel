package com.xjc;

import com.xjc.createexcel.CreateWithHeadMapAndList;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class Test {
    public static void main(String[] args) throws IOException {
        CreateWithHeadMapAndList<Car>createMod = new CreateWithHeadMapAndList<>();
        List<Car>list = new ArrayList<>();
        for (int i=0;i<5;i++){
            Car car = new Car();
            car.setId(i);
            car.setName("car-"+i);
            car.setFactory("factory-"+i);
            list.add(car);
        }
        LinkedHashMap<String,String>map = new LinkedHashMap<>();
        map.put("编号","id");
        map.put("车名","name");
        map.put("厂名","factory");
        HSSFWorkbook workbook = createMod.create(map,list);
        FileOutputStream fops = new FileOutputStream("/opt/test.xls");
        workbook.write(fops);
    }
}
