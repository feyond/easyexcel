package com.github.feyond;


import com.github.feyond.excel.ExportExcel;
import com.github.feyond.excel.OperateType;
import com.github.feyond.excel.annotation.ExcelField;
import com.github.feyond.excel.annotation.ExcelSheet;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.IOException;
import java.util.*;

/**
 * @author chenfy
 * @create 2017-10-23 21:20
 **/
public class ExportTests {
    public static void main(String[] args) throws IOException {
        List<Map> list = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put("key1", "value1");
        map.put("key2", OperateType.EXPORT);
        map.put("key3", Integer.valueOf(24));
        map.put("key4", Double.valueOf(1.24));
        map.put("key5", new Date());
        map.put("key6", Boolean.TRUE);
        map.put("key7", new String[]{"aa","bb","cc"});
        map.put("key8", Arrays.asList("aa","bb","cc"));
        map.put("key9", Arrays.asList(OperateType.EXPORT, OperateType.IMPORT, OperateType.BOTH));
        list.add(map);
        /*new ExportExcel()
                .sheet("Sheet1", new String[]{"字段1", "字段2","字段3", "字段4","字段5", "字段6","字段7", "字段8", "字段9"})
                .setDataMapList(Arrays.asList("key1", "key2","key3", "key4","key5", "key6","key7", "key8","key9"), list)
                .and().writeFile("X:\\home\\1.xlsx");*/

        List<A> aList = new ArrayList<>();
        A a = new A();
        a.setKey1("value1");
        a.setKey2(OperateType.EXPORT);
        a.setKey3(Integer.valueOf(24));
        a.setKey4(1.24);
        a.setKey5(new Date());
        a.setKey6(Boolean.TRUE);
        a.setKey7(new String[]{"aa","bb","cc"});
        a.setKey8(Arrays.asList("aa","bb","cc"));
        a.setKey9(Arrays.asList(OperateType.EXPORT, OperateType.IMPORT, OperateType.BOTH));
        aList.add(a);
        aList.add(a);
        aList.add(a);
        new ExportExcel()
                .sheet("Sheet2", Arrays.asList("字段1", "字段2","字段3", "字段4","字段5", "字段6","字段7", "字段8", "字段9"))
                .setDataList(Arrays.asList("key1", "key2","key3", "key4","key5", "key6","key7", "key8","key9"), aList)
                .and().sheet("Sheet1", new String[]{"字段1", "字段2","字段3", "字段4","字段5", "字段6","字段7", "字段8", "字段9"})
                .setDataMapList(Arrays.asList("key1", "key2","key3", "key4","key5", "key6","key7", "key8","key9"), list)
                .and()
                .sheet(A.class).setDataList(aList)
                .and()
                .sheet("HELLO SHEET 1", A.class, 1).setDataList(aList)
                .and().sheet("HELLO SHEET 2", A.class, 2).setDataList(aList)
                .and().sheet("HELLO SHEET 3", A.class, 0, 1).setDataList(aList)
                .and().writeFile("X:\\home\\2.xlsx");


    }

    @Data
    @NoArgsConstructor
    @ExcelSheet("HELLO SHEET")
    public static class A {

        @ExcelField(header = "字段1", groups = {1}, sort = 10)
        private String key1;
        @ExcelField(header = "字段2", groups = {1}, sort = 20)
        private OperateType key2;
        @ExcelField(header = "字段3", groups = {2}, sort = 30)
        private Integer key3;
        @ExcelField(header = "字段4", groups = {2}, sort = 40)
        private double key4;
        @ExcelField(header = "字段5", groups = {1, 2}, sort = 50, format = "YYYY-MM-DD")
        private Date key5;
        @ExcelField(header = "字段6", groups = {1, 2}, sort = 60)
        private Boolean key6;
        @ExcelField(header = "字段7", sort = 70)
        private String[] key7;
        @ExcelField(header = "字段8", sort = 80)
        private List<String> key8;
        @ExcelField(header = "字段9", groups = {3}, sort = 100)
        private List<OperateType> key9;
    }
}
