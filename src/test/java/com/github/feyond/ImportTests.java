package com.github.feyond;

import com.github.feyond.excel.ImportExcel;
import com.github.feyond.excel.OperateType;
import com.github.feyond.excel.annotation.ExcelField;
import com.github.feyond.excel.annotation.ExcelSheet;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

import java.io.IOException;
import java.util.Date;
import java.util.List;

/**
 * @author chenfy
 * @create 2017-10-24 14:04
 **/
public class ImportTests {
    public static void main(String[] args) throws IOException {
        System.out.println(new ImportExcel("X:\\home\\2.xlsx")
                .getDataList(0, 2));

        System.out.println(new ImportExcel("X:\\home\\2.xlsx")
                .getDataListWithHeader(2, 1));

        System.out.println(new ImportExcel("X:\\home\\2.xlsx").getDataList(2,2, A.class));
        System.out.println(new ImportExcel("X:\\home\\2.xlsx").getDataList(3,2, A.class, 1));

    }

    @Data
    @NoArgsConstructor
    @ExcelSheet("HELLO SHEET")
    @ToString
    public static class A {

        @ExcelField(header = "字段1", groups = {1}, sort = 10)
        private String key1;
        @ExcelField(header = "字段2", groups = {1}, sort = 20)
        private OperateType key2;
        @ExcelField(header = "字段3", groups = {2}, sort = 30)
        private Integer key3;
        @ExcelField(header = "字段4", groups = {2}, sort = 40)
        private double key4;
        @ExcelField(header = "字段5", groups = {1, 2}, sort = 50)
        private Date key5;
        @ExcelField(header = "字段6", groups = {1, 2}, sort = 60)
        private Boolean key6;
        @ExcelField(header = "字段7", sort = 70)
        private String[] key7;
        @ExcelField(header = "字段8", sort = 80)
        private List<String> key8;
        @ExcelField(header = "字段9", groups = {3}, sort = 90)
        private List<OperateType> key9;
    }
}
