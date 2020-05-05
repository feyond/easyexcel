package com.github.feyond.excel;

import lombok.Getter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * @author
 * @create 2017-09-30 12:31
 **/
public class ImportExcel {

    /**
     * 工作薄对象
     */
    @Getter
    private Workbook wb;

    public ImportExcel(File file) throws IOException {
        if (!file.exists()) {
            throw new RuntimeException("导入文档为空!");
        } else if (!file.isFile() || !file.canRead()) {
            throw new RuntimeException("导入文档不可读!");
        } else if (StringUtils.endsWithIgnoreCase(file.getName(), ExcelConstants.XLS_SUFFIX)) {
            this.wb = new HSSFWorkbook(new FileInputStream(file));
        } else if (StringUtils.endsWithIgnoreCase(file.getName(), ExcelConstants.XLSX_SUFFIX)) {
            this.wb = new XSSFWorkbook(new FileInputStream(file));
        } else {
            throw new RuntimeException("文档格式不正确!");
        }
    }

    public ImportExcel(String filepath) throws IOException {
        this(new File(filepath));
    }

    public <E> List<E> getDataList(int sheetIndex, int dataRow, boolean readMergedCell, Class<E> cls, int... groups) {
        AnnotationSheetWrapper sheetWrapper = new AnnotationSheetWrapper(this, sheetIndex, readMergedCell);
        return sheetWrapper.getDataList(dataRow, cls, groups);
    }

    public List<Map<Integer, Object>> getDataList(int sheetIndex, int dataRow) {
        SheetWrapper sheet = new SheetWrapper(this, sheetIndex);
        return sheet.getDataList(dataRow);
    }

    public List<Map<String, Object>> getDataListWithHeader(int sheetIndex, int headerRow) {
        SheetWrapper sheet = new SheetWrapper(this, sheetIndex);
        return sheet.getDataListWithHeader(headerRow);
    }

    public List<String> getHeaders(int sheetIndex, int headerRowIndex) {
        AnnotationSheetWrapper sheetWrapper = new AnnotationSheetWrapper(this, sheetIndex, false);
        return sheetWrapper.getHeaders(headerRowIndex);
    }
}
