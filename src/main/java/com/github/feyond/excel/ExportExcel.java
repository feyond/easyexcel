package com.github.feyond.excel;

import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.List;

/**
 * 导出Excel文件（导出“XLSX”格式，支持大数据量导出）
 */
@Slf4j
public class ExportExcel {

    /**
     * 工作薄对象
     */
    @Getter
    private SXSSFWorkbook wb;


    public ExportExcel() {
        this.wb = new SXSSFWorkbook(500);
    }

    public AnnotationSheetWrapper sheet(String sheetName, Class<?> cls, int... groups) {
        return new AnnotationSheetWrapper(this, sheetName, cls, groups);
    }

    public AnnotationSheetWrapper sheet(Class<?> cls, int... groups) {
        return new AnnotationSheetWrapper(this, null, cls, groups);
    }

    public SheetWrapper sheet(String sheetName, String[] headers) {
        return new SheetWrapper(this, sheetName, headers);
    }

    public SheetWrapper sheet(String sheetName, List<String> headerList) {
        return new SheetWrapper(this, sheetName, headerList);
    }


    /**
     * 浏览器下载
     * @param response
     * @param filename
     * @return
     * @throws IOException
     */
    public ExportExcel download(HttpServletResponse response, String filename) throws IOException {
        if(response != null) {
            // response对象不为空,响应到浏览器下载
            String name = StringUtils.endsWithIgnoreCase(filename, ExcelConstants.XLSX_SUFFIX) ? filename : filename + ExcelConstants.XLSX_SUFFIX;
            response.setContentType(ExcelConstants.XLSX_CONTENT_TYPE);
            response.setHeader("Content-disposition", "attachment; filename=" + URLEncoder.encode(name, "UTF-8"));
            write(response.getOutputStream());
        }
        return this;
    }

    /**
     * 输出数据流
     *
     * @param os 输出数据流
     */
    public ExportExcel write(OutputStream os) throws IOException {
        wb.write(os);
        return this;
    }

    /**
     * 输出到文件
     *
     * @param fileName 输出文件名
     */
    public ExportExcel writeFile(File file) throws FileNotFoundException, IOException {
        FileOutputStream os = new FileOutputStream(file);
        this.write(os);
        return this;
    }


    /**
     * 输出到文件
     *
     * @param fileName 输出文件名
     */
    public ExportExcel writeFile(String name) throws FileNotFoundException, IOException {
        FileOutputStream os = new FileOutputStream(name);
        this.write(os);
        return this;
    }

    /**
     * 清理临时文件
     */
    public void dispose() {
        wb.dispose();
    }

}
