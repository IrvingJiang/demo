package com.example.demo.excel;

import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.lang.NonNull;

import javax.servlet.http.HttpServletResponse;
import java.util.Map;
import java.util.function.Consumer;

/**
 * @Description:
 * @author:Irving
 * @date 2023/1/12 15:21
 */
public class ExcelExportContext {

    @NonNull
    private HttpServletResponse response;
    @NonNull
    private String excelTemplatePath;
    @NonNull
    private TemplateExportParams params;
    @NonNull
    private String targetExcelName;
    /**
     * 对应导出模板提取数据的key
     */
    @NonNull
    private String dataKey;

    /**
     * 拓展参数
     */
    private Map extData;

    /**
     * 后置处理
     */
    private Consumer<Workbook> after;

    @NonNull
    public HttpServletResponse getResponse() {
        return response;
    }

    @NonNull
    public String getExcelTemplatePath() {
        return excelTemplatePath;
    }

    @NonNull
    public TemplateExportParams getParams() {
        return params;
    }

    @NonNull
    public String getTargetExcelName() {
        return targetExcelName;
    }

    @NonNull
    public String getDataKey() {
        return dataKey;
    }

    public Map<String, Object> getExtData() {
        return extData;
    }

    public Consumer<Workbook> getAfter() {
        return after;
    }

    public void setResponse(@NonNull HttpServletResponse response) {
        this.response = response;
    }

    public void setExcelTemplatePath(@NonNull String excelTemplatePath) {
        this.excelTemplatePath = excelTemplatePath;
    }

    public void setParams(@NonNull TemplateExportParams params) {
        this.params = params;
    }

    public void setTargetExcelName(@NonNull String targetExcelName) {
        this.targetExcelName = targetExcelName;
    }

    public void setDataKey(@NonNull String dataKey) {
        this.dataKey = dataKey;
    }

    public void setExtData(Map extData) {
        this.extData = extData;
    }

    public void setAfter(Consumer<Workbook> after) {
        this.after = after;
    }
}
