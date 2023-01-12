package com.example.demo.excel.kit;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import com.alibaba.fastjson.JSON;
import com.example.demo.excel.ExcelExportContext;
import com.example.demo.excel.ExcelImportContext;
import com.google.common.collect.Maps;
import org.apache.commons.collections4.MapUtils;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.List;
import java.util.Map;

/**
 * @Description:excel导入导出工具类
 * @author:Irving
 */
public class ExcelKit {

    public static <T> List<T> importExcel(ExcelImportContext importContext, Class<T> clz) throws Exception {
        return ExcelImportUtil.importExcel(importContext.getFile().getInputStream(), clz, importContext.getImportParams());
    }

    public static <T> void exportExcel(List<T> lefts, ExcelExportContext excelExportContext) throws Exception {

        TemplateExportParams params = new TemplateExportParams(excelExportContext.getExcelTemplatePath());

        Map<String, Object> temp = Maps.newHashMap();
        temp.put(excelExportContext.getDataKey(), lefts);

        Map data = JSON.parseObject((JSON.toJSONString(temp)), Map.class);

        //非表格数据填充
        if (MapUtils.isNotEmpty(excelExportContext.getExtData())) {
            data.putAll(excelExportContext.getExtData());
        }

        Workbook workbook = ExcelExportUtil.exportExcel(params, data);
        if (workbook == null) {
            //读取模板失败
            return;
        }

        //后置不同业务场景对excel的定制化逻辑处理
        if (excelExportContext.getAfter() != null) {
            excelExportContext.getAfter().accept(workbook);
        }

        outputStream(excelExportContext.getTargetExcelName(), excelExportContext.getResponse(), workbook);
    }

    public static String getCellValue(MultipartFile file,Integer rowNum,Integer cellNum) throws IOException {
        //创建一个工作铺
        Workbook workbook = WorkbookFactory.create(file.getInputStream());
        //获取第一个excel表
        Sheet sheet = workbook.getSheetAt(0);

        Row row=sheet.getRow(rowNum);
        Cell cell=row.getCell(cellNum);
        return cell.getStringCellValue();
    }



    private static void outputStream(String fileName, HttpServletResponse response, Workbook workbook) {
        // 重置响应对象
        response.reset();
        try {
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Cache-Control", "no-cache");
        response.setDateHeader("Expires", 0);
        try {
            OutputStream output = response.getOutputStream();
            BufferedOutputStream bufferedOutPut = new BufferedOutputStream(output);
            workbook.write(bufferedOutPut);
            bufferedOutPut.flush();
            bufferedOutPut.close();
            output.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
