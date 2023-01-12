package com.example.demo.controller;

import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import com.alibaba.fastjson.JSON;
import com.example.demo.excel.ExcelExportContext;
import com.example.demo.excel.ExcelImportContext;
import com.example.demo.excel.ImportHandler;
import com.example.demo.excel.entity.Left;
import com.example.demo.excel.kit.ExcelKit;
import com.google.common.collect.Maps;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.springframework.util.CollectionUtils;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * @Description:
 * @author:Irving
 */
@RequestMapping("/excel")
@RestController
public class ExcelController {


    @PostMapping("/import")
    public List<Left> excelImport(MultipartFile multipartFile) throws Exception {
        //如果要获取月份
        String value = ExcelKit.getCellValue(multipartFile, 1, 4);
        String month = null;
        if (StringUtils.isNotEmpty(value)) {
            month = value.replaceAll("月份：","");
        }
        List<Left> leftList = ExcelKit.importExcel(initExcelImportContext(multipartFile), Left.class);
        //填充月份
        if (!CollectionUtils.isEmpty(leftList)) {
            String finalMonth = month;
            leftList.forEach(e -> e.setYf(finalMonth));
        }
        return leftList;
    }


    @PostMapping("/export")
    public void excelExport(HttpServletResponse response) throws Exception {
        List<Left> lefts = JSON.parseArray(getMockData(), Left.class);
        ExcelKit.exportExcel(lefts, initExcelExportContext(response));
    }


    private ExcelExportContext initExcelExportContext(HttpServletResponse response) throws IOException {

        ExcelExportContext excelExportContext = new ExcelExportContext();

        //模板路径
        excelExportContext.setExcelTemplatePath(getTemplatePath());
        //导出的excel名称
        excelExportContext.setTargetExcelName("表名自行定义.xlsx");
        //easypoi导出相关调参
        TemplateExportParams templateExportParams = new TemplateExportParams();
        templateExportParams.setHeadingStartRow(4);
        excelExportContext.setParams(templateExportParams);
        excelExportContext.setResponse(response);
        excelExportContext.setDataKey("leftlist");

        //设置需要额外填充的数据
        Map ext = Maps.newHashMap();
        ext.put("title", "这是标题");
        ext.put("month", "这是月份");
        excelExportContext.setExtData(ext);

        //设置定制化处理逻辑
        excelExportContext.setAfter(book -> {
            Sheet firstSheet = book.getSheetAt(0);
            List<CellRangeAddress> mergedRegions = firstSheet.getMergedRegions();
            for (CellRangeAddress mergedRegion : mergedRegions) {
                RegionUtil.setBorderBottom(BorderStyle.THIN, mergedRegion, firstSheet);
                RegionUtil.setBorderTop(BorderStyle.THIN, mergedRegion, firstSheet);
                RegionUtil.setBorderLeft(BorderStyle.THIN, mergedRegion, firstSheet);
                RegionUtil.setBorderRight(BorderStyle.THIN, mergedRegion, firstSheet);
            }
        });
        return excelExportContext;
    }


    private ExcelImportContext initExcelImportContext(MultipartFile multipartFile) {
        ExcelImportContext excelImportContext = new ExcelImportContext();
        ImportParams importParams = new ImportParams();
        //参数excel模板不同也少许差异
        //设置标题行
        importParams.setTitleRows(1);
        //表头设置
        importParams.setHeadRows(3);
        //设置读取的sheet
        importParams.setSheetNum(1);
        importParams.setDataHandler(new ImportHandler());
        excelImportContext.setImportParams(importParams);
        excelImportContext.setFile(multipartFile);
        return excelImportContext;
    }

    private String getTemplatePath() throws IOException {
        File directory = new File("src/main/resources");
        String reportPath = directory.getCanonicalPath();
        String resource = reportPath + "/templates/exportTemplate.xlsx";
        return resource;
    }


    private String getMockData() {
        return "[\n" +
                "    {\n" +
                "        \"index\": 1,\n" +
                "        \"bm\": \"无敌\",\n" +
                "        \"xmmc\": \"地方\",\n" +
                "        \"xmfzr\": \"的地方是\",\n" +
                "        \"xmnd\": 2023,\n" +
                "        \"zjly\": \"财政资金\",\n" +
                "        \"xmys\": 12321,\n" +
                "        \"yszc\": 12321,\n" +
                "        \"yszxl\": \"0.99\",\n" +
                "        \"xmrwwcqk\": \"231\",\n" +
                "        \"xmrwwcl\": \"0.88\",\n" +
                "        \"czwt\": \"s'd'fa's\",\n" +
                "        \"xygzjh\": \"321321\",\n" +
                "        \"rights\": [\n" +
                "            {\n" +
                "                \"sfnrndwxcgjh\": \"纳入年度外协采购计划\",\n" +
                "                \"wxcgjhmc\": \"321\",\n" +
                "                \"wxjfys\": 11,\n" +
                "                \"wxjfzc\": 77,\n" +
                "                \"wxrwjdap\": \"0.88\",\n" +
                "                \"wxrwwcqk\": \"的额\",\n" +
                "                \"bz\": \"3\"\n" +
                "            },\n" +
                "            {\n" +
                "                \"sfnrndwxcgjh\": \"单报单批\",\n" +
                "                \"wxcgjhmc\": \"2313\",\n" +
                "                \"wxjfys\": 11,\n" +
                "                \"wxjfzc\": 88,\n" +
                "                \"wxrwjdap\": \"0.11\",\n" +
                "                \"wxrwwcqk\": \"多发点\",\n" +
                "                \"bz\": \"4\"\n" +
                "            },\n" +
                "            {\n" +
                "                \"sfnrndwxcgjh\": \"单报单批\",\n" +
                "                \"wxcgjhmc\": \"1231\",\n" +
                "                \"wxjfys\": 21,\n" +
                "                \"wxjfzc\": 99,\n" +
                "                \"wxrwjdap\": \"0.22\",\n" +
                "                \"wxrwwcqk\": \"打发打发\",\n" +
                "                \"bz\": \"5\"\n" +
                "            },\n" +
                "            {\n" +
                "                \"sfnrndwxcgjh\": \"单报单批\",\n" +
                "                \"wxcgjhmc\": \"2321\",\n" +
                "                \"wxjfys\": 22,\n" +
                "                \"wxjfzc\": 101,\n" +
                "                \"wxrwjdap\": \"0.33\",\n" +
                "                \"wxrwwcqk\": \"d'fa\",\n" +
                "                \"bz\": \"t\"\n" +
                "            },\n" +
                "            {\n" +
                "                \"sfnrndwxcgjh\": \"单报单批\",\n" +
                "                \"wxcgjhmc\": \"21321\",\n" +
                "                \"wxjfys\": 33,\n" +
                "                \"wxjfzc\": 102,\n" +
                "                \"wxrwjdap\": \"0.44\",\n" +
                "                \"wxrwwcqk\": \"a'f'd'f's\",\n" +
                "                \"bz\": \"1\"\n" +
                "            }\n" +
                "        ]\n" +
                "    },\n" +
                "    {\n" +
                "        \"index\": 2,\n" +
                "        \"bm\": \"地方\",\n" +
                "        \"xmmc\": \"十分士大夫\",\n" +
                "        \"xmfzr\": \"士大夫\",\n" +
                "        \"xmnd\": 2023,\n" +
                "        \"zjly\": \"财政资金\",\n" +
                "        \"xmys\": 3231,\n" +
                "        \"yszc\": 312321,\n" +
                "        \"yszxl\": \"0.13\",\n" +
                "        \"xmrwwcqk\": \"23213\",\n" +
                "        \"xmrwwcl\": \"0.88\",\n" +
                "        \"czwt\": \"212\",\n" +
                "        \"xygzjh\": \"12312\",\n" +
                "        \"rights\": [\n" +
                "            {\n" +
                "                \"sfnrndwxcgjh\": \"单报单批\",\n" +
                "                \"wxcgjhmc\": \"21321\",\n" +
                "                \"wxjfys\": 33,\n" +
                "                \"wxjfzc\": 103,\n" +
                "                \"wxrwjdap\": \"0.42\",\n" +
                "                \"wxrwwcqk\": \"的发射点\",\n" +
                "                \"bz\": \"3\"\n" +
                "            },\n" +
                "            {\n" +
                "                \"sfnrndwxcgjh\": \"单报单批\",\n" +
                "                \"wxcgjhmc\": \"12312\",\n" +
                "                \"wxjfys\": 33,\n" +
                "                \"wxjfzc\": 104,\n" +
                "                \"wxrwjdap\": \"0.53\",\n" +
                "                \"wxrwwcqk\": \"打发打发\",\n" +
                "                \"bz\": \"5\"\n" +
                "            },\n" +
                "            {\n" +
                "                \"sfnrndwxcgjh\": \"单报单批\",\n" +
                "                \"wxcgjhmc\": \"12312\",\n" +
                "                \"wxjfys\": 66,\n" +
                "                \"wxjfzc\": 105,\n" +
                "                \"wxrwjdap\": \"0.34\",\n" +
                "                \"wxrwwcqk\": \"啊发的发\",\n" +
                "                \"bz\": \"9\"\n" +
                "            }\n" +
                "        ]\n" +
                "    }\n" +
                "]";
    }
}
