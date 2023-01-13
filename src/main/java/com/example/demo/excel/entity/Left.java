package com.example.demo.excel.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelCollection;

import java.math.BigDecimal;
import java.util.List;

/**
 * @Description:左表数据
 * @author:Irving
 */
public class Left {
    @Excel(name="序号")
    private Integer index;

    @Excel(name="月份")
    private String yf;

    /**
     * 部门
     */
    @Excel(name="部门")
    private String bm;


    /**
     * 项目名称 系统导入
     */
    @Excel(name="项目名称（系统导入）")
     private String xmmc;

    /**
     * 项目负责人
     */
    @Excel(name="项目负责人")
    private String xmfzr;


    /**
     * 项目年度
     */
    @Excel(name="项目年度")
    private Integer xmnd;

    /**
     * 资金来源 枚举
     * 财政资金/
     * 银行存款专项/
     * 自有资金
     */
    @Excel(name = "资金来源（下拉框，财政资金/银行存款专项/自有资金）")
    private String zjly;


    /**
     * 项目预算（系统导入）
     */
    @Excel(name="项目预算（系统导入）")
    private BigDecimal xmys;


    /**
     * 预算支出
     */
    @Excel(name="预算支出")
    private BigDecimal yszc;

    /**
     * 预算执行率
     */
    @Excel(name="预算执行率（=预算支出/项目预算。由系统自动计算得出）")
    private String yszxl;


    /**
     * 项目任务完成情况
     */
    @Excel(name="项目任务完成情况（含进度情况、绩效目标完成情况等）")
    private String xmrwwcqk;

    /**
     * 项目任务完成率
     */
    @Excel(name="项目任务完成率（%）")
    private String xmrwwcl;

    /**
     * 存在问题
     */
    @Excel(name="存在问题")
    private String czwt;

    /**
     * 下月工作计划
     */
    @Excel(name="下月工作计划")
    private String xygzjh;


    @ExcelCollection(name="外协采购计划进度")
    private List<Right> rights;

    public String getBm() {
        return bm;
    }

    public void setBm(String bm) {
        this.bm = bm;
    }

    public String getXmmc() {
        return xmmc;
    }

    public void setXmmc(String xmmc) {
        this.xmmc = xmmc;
    }

    public String getXmfzr() {
        return xmfzr;
    }

    public void setXmfzr(String xmfzr) {
        this.xmfzr = xmfzr;
    }

    public Integer getXmnd() {
        return xmnd;
    }

    public void setXmnd(Integer xmnd) {
        this.xmnd = xmnd;
    }

    public String getZjly() {
        return zjly;
    }

    public void setZjly(String zjly) {
        this.zjly = zjly;
    }

    public BigDecimal getXmys() {
        return xmys;
    }

    public void setXmys(BigDecimal xmys) {
        this.xmys = xmys;
    }

    public BigDecimal getYszc() {
        return yszc;
    }

    public void setYszc(BigDecimal yszc) {
        this.yszc = yszc;
    }

    public String getYszxl() {
        return yszxl;
    }

    public void setYszxl(String yszxl) {
        this.yszxl = yszxl;
    }

    public String getXmrwwcqk() {
        return xmrwwcqk;
    }

    public void setXmrwwcqk(String xmrwwcqk) {
        this.xmrwwcqk = xmrwwcqk;
    }

    public String getXmrwwcl() {
        return xmrwwcl;
    }

    public void setXmrwwcl(String xmrwwcl) {
        this.xmrwwcl = xmrwwcl;
    }

    public String getCzwt() {
        return czwt;
    }

    public void setCzwt(String czwt) {
        this.czwt = czwt;
    }

    public String getXygzjh() {
        return xygzjh;
    }

    public void setXygzjh(String xygzjh) {
        this.xygzjh = xygzjh;
    }

    public Integer getIndex() {
        return index;
    }

    public void setIndex(Integer index) {
        this.index = index;
    }

    public List<Right> getRights() {
        return rights;
    }

    public void setRights(List<Right> rights) {
        this.rights = rights;
    }

    public String getYf() {
        return yf;
    }

    public void setYf(String yf) {
        this.yf = yf;
    }
}
