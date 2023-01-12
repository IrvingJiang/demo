package com.example.demo.excel.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;

import java.math.BigDecimal;

/**
 * @Description:右表数据
 * @author:Irving
 */
public class Right {

    /**
     * 是否纳入年度外协采购计划
     * 枚举
     * 纳入年度外协采购计划
     * 单报单批
     */
    @Excel(name = "是否纳入年度外协采购计划（下拉框：纳入年度外协采购计划/单报单批）")
    private String sfnrndwxcgjh;


    /**
     * 外协采购计划名称
     */
    @Excel(name="外协采购计划名称（年初计划由系统导入作，单报单批手动填）")
    private String wxcgjhmc;

    /**
     * 外协经费预算
     */
    @Excel(name="外协经费预算（年初计划由系统自动带入）")
    private BigDecimal wxjfys;
    /**
     * 外协经费支出
     */
    @Excel(name="外协经费支出")
    private BigDecimal wxjfzc;

    /**
     * 外协任务进度安排
     */
    @Excel(name="外协任务进度安排")
    private String wxrwjdap;

    /**
     * 外协任务完成情况
     */
    @Excel(name="外协任务完成情况（含进度情况、存在问题、下月工作计划等）")
    private String wxrwwcqk;

    /**
     * 备注
     */
    @Excel(name="备注")
        private String bz;

    public String getSfnrndwxcgjh() {
        return sfnrndwxcgjh;
    }

    public void setSfnrndwxcgjh(String sfnrndwxcgjh) {
        this.sfnrndwxcgjh = sfnrndwxcgjh;
    }

    public String getWxcgjhmc() {
        return wxcgjhmc;
    }

    public void setWxcgjhmc(String wxcgjhmc) {
        this.wxcgjhmc = wxcgjhmc;
    }

    public BigDecimal getWxjfys() {
        return wxjfys;
    }

    public void setWxjfys(BigDecimal wxjfys) {
        this.wxjfys = wxjfys;
    }

    public BigDecimal getWxjfzc() {
        return wxjfzc;
    }

    public void setWxjfzc(BigDecimal wxjfzc) {
        this.wxjfzc = wxjfzc;
    }

    public String getWxrwjdap() {
        return wxrwjdap;
    }

    public void setWxrwjdap(String wxrwjdap) {
        this.wxrwjdap = wxrwjdap;
    }

    public String getWxrwwcqk() {
        return wxrwwcqk;
    }

    public void setWxrwwcqk(String wxrwwcqk) {
        this.wxrwwcqk = wxrwwcqk;
    }

    public String getBz() {
        return bz;
    }

    public void setBz(String bz) {
        this.bz = bz;
    }
}
