package com.lijia.excel;

import com.github.crab2died.annotation.ExcelField;

public class HanZheng {

    @ExcelField(title = "序号", order = 1)
    private String id;

    @ExcelField(title = "基金名称", order = 2)
    private String jjmc;

    @ExcelField(title = "批次号", order = 3)
    private String pch;

    @ExcelField(title = "序列号", order = 4)
    private String xlh;

    @ExcelField(title = "索引号", order = 5)
    private String syh;

    @ExcelField(title = "项目编号", order = 6)
    private String xmbh;

    @ExcelField(title = "被投资公司名称", order = 7)
    private String btzgsmc;

    @ExcelField(title = "联系人", order = 8)
    private String lxr;

    @ExcelField(title = "被投资公司地址", order = 9)
    private String btzgsdz;

    @ExcelField(title = "联系电话", order = 10)
    private String lxdh;

    @ExcelField(title = "投资类型", order = 11)
    private String tzlx;

    @ExcelField(title = "No. of shares", order = 12)
    private String nos;

    @ExcelField(title = "Share Percentage", order = 13)
    private String sp;

    @ExcelField(title = "Cost", order = 14)
    private String cost;

    @ExcelField(title = "Accrued Interest-Before 2019", order = 15)
    private String aib;

    @ExcelField(title = "Accrued Interest-During 2019", order = 16)
    private String aid;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getJjmc() {
        return jjmc;
    }

    public void setJjmc(String jjmc) {
        this.jjmc = jjmc;
    }

    public String getPch() {
        return pch;
    }

    public void setPch(String pch) {
        this.pch = pch;
    }

    public String getXlh() {
        return xlh;
    }

    public void setXlh(String xlh) {
        this.xlh = xlh;
    }

    public String getSyh() {
        return syh;
    }

    public void setSyh(String syh) {
        this.syh = syh;
    }

    public String getXmbh() {
        return xmbh;
    }

    public void setXmbh(String xmbh) {
        this.xmbh = xmbh;
    }

    public String getBtzgsmc() {
        return btzgsmc;
    }

    public void setBtzgsmc(String btzgsmc) {
        this.btzgsmc = btzgsmc;
    }

    public String getLxr() {
        return lxr;
    }

    public void setLxr(String lxr) {
        this.lxr = lxr;
    }

    public String getBtzgsdz() {
        return btzgsdz;
    }

    public void setBtzgsdz(String btzgsdz) {
        this.btzgsdz = btzgsdz;
    }

    public String getLxdh() {
        return lxdh;
    }

    public void setLxdh(String lxdh) {
        this.lxdh = lxdh;
    }

    public String getTzlx() {
        return tzlx;
    }

    public void setTzlx(String tzlx) {
        this.tzlx = tzlx;
    }

    public String getNos() {
        return nos;
    }

    public void setNos(String nos) {
        this.nos = nos;
    }

    public String getSp() {
        return sp;
    }

    public void setSp(String sp) {
        this.sp = sp;
    }

    public String getCost() {
        return cost;
    }

    public void setCost(String cost) {
        this.cost = cost;
    }

    public String getAib() {
        return aib;
    }

    public void setAib(String aib) {
        this.aib = aib;
    }

    public String getAid() {
        return aid;
    }

    public void setAid(String aid) {
        this.aid = aid;
    }
}
