package com.liutao.excel.model;

import com.liutao.excel.common.annotation.ExcelColumn;

import java.math.BigDecimal;

/**
 * 卡导出实体
 */
public class CardExport {
    // 编号
    @ExcelColumn(name = "编号", index = 0, width = 20)
    private String cardNum;
    // 部门
    @ExcelColumn(name = "部门", index = 1, width = 30)
    private String department;
    // 姓名
    @ExcelColumn(name = "姓名", index = 2, width = 20)
    private String userName;
    // 手机号
    @ExcelColumn(name = "手机号", index = 3, width = 25)
    private String phoneNumber;
    // 金额
    @ExcelColumn(name = "金额", index = 4, width = 10)
    private BigDecimal price;

    public String getCardNum() {
        return cardNum;
    }

    public void setCardNum(String cardNum) {
        this.cardNum = cardNum;
    }

    public String getDepartment() {
        return department;
    }

    public void setDepartment(String department) {
        this.department = department;
    }

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public String getPhoneNumber() {
        return phoneNumber;
    }

    public void setPhoneNumber(String phoneNumber) {
        this.phoneNumber = phoneNumber;
    }

    public BigDecimal getPrice() {
        return price;
    }

    public void setPrice(BigDecimal price) {
        this.price = price;
    }
}
