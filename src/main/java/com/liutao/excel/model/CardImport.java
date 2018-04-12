package com.liutao.excel.model;

import com.liutao.excel.common.annotation.ExcelColumn;

import java.math.BigDecimal;

/**
 * 卡导入实体
 */
public class CardImport {
    // 编号
    @ExcelColumn(index = 0)
    private String cardNum;
    // 部门
    @ExcelColumn(index = 1)
    private String department;
    // 姓名
    @ExcelColumn(index = 2)
    private String userName;
    // 手机号
    @ExcelColumn(index = 3)
    private String phoneNumber;
    // 金额
    @ExcelColumn(index = 4)
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
