package com.wj;

import com.wj.annotation.Property;

public class Item {

    @Property(columnName = "姓名")
    private String name;

    @Property(columnName = "性别")
    private String sex;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }
}
