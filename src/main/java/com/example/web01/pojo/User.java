package com.example.web01.pojo;

import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableName;
import lombok.Data;

@Data
@TableName("user")
public class User {
    private Long id;
    @TableField("name")
    private String name;
    private Integer age;
    private String email;
}
