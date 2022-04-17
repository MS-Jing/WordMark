package com.lj.wordmark.entity;

import com.lj.wordmark.annotation.MarkField;
import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * @Author luojing
 * @Date 2022/4/15
 */
@Data
@AllArgsConstructor
public class User {
    private Integer index;
    @MarkField("姓名")
    private String name;
    private Integer age;
    private String hobby;
}
