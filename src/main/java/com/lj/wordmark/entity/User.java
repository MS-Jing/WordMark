package com.lj.wordmark.entity;

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
    private String name;
    private Integer age;
    private String hobby;
}
