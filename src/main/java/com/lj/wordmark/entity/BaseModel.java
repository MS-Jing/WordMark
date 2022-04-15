package com.lj.wordmark.entity;

import lombok.Data;

import java.time.LocalDate;

/**
 * @Author luojing
 * @Date 2022/4/14
 */
@Data
public class BaseModel {
    private LocalDate createTime;
    private LocalDate updateTime;
}
