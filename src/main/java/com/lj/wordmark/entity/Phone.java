package com.lj.wordmark.entity;

import com.lj.wordmark.annotation.MarkField;
import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * @Author luojing
 * @Date 2022/4/17
 */
@Data
@AllArgsConstructor
public class Phone {
    @MarkField("手机号")
    private String num;
    @MarkField("手机类型")
    private String type;
}
