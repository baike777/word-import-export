package com.muyangren.entity;

import lombok.Data;
import lombok.EqualsAndHashCode;

import java.io.Serializable;

/**
 * @author: muyangren
 * @Date: 2023/1/16
 * @Description: com.muyangren.entity
 * @Version: 1.0
 */
@Data
@EqualsAndHashCode(callSuper = false)
public class Templates implements Serializable {
    private static final long serialVersionUID = 1L;

    //标题
    private String title;
    //基本信息1
    private String basicInfoOne;
    //基本信息2
    private String basicInfoTwo;
    //基本信息3
    private String basicInfoThree;

    //富文本1
    private String richTextOne;
    //富文本2
    private String richTextTwo;
    //富文本3
    private String richTextThree;


}
