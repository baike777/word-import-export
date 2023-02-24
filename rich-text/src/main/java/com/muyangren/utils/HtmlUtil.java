package com.muyangren.utils;

import com.muyangren.enums.StyleEnum;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author: muyangren
 * @Date: 2023/1/13
 * @Description: com.muyangren.utils
 * @Version: 1.0
 */
public class HtmlUtil {

    /**
     * 匹配字符串中的img标签正则表达式
     */
    private static final String IMG_STYPE_REGULAR ="<(img|IMG)(.*?)(>|></img>|/>)";

    /**
     * 匹配字符串中的src标签正则表达式
     */
    private static final String URL_REGULAR = "(src|SRC)=(\"|\')(.*?)(\"|\')";

    /**
     * 匹配字符串中的style标签中的宽高(关键看前端用的是什么富文本)
     */
    private static final String WIDTH_REGEX = "width:(?<width>\\d+([.]\\d+)?)(px|pt)";
    //private static final String WIDTH_REGEX = "width:(?<width>\\d+([.]\\d+)?)(px;|pt;)";

    private static final String HEIGHT_REGEX = "height:(?<height>\\d+([.]\\d+)?)(px;|pt;)";

    /**
     * 通过正则表达式去获取html中的src
     *
     * @param content
     * @return
     */
    public static List<String> regexMatchPicture(String content) {
        //1、用来存储获取到的图片地址
        List<String> imgList = new ArrayList<>();
        //2、匹配字符串中的img标签
        Pattern p = Pattern.compile(IMG_STYPE_REGULAR);
        Matcher matcher = p.matcher(content);
        boolean hasPic = matcher.find();
        //3、判断是否含有图片
        if (hasPic) {
            //3.1、如果含有图片，那么持续进行查找，直到匹配不到
            while (hasPic) {
                //3.2、获取第二个分组的内容，也就是 (.*?)匹配到的
                String group = matcher.group(2);
                //3.3、匹配图片的地址
                Pattern srcText = Pattern.compile(URL_REGULAR);
                Matcher matcher2 = srcText.matcher(group);
                if (matcher2.find()) {
                    //3.3.1、把获取到的图片地址添加到列表中
                    imgList.add(matcher2.group(3));
                }
                //3.4、判断是否还有img标签
                hasPic = matcher.find();
            }
        }
        return imgList;
    }

    /**
     * 通过正则表达式去获取html中的src中的宽高
     *
     * @param content
     * @return
     */
    public static List<HashMap<String, String>> regexMatchWidthAndHeight(String content) {
        //1、用来存储获取到的图片地址
        List<HashMap<String, String>> srcList = new ArrayList<>();

        //2、匹配字符串中的img标签
        Pattern p = Pattern.compile(IMG_STYPE_REGULAR);

        Matcher matcher = p.matcher(content);
        boolean hasPic = matcher.find();
        //3、判断是否含有图片
        if (hasPic) {
            //3.1、如果含有图片，那么持续进行查找，直到匹配不到
            while (hasPic) {
                HashMap<String, String> hashMap = new HashMap<>(4);
                //3.1.2、获取第二个分组的内容，也就是 (.*?)匹配到的
                String group = matcher.group(2);
                hashMap.put(StyleEnum.FILE_URL.getValue(), group);
                //3.1.2、匹配图片的地址
                Pattern srcText = Pattern.compile(WIDTH_REGEX);
                Matcher matcher2 = srcText.matcher(group);
                String imgWidth = null;
                String imgHeight = null;
                if (matcher2.find()) {
                    imgWidth = matcher2.group(StyleEnum.WIDTH.getValue());
                }
                srcText = Pattern.compile(HEIGHT_REGEX);
                matcher2 = srcText.matcher(group);
                if (matcher2.find()) {
                    imgHeight = matcher2.group(StyleEnum.HEIGHT.getValue());
                }
                hashMap.put(StyleEnum.WIDTH.getValue(), imgWidth);
                hashMap.put(StyleEnum.HEIGHT.getValue(), imgHeight);
                srcList.add(hashMap);
                //判断是否还有img标签
                hasPic = matcher.find();
            }

            //4、判断srcList是否为空
            if (srcList.isEmpty()){
                return Collections.emptyList();
            }
            dealWithSrcList(srcList);
        }
        return srcList;
    }

    public static void dealWithSrcList(List<HashMap<String, String>> srcList){
        for (HashMap<String, String> imagesFile : srcList) {
            String height = imagesFile.get(StyleEnum.HEIGHT.getValue());
            String width = imagesFile.get(StyleEnum.WIDTH.getValue());
            String fileUrl = imagesFile.get(StyleEnum.FILE_URL.getValue());
            //注：
            //1）若图片宽度超过word的最大宽值,图片将会变形、所以通过等比缩小来处理
            //2）word文档中：宽大最大650px(17厘米)
            //3）有的标签返回时pt像素,此处主要看前端是用那种富文本了、可能有出现以下几种情况1、height:auto (正则匹配不到) 2、无宽高标签
            if (width != null) {
                BigDecimal widthDecimal = new BigDecimal(width);
                BigDecimal maxWidthWord = new BigDecimal(StyleEnum.MAX_WIDTH.getValue());
                //超过最大值
                if (widthDecimal.compareTo(maxWidthWord) > 0) {
                    BigDecimal divide = widthDecimal.divide(maxWidthWord, 2, RoundingMode.HALF_UP);
                    //设置宽值为650px
                    fileUrl = fileUrl.replace("width:" + width, "width:" + maxWidthWord);
                    //特殊处理高
                    if (height !=null) {
                        BigDecimal heightDecimal = new BigDecimal(height);
                        BigDecimal divide1 = heightDecimal.divide(divide, 1, RoundingMode.HALF_UP);
                        fileUrl = fileUrl.replace("height:" + height, "height:" + divide1);
                    } else {
                        fileUrl = fileUrl.replace("height:auto", "height:350px");
                    }
                    imagesFile.put(StyleEnum.NEW_FILE_URL.getValue(), fileUrl);
                } else {
                    imagesFile.put(StyleEnum.NEW_FILE_URL.getValue(), null);
                }
            }
        }
    }
}
