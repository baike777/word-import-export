package com.muyangren.utils;

import fr.opensagres.odfdom.converter.core.utils.StringUtils;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
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
         * 通过正则表达式去获取html中的src
         *
         * @param content
         * @return
         */
        public static List<String> regexMatchPicture(String content) {
        //用来存储获取到的图片地址
        List<String> srcList = new ArrayList<>();
        //匹配字符串中的img标签
        Pattern p = Pattern.compile("<(img|IMG)(.*?)(>|></img>|/>)");
        Matcher matcher = p.matcher(content);
        boolean hasPic = matcher.find();
        //判断是否含有图片
        if (hasPic) {
            //如果含有图片，那么持续进行查找，直到匹配不到
            while (hasPic) {
                //获取第二个分组的内容，也就是 (.*?)匹配到的
                String group = matcher.group(2);
                //匹配图片的地址
                Pattern srcText = Pattern.compile("(src|SRC)=(\"|\')(.*?)(\"|\')");
                Matcher matcher2 = srcText.matcher(group);
                if (matcher2.find()) {
                    //把获取到的图片地址添加到列表中
                    srcList.add(matcher2.group(3));
                }
                //判断是否还有img标签
                hasPic = matcher.find();
            }
        }
        return srcList;
    }

        /**
         * 通过正则表达式去获取html中的src中的宽高
         *
         * @param content
         * @return
         */
        public static List<HashMap<String, String>> regexMatchWidthAndHeight(String content) {
        //用来存储获取到的图片地址
        List<HashMap<String, String>> srcList = new ArrayList<>();
        //匹配字符串中的img标签
        Pattern p = Pattern.compile("<(img|IMG)(.*?)(>|></img>|/>)");
        //匹配字符串中的style标签中的宽高(关键看前端用的是什么富文本)
        String regexWidth = "width:(?<width>\\d+([.]\\d+)?)(px|pt)";
        //String regexWidth = "width:(?<width>\\d+([.]\\d+)?)(px;|pt;)";
        String regexHeight = "height:(?<height>\\d+([.]\\d+)?)(px;|pt;)";
        Matcher matcher = p.matcher(content);
        boolean hasPic = matcher.find();
        //判断是否含有图片
        if (hasPic) {
            //如果含有图片，那么持续进行查找，直到匹配不到
            while (hasPic) {
                HashMap<String, String> hashMap = new HashMap<>();
                //获取第二个分组的内容，也就是 (.*?)匹配到的
                String group = matcher.group(2);
                hashMap.put("fileUrl", group);
                //匹配图片的地址
                Pattern srcText = Pattern.compile(regexWidth);
                Matcher matcher2 = srcText.matcher(group);
                String imgWidth = null;
                String imgHeight = null;
                if (matcher2.find()) {
                    imgWidth = matcher2.group("width");
                }
                srcText = Pattern.compile(regexHeight);
                matcher2 = srcText.matcher(group);
                if (matcher2.find()) {
                    imgHeight = matcher2.group("height");
                }
                hashMap.put("width", imgWidth);
                hashMap.put("height", imgHeight);
                srcList.add(hashMap);
                //判断是否还有img标签
                hasPic = matcher.find();
            }
            for (HashMap<String, String> imagesFile : srcList) {
                String height = imagesFile.get("height");
                String width = imagesFile.get("width");
                String fileUrl = imagesFile.get("fileUrl");
                //注：该处是避免图片超过word的最大宽值，若超过则按比例缩小,判断主要是根据实际业务中前端富文本插件对图片的处理，项目中出现过无宽度，或者高度非px数值,所以特殊处理下
                //1厘米=25px(像素)  17厘米(650px) word最大宽值
                if (StringUtils.isNotEmpty(width)) {
                    BigDecimal widthDecimal = new BigDecimal(width);
                    BigDecimal maxWidthWord = new BigDecimal("650.0");
                    if (widthDecimal.compareTo(maxWidthWord) > 0) {
                        BigDecimal divide = widthDecimal.divide(maxWidthWord, 2, RoundingMode.HALF_UP);
                        fileUrl = fileUrl.replace("width:" + width, "width:" + maxWidthWord);
                        if (StringUtils.isNotEmpty(height)) {
                            BigDecimal heightDecimal = new BigDecimal(height);
                            BigDecimal divide1 = heightDecimal.divide(divide, 1, RoundingMode.HALF_UP);
                            fileUrl = fileUrl.replace("height:" + height, "height:" + divide1);
                        } else {
                            fileUrl = fileUrl.replace("height:auto", "height:350px");
                        }
                        imagesFile.put("newFileUrl", fileUrl);
                    } else {
                        imagesFile.put("newFileUrl", "");
                    }
                }
            }
        }
        return srcList;
    }

}
