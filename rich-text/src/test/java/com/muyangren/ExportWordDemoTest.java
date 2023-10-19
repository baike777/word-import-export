package com.muyangren;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.muyangren.entity.Templates;
import com.muyangren.enums.StyleEnum;
import com.muyangren.utils.FileUtil;
import com.muyangren.utils.HtmlUtil;
import org.ddr.poi.html.HtmlRenderPolicy;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.*;


/**
 * @author: muyangren
 * @Date: 2023/1/16
 * @Description: 导出word文档 word文档和富文本是两种东西，所以肯定存在一些兼容问题，仅能支持大部分样式
 * @Version: 1.0
 */
class ExportWordDemoTest {
    @Test
    void exportWord() throws IOException {
        // 1、初始化数据
        Templates templates = initTemplates();

        // 2、填充数据
        Map<String, Object> data = new HashMap<>(8);
        data.put("title", templates.getTitle());
        data.put("basicInfoOne", templates.getBasicInfoOne());
        data.put("basicInfoTwo", templates.getBasicInfoTwo());
        data.put("basicInfoThree", templates.getBasicInfoThree());

        // 3、判断图片是否超出word最大宽度，否则需要等比缩小
        // 3.1、处理前图片样式
        //data.put("richText1", templates.getRichTextOne());
        //data.put("richText2", templates.getRichTextTwo());
        //data.put("richText3", templates.getRichTextThree());

        // 3.2、处理后图片样式
        data.put("richText1", dealWithPictureWidthAndHeight(templates.getRichTextOne()));
        data.put("richText2", dealWithPictureWidthAndHeight(templates.getRichTextTwo()));
        data.put("richText3", dealWithPictureWidthAndHeight(templates.getRichTextThree()));

        // 4、指定插件
        HtmlRenderPolicy htmlRenderPolicy = new HtmlRenderPolicy();
        ConfigureBuilder builder = Configure.builder();
        Configure config = builder.build();
        config.customPolicy("richText1", htmlRenderPolicy);
        config.customPolicy("richText2", htmlRenderPolicy);
        config.customPolicy("richText3", htmlRenderPolicy);

        // 临时路径，实际项目里不需要
        String filePath = FileUtil.getProjectPath() + "document" + File.separator + "牧羊人导出实例.docx";
        // 5、动态导出
        dynamicExportWord(data, new File(filePath), new ClassPathResource("templates" + File.separator + "exportWordTemplate.docx"), config, null, false);
    }

    private Templates initTemplates() {
        // 1、数据来源,实际项目里是通过sql查询出来的
        Templates templates = new Templates();
        templates.setTitle("标题1");
        templates.setBasicInfoOne("基本信息1");
        templates.setBasicInfoTwo("基本信息2");
        templates.setBasicInfoThree("基本信息3");

        // 随便拿一张是图片的就行
        File file = new File("C:/temp/word/media/image1.jpeg");
        // 1.1、本地路径，不支持  解决方法如下：1)转为base64(又长又臭 依托答辩)、2)上传到自己文件服务器里
        templates.setRichTextOne("<p style=\"white-space:pre-wrap;\"><span style=\"font-family:'等线';white-space:pre-wrap;\">V50kfc 我倒要试试有多好吃</span></p><p style=\"white-space:pre-wrap;\"><img src=" + file.toString() + " style=\"width:194.1pt;height:126.05pt;\"/><span id=\"_GoBack\"/></p><p style=\"white-space:pre-wrap;\"><span style=\"font-family:'等线';white-space:pre-wrap;\">");

        // 1.2、base64格式
        String baseString = base64String(file);
        baseString = "data:image/jpeg;base64," + baseString;
        templates.setRichTextTwo("<p style=\"white-space:pre-wrap;\"><span style=\"font-family:'等线';white-space:pre-wrap;\">V50kfc 我倒要试试有多好吃</span></p><p style=\"white-space:pre-wrap;\"><img src=" + baseString + " style=\"width:194.1pt;height:126.05pt;\"/><span id=\"_GoBack\"/></p><p style=\"white-space:pre-wrap;\"><span style=\"font-family:'等线';white-space:pre-wrap;\">");

        // 1.3、实际项目中,前端富文本获取图片,会调取上传接口返回url,所以此块不用太纠结，如果是存储在某个路径下，那么在导出时 后端处理转为base64即可
        templates.setRichTextThree("<p><img src=\"http://xxx.xxx.xxx.xxxx:80/upload/20230111/image1.jpeg\" style=\"height:1600px; width:2560px\" /></p>");
        return templates;
    }

    private String dealWithPictureWidthAndHeight(String content) {
        List<HashMap<String, String>> imagesFiles = HtmlUtil.regexMatchWidthAndHeight(content);
        if (!imagesFiles.isEmpty()) {
            for (HashMap<String, String> imagesFile : imagesFiles) {
                String newFileUrl = imagesFile.get(StyleEnum.NEW_FILE_URL.getValue());
                String fileUrl = imagesFile.get(StyleEnum.FILE_URL.getValue());
                if (newFileUrl != null) {
                    content = content.replace(fileUrl, newFileUrl);
                }
            }
        }
        return content;
    }

    public String base64String(File file) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (FileInputStream fileInputStream = new FileInputStream(file)) {
            byte[] bytes = new byte[1024];
            //用来定义一个准备接收图片总长度的局部变量
            int len;
            //将流的内容读取到bytes中
            while ((len = fileInputStream.read(bytes)) > 0) {
                //将bytes内存中的内容从0开始到总长度输出出去
                out.write(bytes, 0, len);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        //通过util包中的Base64类对字节数组进行base64编码
        return Base64.getEncoder().encodeToString(out.toByteArray());
    }

    /**
     * @param data      填充数据
     * @param filePath  临时路径
     * @param resource  模板路径
     * @param response  通过浏览器下载需要
     * @param isBrowser true-通过浏览器下载  false-下载到临时路径
     */
    public void dynamicExportWord(Map<String, Object> data, File filePath, Resource resource, Configure config, HttpServletResponse response, boolean isBrowser) throws IOException {
        XWPFTemplate render = XWPFTemplate.compile(resource.getInputStream(), config).render(data);
        if (Boolean.FALSE.equals(isBrowser)) {
            File parentFile = filePath.getParentFile();
            if (!parentFile.exists()) {
                parentFile.mkdirs();
            }
            FileOutputStream out = new FileOutputStream(filePath);
            render.writeAndClose(out);
        } else {
            ServletOutputStream out = response.getOutputStream();
            render.writeAndClose(out);
        }
    }
}
