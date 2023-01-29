package com.muyangren;

import com.muyangren.entity.Templates;
import com.muyangren.utils.FileUtil;
import com.muyangren.utils.HtmlUtil;
import com.muyangren.utils.WordUtil;
import fr.opensagres.poi.xwpf.converter.core.FileImageExtractor;
import fr.opensagres.poi.xwpf.converter.core.FileURIResolver;
import fr.opensagres.poi.xwpf.converter.core.XWPFConverterException;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

class ImportWord {

    /**
     * 导入模板，获取富文本信息以及非富文本信息
     */
    @Test
    void importWord() throws IOException {
        Templates templates = new Templates();
        //1、获取文件路径
        //实际情况中，该MultipartFile 是Controller层获取的(如语雀所示)
        File file = new ClassPathResource("templates" + File.separator + "牧羊人导入模板.docx").getFile();
        MultipartFile fileItem = FileUtil.createFileItem(file,file.toString());
        //2、处理非富文本信息
        List<Map<String, String>> mapList = WordUtil.readWord(fileItem);
        //3、下载文件到临时路径下
        File destFile = FileUtil.fileDownloadToLocalPath(fileItem);
        //4、处理富文本信息(含图片),并且定义一个实体类接收
        dealWithTemplatesRichText(templates,destFile);
        //5、替换案例富文本信息中的图片(如果有)路径并删除临时文件和临时图片
        dealWithTemplatesRichTextToPicture(templates);
        //6、最后调用方法保存templates即可
    }

    private void dealWithTemplatesRichTextToPicture(Templates templates) {
        //保存临时图片路径，统一删除
        Set<File> files = new HashSet<>();
        String richTextOne = templates.getRichTextOne();
        String richTextTwo = templates.getRichTextTwo();
        String richTextThree = templates.getRichTextThree();
        if (StringUtils.isNotEmpty(richTextOne)) {
            //参数分别是实体类、minio（其他）方法、临时图片路径集合
            String content = dealWithTemplatesRichTextToPictureChild(richTextOne, "ossBuilder", files);
            templates.setRichTextOne(content);
        }
        if (StringUtils.isNotEmpty(richTextTwo)) {
            String content = dealWithTemplatesRichTextToPictureChild(richTextTwo, "ossBuilder", files);
            templates.setRichTextTwo(content);
        }
        if (StringUtils.isNotEmpty(richTextThree)) {
            String content = dealWithTemplatesRichTextToPictureChild(richTextThree, "ossBuilder", files);
            templates.setRichTextThree(content);
        }
        if (files.size() > 0) {
            for (File file : files) {
                //自行删除
                //FileUtil.deleteQuietly(file);
            }
        }
    }

    private String dealWithTemplatesRichTextToPictureChild(String content, String ossBuilder, Set<File> files) {
        //1、正则匹配src
        List<String> imagesFiles = HtmlUtil.regexMatchPicture(content);
        //2、判空
        if (imagesFiles.size()>0) {
            for (String imagesFile : imagesFiles) {
                File file = new File(imagesFile);
                MultipartFile fileItem = FileUtil.createFileItem(file, file.getName());
                boolean aBoolean = true;
                while (Boolean.TRUE.equals(aBoolean)) {
                    //这里是调用上传接口返回一个url，然后替换即可
                    //BladeFile bladeFile = ossBuilder.template().putFile(fileItem);
                    //if (Func.isNotEmpty(bladeFile)) {
                    //    String link = bladeFile.getLink();
                    //    content = content.replace(imagesFile, link);
                        //删除临时图片(统一删除 如上传同一张图片，第二次会找不到图片)
                        //files.add(file);
                        //aBoolean = false;
                    //}
                }
            }
        }
        return content;
    }

    private void dealWithTemplatesRichText(Templates templates, File destFile) {
        if (!destFile.exists()) {
            //导入模板失败,请重新上传！
        } else {
            //1、判断是否为docx文件
            if (destFile.getName().endsWith(".docx") || destFile.getName().endsWith(".DOCX")) {
                // 1)加载word文档生成XWPFDocument对象
                try (FileInputStream in = new FileInputStream(destFile); XWPFDocument document = new XWPFDocument(in)) {
                    // 2)解析XHTML配置（这里设置IURIResolver来设置图片存放的目录）
                    File imageFolderFile = new File(String.valueOf(destFile.getParentFile()));
                    XHTMLOptions options = XHTMLOptions.create().URIResolver(new FileURIResolver(imageFolderFile));
                    options.setExtractor(new FileImageExtractor(imageFolderFile));
                    options.setIgnoreStylesIfUnused(false);
                    options.setFragment(true);
                    //2、使用字符数组流获取解析的内容
                    ByteArrayOutputStream bass = new ByteArrayOutputStream();
                    XHTMLConverter.getInstance().convert(document, bass, options);
                    String richTextContent = bass.toString();
                    //自定义用于截取的波浪线（可以用其他代替，只要你能够区分）
                    String[] tableSplit = richTextContent.split("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~</span></p>");
                    //实际项目中需要判断长度是否与预期的一样，避免模板被破坏了，导致内容下表有所改动
                    int length = tableSplit.length;
                    if (length==7){
                        templates.setRichTextOne(tableSplit[2]);
                        templates.setRichTextTwo(tableSplit[4]);
                        templates.setRichTextThree(tableSplit[6]);
                    }else {
                        //导入模板被破坏,请重新上传！
                    }
                } catch (IOException | XWPFConverterException e) {
                    e.printStackTrace();
                } finally {
                    //处理完就删除文件
                    //FileUtil.deleteQuietly(destFile);
                }
            }
        }

    }
}
