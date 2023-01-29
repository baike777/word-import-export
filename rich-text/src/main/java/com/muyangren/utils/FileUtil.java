package com.muyangren.utils;

import java.io.*;
import java.rmi.ServerException;
import java.util.UUID;


import org.apache.commons.lang3.StringUtils;
import org.apache.http.entity.ContentType;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileItemFactory;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;

import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.commons.CommonsMultipartFile;


/**
 * @author: muyangren
 * @Date: 2023/1/14
 * @Description: com.muyangren.utils
 * @Version: 1.0
 */
public interface FileUtil {
    /**
     * 创建FileItem
     * @param file
     * @param fieldName
     * @return
     */
    public static MultipartFile createFileItem(File file, String fieldName) {
        FileItemFactory factory = new DiskFileItemFactory(16, null);
        FileItem item = factory.createItem(fieldName, ContentType.MULTIPART_FORM_DATA.toString(), true, file.getName());
        int bytesRead = 0;
        byte[] buffer = new byte[8192];
        try {
            FileInputStream fis = new FileInputStream(file);
            OutputStream os = item.getOutputStream();
            while ((bytesRead = fis.read(buffer, 0, 8192)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
            os.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return new CommonsMultipartFile(item);
    }
    /**
     * 下载到本地路径
     * @param file
     * @return
     * @throws IOException
     */
    public static File fileDownloadToLocalPath(MultipartFile file) {
        File destFile = null;
        try {
            //获取文件名称
            if (StringUtils.isEmpty(file.getOriginalFilename())){
                throw new ServerException("导入模板失败！");
            }
            String fileName = file.getOriginalFilename();
            //获取文件后缀
            String pref = fileName.lastIndexOf(".") != -1 ? fileName.substring(fileName.lastIndexOf(".") + 1) : null;
            //临时文件
            //临时文件名避免重复
            String uuidFile = UUID.randomUUID().toString().replace("-", "") + "." + pref;
            destFile = new File(FileUtil.getProjectPath() + uuidFile);
            if (!destFile.getParentFile().exists()) {
                destFile.getParentFile().mkdirs();
            }
            file.transferTo(destFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return destFile;
    }

    /**
     * @return 文件路径
     */
    public static String getProjectPath(){
        String os = System.getProperty("os.name").toLowerCase();
        //windows下
        if (os.indexOf("windows")>=0) {
            return "C://temp/";
        }else{
            return "/usr/local/temp/";
        }
    }
}
