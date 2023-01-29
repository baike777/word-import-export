package com.muyangren.utils;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * @author: muyangren
 * @Date: 2023/1/14
 * @Description: com.muyangren.utils
 * @Version: 1.0
 */
public class WordUtil {

    static final String suffix = ".docx";


    /**
     * 文档返回数据
     * @param file
     * @return
     * @throws Exception
     */
    public static List<Map<String, String>> readWord(MultipartFile file){
        try {
            InputStream in = file.getInputStream();
            System.out.println("解析的文件名："+file.getOriginalFilename());
            if(file.getOriginalFilename().toLowerCase().endsWith(suffix)){
                // 2007+版
                return readWord2007(in);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 该读取方法针对导入案例
     * @param in
     * @return
     * @throws IOException
     */
    public static List<Map<String, String>> readWord2007(InputStream in) throws IOException {
        XWPFDocument document = new XWPFDocument(in);
        List<Map<String, String>> list = new ArrayList<>();
        try {
            Iterator<XWPFTable> iterator = document.getTablesIterator();
            XWPFTable table = iterator.next();
            // 获取到表格的行数
            List<XWPFTableRow> rows = table.getRows();
            // 读取表格的每一行
            Map<String, String> map = new LinkedHashMap<>();
            for (int i = 0; i < rows.size(); i++) {
                XWPFTableRow row = rows.get(i);
                List<XWPFTableCell> cells = row.getTableCells();
                String title = cells.get(0).getText();
                if(StringUtils.isNotBlank(title)){
                    title = title.trim();
                }
                String content = cells.get(1).getText();
                if(StringUtils.isNotBlank(content)){
                    content = content.trim();
                }
                map.put(title, content);

            }
            list.add(map);
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            document.close();
            in.close();
        }
        return list;
    }
}
