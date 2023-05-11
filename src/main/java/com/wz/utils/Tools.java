package com.wz.utils;

import java.io.*;

import cn.hutool.core.codec.Base64;

/**
 * @author 王哲
 * @ClassName Tools
 * @create 2023--五月--下午10:28
 * @Description Tools
 * @Version V1.0
 */
public class Tools {
    /**
     * 获得指定图片文件的base64编码数据
     * @param filePath 文件路径
     * @return base64编码数据
     */
    public static String getBase64ByPath(String filePath) {
        if(!hasLength(filePath)){
            return "";
        }
        File file = new File(filePath);
        if(!file.exists()) {
            return "";
        }
        InputStream in = null;
        byte[] data = null;
        try {
            in = new FileInputStream(file);
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        }
        try {
            assert in != null;
            data = new byte[in.available()];
            in.read(data);
            in.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        Base64 encoder = new Base64();
        return encoder.encode(data);
    }
    /**
     * @desc 判断字符串是否有长度
     */
    public static boolean hasLength(String str) {
        return org.springframework.util.StringUtils.hasLength(str);
    }
}
