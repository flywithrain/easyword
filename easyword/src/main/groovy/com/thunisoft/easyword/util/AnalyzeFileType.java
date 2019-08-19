package com.thunisoft.easyword.util;

import com.thunisoft.easyword.constant.FileTypeEnum;
import org.jetbrains.annotations.NotNull;

import java.io.IOException;
import java.io.InputStream;
import java.util.Locale;

/**
 * 2019/8/13 14:59
 * analyzeFileType
 *
 * @author 657518680@qq.com
 * @since 1.0.0
 */
public final class AnalyzeFileType {

    /**
     * the max length of hex code
     */
    private static final int MAX_LENGTH = 14;

    private AnalyzeFileType() {
    }

    /**
     * 2019/8/19
     * analyze the type of image through the file header
     *
     * @param inputStream the inputStream of the picture
     * @return -1:is not image; other:the type of the picture{@link FileTypeEnum}
     * @throws IOException IOException
     * @author 657518680@qq.com
     * @since 1.0.0
     */
    public static int getFileType(@NotNull InputStream inputStream) throws IOException {
        int temp;
        int i = 0;
        StringBuilder builder = new StringBuilder();
        while ((temp = inputStream.read()) != -1 && i < MAX_LENGTH) {
            builder.append(String.format("%02X", temp));
            i++;
        }
        String hex = builder.toString().toUpperCase(Locale.ENGLISH);
        for (FileTypeEnum typeEnum : FileTypeEnum.values()) {
            if (hex.startsWith(typeEnum.getHex())) {
                return typeEnum.getValue();
            }
        }
        return -1;
    }

}
