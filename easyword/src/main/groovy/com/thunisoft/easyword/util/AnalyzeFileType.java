package com.thunisoft.easyword.util;

import com.thunisoft.easyword.constant.FileTypeEnum;
import org.jetbrains.annotations.NotNull;

import java.util.Locale;

/**
 * 2019/8/13 14:59
 * analyzeFileType
 *
 * @author 657518680@qq.com
 * @since alpha
 * @version beta
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
     * @param bytes the byte array of the picture
     * @return -1:is not image; other:the type of the picture{@link FileTypeEnum}
     * @author 657518680@qq.com
     * @since alpha
     */
    public static int getFileType(@NotNull byte[] bytes) {
        int i = 0;
        StringBuilder builder = new StringBuilder();
        while (i < MAX_LENGTH && i < bytes.length) {
            builder.append(String.format("%02X", bytes[i]));
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
