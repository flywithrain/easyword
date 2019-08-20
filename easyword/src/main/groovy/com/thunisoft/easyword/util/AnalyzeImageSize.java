package com.thunisoft.easyword.util;

import org.jetbrains.annotations.NotNull;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * 2019/8/20 12:05
 *
 * @author 657518680@qq.com
 * @since beta
 * @version beta
 */
public final class AnalyzeImageSize {

    private AnalyzeImageSize() {

    }

    /**
     * 2019/8/20 14:28
     * get the original size in pixels of the picture
     *
     * @return the width and height of image
     * @param inputStream the inputStream of image
     * @author 657518680@qq.com
     * @since beta
     */
    public static Map<String, Integer> getImageSize(@NotNull InputStream inputStream) throws IOException {
        Map<String, Integer> result = new HashMap<>(2);
        BufferedImage bufferedImage = ImageIO.read(inputStream);
        result.put("width", bufferedImage.getWidth());
        result.put("height", bufferedImage.getHeight());
        return result;
    }

}
