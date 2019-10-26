package com.thunisoft.easyword.bo;

import java.io.InputStream;

/**
 * 2019/10/26 13:40
 *
 * @author wangxiaoyu 657518680@qq.com
 * @since 1.2.5
 */
public interface Customization4Picture extends Customization {

    /**
     * 2019/8/19
     * get the inputStream of the picture in EasyWord
     *
     * @return the inputStream of the picture
     * @author 657518680@qq.com
     * @since alpha
     */
    InputStream getPicture();

    /**
     * 2019/8/19
     * get the name of the picture inputStream in EasyWord
     *
     * @return the name of the picture
     * @author 657518680@qq.com
     * @since alpha
     */
    String getPictureName();

    /**
     * 2019/8/19
     * get the width in pixel of the picture inputStream in EasyWord
     * if <=0 will use the native size of the image both height {@link Customization4Text#getHeight()}
     *
     * @return the width of the picture
     * @author 657518680@qq.com
     * @since alpha
     */
    int getWidth();

    /**
     * 2019/8/19
     * get the height in pixel of the picture inputStream in EasyWord
     * if <=0 will use the native size of the image both width {@link Customization4Text#getWidth()}
     *
     * @return the height of the picture
     * @author 657518680@qq.com
     * @since alpha
     */
    int getHeight();

}
