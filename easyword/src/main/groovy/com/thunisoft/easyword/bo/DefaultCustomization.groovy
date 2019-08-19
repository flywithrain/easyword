package com.thunisoft.easyword.bo
/**
 * 2019/8/13 9:35
 * DefaultCustomization
 *
 * @author 657518680@qq.com
 * @since 1.0.0
 */
class DefaultCustomization implements Customization {

    private String text = ''
    private InputStream picture
    private String pictureName
    private int width
    private int height

    DefaultCustomization() {

    }

    DefaultCustomization(String text) {
        this.text = text
    }

    void setText(String text) {
        if (text == null) {
            this.text = ''
        } else {
            this.text = text
        }
    }

    String getText() {
        return text
    }

    InputStream getPicture() {
        return picture
    }

    String getPictureName() {
        return pictureName
    }

    int getWidth() {
        return width
    }

    int getHeight() {
        return height
    }

}
