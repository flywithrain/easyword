package com.thunisoft.easyword.bo
/**
 * 2019/8/13 9:35
 * DefaultCustomization
 *
 * @author 657518680@qq.com
 * @since alpha
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

    void setPicture(InputStream picture) {
        this.picture = picture
    }

    void setPictureName(String pictureName) {
        this.pictureName = pictureName
    }

    void setWidth(int width) {
        this.width = width
    }

    void setHeight(int height) {
        this.height = height
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
