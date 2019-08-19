package com.thunisoft.easyword.constant;

import org.apache.poi.xwpf.usermodel.Document;

/**
 * 2019/8/13 15:03
 * FileType
 *
 * @author 657518680@qq.com
 * @since 1.0.0
 */
public enum FileTypeEnum {

    /**
     * JPEG.
     */
    JPEG("FFD8FF", Document.PICTURE_TYPE_JPEG),

    /**
     * PNG.
     */
    PNG("89504E47", Document.PICTURE_TYPE_PNG),

    /**
     * GIF.
     */
    GIF("47494638", Document.PICTURE_TYPE_GIF),

    /**
     * TIFF.
     */
    TIFF("49492A00", Document.PICTURE_TYPE_TIFF),

    /**
     * Windows Bitmap.
     */
    BMP("424D", Document.PICTURE_TYPE_BMP),

    /**
      * Extended (Enhanced) Windows MetaFile Format, printer spool file
      */
    EMF("0100000058000000", Document.PICTURE_TYPE_EMF),

    /**
      * Graphics MetaFile
      */
    WMF("01000900", Document.PICTURE_TYPE_WMF),

    /**
      * device-independent bitmap image
      */
    DIB("424D", Document.PICTURE_TYPE_DIB),

    /**
     * Postscript.
     */
    EPS("252150532D41646F6265", Document.PICTURE_TYPE_EPS),

    /**
      * WordPerfect Graphics
      */
    WPG("FF575047", Document.PICTURE_TYPE_WPG),

    /**
     * CAD.
     */
    DWG("41433130"),

    /**
     * Adobe PhotoShop.
     */
    PSD("38425053"),

    /**
     * Rich Text Format.
     */
    RTF("7B5C727466"),

    /**
     * XML.
     */
    XML("3C3F786D6C"),

    /**
     * HTML.
     */
    HTML("68746D6C3E"),

    /**
     * Email [thorough only].
     */
    EML("44656C69766572792D646174653A"),

    /**
     * Outlook Express.
     */
    DBX("CFAD12FEC5FD746F"),

    /**
     * Outlook (pst).
     */
    PST("2142444E"),

    /**
     * MS Word/Excel.
     */
    XLS_DOC("D0CF11E0"),

    /**
     * MS Access.
     */
    MDB("5374616E64617264204A"),

    /**
     * WordPerfect.
     */
    WPD("FF575043"),

    /**
     * Adobe Acrobat.
     */
    PDF("255044462D312E"),

    /**
     * Quicken.
     */
    QDF("AC9EBD8F"),

    /**
     * Windows Password.
     */
    PWL("E3828596"),

    /**
     * ZIP Archive.
     */
    ZIP("504B0304"),

    /**
     * RAR Archive.
     */
    RAR("52617221"),

    /**
     * Wave.
     */
    WAV("57415645"),

    /**
     * AVI.
     */
    AVI("41564920"),

    /**
     * Real Audio.
     */
    RAM("2E7261FD"),

    /**
     * Real Media.
     */
    RM("2E524D46"),

    /**
     * MPEG (mpg).
     */
    MPG("000001BA"),

    /**
     * Quicktime.
     */
    MOV("6D6F6F76"),

    /**
     * Windows Media.
     */
    ASF("3026B2758E66CF11"),

    /**
     * MIDI.
     */
    MID("4D546864");

    private String hex;
    private int value;

    FileTypeEnum(String hex) {
        this(hex, -1);
    }

    FileTypeEnum(String hex, int value) {
        this.hex = hex;
        this.value = value;
    }

    public String getHex() {
        return hex;
    }

    public int getValue() {
        return value;
    }

}
