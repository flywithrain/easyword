package com.thunisoft.easyword.bo

import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.poi.xwpf.usermodel.XWPFRun
import org.apache.poi.xwpf.usermodel.XWPFTable
import org.apache.poi.xwpf.usermodel.XWPFTableCell
import org.apache.poi.xwpf.usermodel.XWPFTableRow

/**
 * @author 65751* @date 2019-08-2019/8/18 18:10
 */
class WordConstruct {

    WordConstruct(XWPFParagraph paragraph, XWPFRun run) {
        this(null, null, null, paragraph, run)
    }

    WordConstruct(XWPFTable table, XWPFTableRow row, XWPFTableCell cell, XWPFParagraph paragraph, XWPFRun run) {
        this.table = table
        this.row = row
        this.cell = cell
        this.paragraph = paragraph
        this.run = run
    }

    XWPFTable table
    XWPFTableRow row
    XWPFTableCell cell
    XWPFParagraph paragraph
    XWPFRun run
}
