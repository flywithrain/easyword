package com.thunisoft.easyword.bo

import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.poi.xwpf.usermodel.XWPFRun
import org.apache.poi.xwpf.usermodel.XWPFTable
import org.apache.poi.xwpf.usermodel.XWPFTableCell
import org.apache.poi.xwpf.usermodel.XWPFTableRow

/**
 * @author 65751* @date 2019-08-2019/8/18 18:10
 * @version 2.0.0
 */
public class WordConstruct {

    WordConstruct(XWPFDocument document, XWPFParagraph paragraph, XWPFRun run) {
        this(document, null, null, null, paragraph, run)
    }

    WordConstruct(XWPFDocument document, XWPFTable table,
                  XWPFTableRow row, XWPFTableCell cell,
                  XWPFParagraph paragraph, XWPFRun run) {
        this.table = table
        this.row = row
        this.cell = cell
        this.paragraph = paragraph
        this.run = run
        this.document = document
    }

    XWPFDocument document
    XWPFTable table
    XWPFTableRow row
    XWPFTableCell cell
    XWPFParagraph paragraph
    XWPFRun run

}
