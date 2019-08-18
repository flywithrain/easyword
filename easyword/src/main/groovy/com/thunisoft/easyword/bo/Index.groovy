package com.thunisoft.easyword.bo
/**
 * @author 65751* @date 2019-08-2019/8/18 16:51
 */
class Index {

    Index(int pIndex, int rIndex) {
        this(0, 0, 0, pIndex, rIndex)
    }

    Index(int tableIndex, int rowIndex, int cIndex, int pIndex, int rIndex) {
        this.tableIndex = tableIndex
        this.rowIndex = rowIndex
        this.cIndex = cIndex
        this.pIndex = pIndex
        this.rIndex = rIndex
    }

    int tableIndex
    int rowIndex
    int cIndex
    int pIndex
    int rIndex
}
