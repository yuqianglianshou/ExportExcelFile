package com.moon.exportexcelfile

import jxl.write.WritableCellFormat
import jxl.write.WritableFont

/**
 *
 *@author : lq
 *@date   : 2021/6/17
 *@desc   :
 *
 */
object ExcelUtils {
    init {
        format()
    }

    /**
     * excel 样式设置
     */
    private fun format() {
        //表头样式
        val tableFont1 = WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD)
        tableFont1.colour = jxl.format.Colour.LIGHT_BLUE

        val tableFormat = WritableCellFormat(tableFont1)
        tableFormat.alignment = jxl.format.Alignment.CENTRE
        tableFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN)
        tableFormat.setBackground(jxl.format.Colour.VERY_LIGHT_YELLOW)


    }
}