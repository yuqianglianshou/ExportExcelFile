package com.moon.exportexcelfile

import android.util.Log
import jxl.Workbook
import jxl.write.Label
import jxl.write.WritableCellFormat
import jxl.write.WritableFont
import java.io.File

/**
 *
 *@author : lq
 *@date   : 2021/6/17
 *@desc   : 在指定全路径的 Excel表中写入数据
 *
 */
object ExcelUtils {
    //表头样式
    private lateinit var tableFormat: WritableCellFormat
    init {
        Log.i("lq ", "ExcelUtils init")
        format()
    }
    /**
     * excel 样式设置
     */
    private fun format() {
        val tableFont1 = WritableFont(WritableFont.ARIAL, 11, WritableFont.BOLD)//11号字，加粗
//        tableFont1.colour = jxl.format.Colour.LIGHT_BLUE // 可设置颜色

        tableFormat = WritableCellFormat(tableFont1)
        tableFormat.alignment = jxl.format.Alignment.CENTRE // 表格数据居中显示
        tableFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN)
        tableFormat.setBackground(jxl.format.Colour.GRAY_25)//表格添加背景颜色
    }

    /**
     * 写入数据到Excel中
     * file  所创建的Excel文件  /storage/emulated/0/Android/data/com.moon.exportexcelfile/files/Documents/export.xls
     * list  数据集合 j 行 i 列
     */
    fun writeDataToExcel(file: File, list: ArrayList<ArrayList<String>>) {

        if (list.isEmpty() || list.size == 0) {
            return
        }
        if (file.exists()) {
            Log.i("lq ", " 文件已经存在，删除重建")
            file.delete()
        }
        file.createNewFile()
        var workbook = Workbook.createWorkbook(file)
        try {
            //一张 execl 表可以创建很多表单（sheet），默认有3个，这里使用第一张表单创建数据
            val sheet = workbook.createSheet("表单1", 0)

            for (j in list.indices) {

                var listRow = list[j]

                for (i in listRow.indices) {
                    if (j == 0) {
                        //为表头添加样式
                        sheet.addCell(Label(i, j, listRow[i], tableFormat))
                    } else {
                        sheet.addCell(Label(i, j, listRow[i]))
                    }

                    //根据数据长度设置合适列宽
                    if (listRow[i].length <= 5) {
                        sheet.setColumnView(i, listRow[i].length + 8)
                    } else {
                        sheet.setColumnView(i, listRow[i].length + 5)
                    }
                }

                //设置行高
                sheet.setRowView(j, 350)

            }
            workbook.write()
        } catch (e: Exception) {
            e.printStackTrace()
        } finally {
            workbook?.close()
        }
    }

}