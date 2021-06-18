package com.moon.exportexcelfile

import android.os.Bundle
import android.os.Environment
import android.util.Log
import android.widget.Button
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import java.io.File
import java.util.*

class MainActivity : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        //显示存储路径
        val textView = findViewById<TextView>(R.id.textView)
        //导出按键
        findViewById<Button>(R.id.btn).setOnClickListener {
            val file = creatFile()
            ExcelUtils.writeDataToExcel(file, getData())
            Toast.makeText(this, "创建成功", Toast.LENGTH_SHORT).show()
            Log.i("lq ", " ok")
            textView.text = "文件目录为  " + file.absolutePath
        }
    }
    /**
     * 创建一个本地存储路径
     * /storage/emulated/0/Android/data/com.moon.exportexcelfile/files/Documents/export.xls
     */
    private fun creatFile(): File {
        val file =
            File(this.getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS)?.absolutePath + File.separator + "export.xls")

        Log.i("lq ", " file == " + file.absolutePath)
        return file
    }
    /**
     * 生成 20行8列 数据，第一行为表头
     */
    private fun getData(): ArrayList<ArrayList<String>> {
        val list = ArrayList<ArrayList<String>>()
        for (j in 1..20) {
            val listRow = ArrayList<String>()
            for (i in 1..8) {
                if (j == 1) {
                    //添加表头
                    listRow.add("表头 $i ")
                } else {
                    listRow.add("$j 行-$i 列")
                }
            }
            list.add(listRow)
        }
        Log.i("lq ", " list == $list")
        return list
    }

}