package com.moon.exportexcelfile

import android.content.Context
import android.content.Intent
import android.net.Uri
import android.os.Build
import android.util.Log
import android.widget.Toast
import androidx.core.content.FileProvider
import java.io.File
import java.util.*

/**
 *
 *@author : lq
 *@date   : 5/10/21
 *@desc   : 文件打开工具类
 *
 */
class FileUtils {

    private var mContext: Context

    constructor(mContext: Context) {
        this.mContext = mContext
    }

    /**声明各种类型文件的dataType */
    private val DATA_TYPE_ALL = "*/*" //未指定明确的文件类型，不能使用精确类型的工具打开，需要用户选择

    private val DATA_TYPE_APK = "application/vnd.android.package-archive"
    private val DATA_TYPE_VIDEO = "video/*"
    private val DATA_TYPE_AUDIO = "audio/*"
    private val DATA_TYPE_HTML = "text/html"
    private val DATA_TYPE_IMAGE = "image/*"
    private val DATA_TYPE_PPT = "application/vnd.ms-powerpoint"
    private val DATA_TYPE_EXCEL = "application/vnd.ms-excel"
    private val DATA_TYPE_WORD = "application/msword"
    private val DATA_TYPE_CHM = "application/x-chm"
    private val DATA_TYPE_TXT = "text/plain"
    private val DATA_TYPE_PDF = "application/pdf"

    /**
     * 产生除了视频、音频、网页文件外，打开其他类型文件的Intent
     * @param filePath 文件路径
     * @param dataType 文件类型
     * @return
     */
    private fun generateCommonIntent(
        filePath: String,
        dataType: String
    ): Intent? {
        val intent = Intent()
        intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
        intent.setAction(Intent.ACTION_VIEW)
        val file = File(filePath)
        val uri = getUri(intent, file)
        intent.setDataAndType(uri, dataType)
        return intent
    }


    /**
     * 产生打开视频或音频的Intent
     * @param filePath 文件路径
     * @param dataType 文件类型
     * @return
     */
    private fun generateVideoAudioIntent(
        filePath: String,
        dataType: String
    ): Intent? {
        val intent =
            Intent(Intent.ACTION_VIEW)
        intent.addFlags(Intent.FLAG_ACTIVITY_CLEAR_TOP)
        intent.putExtra("oneshot", 0)
        intent.putExtra("configchange", 0)
        val file = File(filePath)
        intent.setDataAndType(getUri(intent, file), dataType)
        return intent
    }

    /**
     * 产生打开网页文件的Intent
     * @param filePath 文件路径
     * @return
     */
    private fun generateHtmlFileIntent(filePath: String): Intent? {
        val uri = Uri.parse(filePath)
            .buildUpon()
            .encodedAuthority("com.android.htmlfileprovider")
            .scheme("content")
            .encodedPath(filePath)
            .build()
        val intent =
            Intent(Intent.ACTION_VIEW)
        intent.setDataAndType(uri, DATA_TYPE_HTML)
        return intent
    }

    /**
     * 获取对应文件的Uri
     * @param intent 相应的Intent
     * @param file 文件对象
     * @return
     */
    private fun getUri(intent: Intent, file: File): Uri? {
        var uri: Uri? = null
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.N) {
            //判断版本是否在7.0以上
            uri = FileProvider.getUriForFile(
                mContext, mContext.getPackageName().toString() + ".fileprovider",
                file
            )
            //添加这一句表示对目标应用临时授权该Uri所代表的文件
            intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        } else {
            uri = Uri.fromFile(file)
        }
        return uri
    }

    /**
     * 打开文件
     * @param filePath 文件的全路径，包括到文件名
     */
    fun openFile(filePath: String) {
        Log.i("lq", "打开文件 == $filePath")
        val file = File(filePath)
        if (!file.exists()) {
            //如果文件不存在
            Toast.makeText(mContext, "文件打开失败，文件已经被移动或者删除", Toast.LENGTH_SHORT).show()
            return
        }
        /* 取得扩展名 */
        val end =
            file.name.substring(file.name.lastIndexOf(".") + 1, file.name.length)
                .toLowerCase(Locale.getDefault())
        /* 依扩展名的类型决定MimeType */
        var intent: Intent? = null
        if (end == "m4a" || end == "mp3" || end == "mid" || end == "xmf" || end == "ogg" || end == "wav") {
            intent = generateVideoAudioIntent(filePath, DATA_TYPE_AUDIO)
        } else if (end == "3gp" || end == "mp4") {
            intent = generateVideoAudioIntent(filePath, DATA_TYPE_VIDEO)
        } else if (end == "jpg" || end == "gif" || end == "png" || end == "jpeg" || end == "bmp") {
            intent = generateCommonIntent(filePath, DATA_TYPE_IMAGE)
        } else if (end == "apk") {
            intent = generateCommonIntent(filePath, DATA_TYPE_APK)
        } else if (end == "html" || end == "htm") {
            intent = generateHtmlFileIntent(filePath)
        } else if (end == "ppt") {
            intent = generateCommonIntent(filePath, DATA_TYPE_PPT)
        } else if (end == "xls" || end == "xlsx") {
            intent = generateCommonIntent(filePath, DATA_TYPE_EXCEL)
        } else if (end == "doc" || end == "docx") {
            intent = generateCommonIntent(filePath, DATA_TYPE_WORD)
        } else if (end == "pdf") {
            intent = generateCommonIntent(filePath, DATA_TYPE_PDF)
        } else if (end == "chm") {
            intent = generateCommonIntent(filePath, DATA_TYPE_CHM)
        } else if (end == "txt") {
            intent = generateCommonIntent(filePath, DATA_TYPE_TXT)
        } else {
            intent = generateCommonIntent(filePath, DATA_TYPE_ALL)
        }
        try {
            mContext.startActivity(intent)
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(mContext, "未找到可以打开文件的软件", Toast.LENGTH_SHORT).show()
        }

    }
}