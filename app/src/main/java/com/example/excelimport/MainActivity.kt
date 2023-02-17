package com.example.excelimport

import android.app.Activity
import android.content.ContentValues.TAG
import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import android.content.Intent
import android.net.Uri
import android.util.Log

import java.io.InputStream
import jxl.WorkbookSettings
import jxl.read.biff.BiffException



class MainActivity : AppCompatActivity() {
    private val REQUEST_CODE = 42
    private var dataList: MutableList<List<String>> = mutableListOf()

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        pickFile()
    }

    private fun pickFile() {
        val intent = Intent(Intent.ACTION_GET_CONTENT)
        intent.type = "application/vnd.ms-excel"
        startActivityForResult(intent, REQUEST_CODE)

    }

    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {

        var arraylist = ArrayList<String>()
        super.onActivityResult(requestCode, resultCode, data)
        if (requestCode == REQUEST_CODE && resultCode == Activity.RESULT_OK) {
            val uri = data?.data
            val inputStream = uri?.let { contentResolver.openInputStream(it) }
            val workbook: jxl.Workbook
            try {
                workbook = jxl.Workbook.getWorkbook(inputStream)
                Log.e(TAG, "Workbook received")
            } catch (e: BiffException) {
                e.printStackTrace()
                return
            }
            for (i in 0 until workbook.numberOfSheets) {
                val sheet = workbook.getSheet(i)
                println("Sheet $i name: ${sheet.name}")
                if (sheet == null) {
                    Log.e(TAG, "Sheet object is null")
                } else {
                    for (i in 0 until sheet.rows) {
                        for (j in 0 until sheet.columns) {
                            val cell = sheet.getCell(j, i)
                           // println("Cell $i,$j value: ${cell.contents}")
                            arraylist.add(cell.contents)
                            Log.e(TAG, "Stored in arrayList")

                        }
                    }
                }
            }
            Log.e(TAG, "Printing arraylist")

            for (element in arraylist) {
                println(element)
            }

          //  val sheet = workbook.getSheet(1)

            workbook.close()
        }
    }

}
