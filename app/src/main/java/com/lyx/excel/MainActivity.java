package com.lyx.excel;

import android.Manifest;
import android.annotation.SuppressLint;
import android.os.Environment;
import android.os.Handler;
import android.os.Message;
import android.support.annotation.NonNull;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.widget.Toast;

import com.lyx.excel.permission.Permission;
import com.lyx.excel.permission.PermissionManager;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class MainActivity extends AppCompatActivity {
    private String regex = "[\u4e00-\u9fa5]{1}[A-Z]{1}[A-Z_0-9]{5}";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
    }

    @SuppressLint("HandlerLeak")
    private Handler handler = new Handler() {
        @Override
        public void handleMessage(Message msg) {
            super.handleMessage(msg);
            String message = (String) msg.obj;

        }
    };

    private PermissionManager permissionManager;

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        permissionManager.onPermissionsResult(requestCode, permissions, grantResults);
    }

    @Override
    protected void onResume() {
        super.onResume();
        scheduleTask();
    }

    private void scheduleTask() {
        if (PermissionManager.checkPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE)) {
            new Thread(new Runnable() {
                @Override
                public void run() {
                    convert();
                }
            }).start();
        } else {
            permissionManager = new PermissionManager(this);
            permissionManager.addPermission(new Permission() {
                @Override
                public String getPermission() {
                    return Manifest.permission.WRITE_EXTERNAL_STORAGE;
                }

                @Override
                public void onApplyResult(boolean succeed) {
                    if (succeed) {
                        new Thread(new Runnable() {
                            @Override
                            public void run() {
                                convert();
                            }
                        }).start();
                    } else {
                        Toast.makeText(MainActivity.this, "没有读写文件权限无法操作！", Toast.LENGTH_SHORT).show();
                    }
                }
            }).apply(this);
        }
    }

    private String readFileDir = "abc";
    private String readFileName = "excel.xls";

    private void convert() {
        List<ExcelCell> excelList = new ArrayList<>();

        try {
            String path = Environment.getExternalStorageDirectory().getAbsolutePath();
            File file = new File(path + "/" + readFileDir + "/" + readFileName);

            Workbook book = Workbook.getWorkbook(file);

            Sheet[] sheets = book.getSheets();
            if (sheets.length == 0) {
                return;
            }

            for (Sheet sheet : sheets) {
                ExcelCell excelCell = new ExcelCell();
                excelCell.tableName = sheet.getName();

                List<Info> list = new ArrayList<>();

                int row = sheet.getRows();

                for (int i = 1; i < row; i++) {  //i = 1   从2行开始读取
                    Cell cell = sheet.getCell(2, i); //（列，行）2--> 读取第三列（即C列）
                    String data = cell.getContents();
                    String[] content = data.split("装：");

                    Info inf = new Info();
                    if (content.length > 1) {
//                    Log.d("MainActivity", "content: " + content[0]);

                        String[] info = content[0].split("，|,| ", 3);  //分为3个数组

                        if (info.length == 3) {
                            String[] pStr = info[0].split("[\\d]+");
                            if (pStr.length > 1) {
                                inf.provider = pStr[1].substring(1);
                            } else {
                                inf.provider = info[0];
                            }

                            inf.card = info[1];


                            //找出车牌和个人信息
                            String str = info[2];
                            Matcher m = Pattern.compile(regex).matcher(str);

                            if (m.find()) {
                                String number = m.group();
                                inf.number = number;

                                String information = null;

                                int index = str.indexOf(number);
                                if (index == 0) {
                                    information = str.substring(number.length(), str.length());
                                } else if (index == str.length() - number.length()) {
                                    information = str.substring(0, str.length() - number.length());
                                } else {
                                    information = str.substring(0, index) + str.substring(index + number.length(), str.length());
                                }

                                information = information.replaceAll("，", " ");
                                information = information.replaceAll(",", " ");
                                information = information.replaceAll("；", " ");
                                information = information.replaceAll(" ", "");

                                inf.information = information.trim();
                            }

                        }
                    }

                    Log.i("ExcelActivity", "" + inf.toString());
                    list.add(inf);
                }

                excelCell.list = list;
                excelList.add(excelCell);
            }
            book.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        Log.w("ExcelActivity", "============== 读取完毕 ===============");

        WritableSheet mWritableSheet;
        WritableWorkbook mWritableWorkbook;

        String path = Environment.getExternalStorageDirectory().getAbsolutePath();

        try {
            for (int j = 0; j < excelList.size(); j++) {
                ExcelCell excelCell = excelList.get(j);
                List<Info> list = excelCell.list;

                // 输出Excel的路径
                String filePath = path + "/" + readFileDir + "/excel_" + excelCell.tableName + ".xls";
                // 新建一个文件
                OutputStream os = new FileOutputStream(filePath);
                // 创建Excel工作簿
                mWritableWorkbook = Workbook.createWorkbook(os);
                // 创建Sheet表
                mWritableSheet = mWritableWorkbook.createSheet(excelCell.tableName, 0);

                if (null == mWritableSheet) {
                    return;
                }

                for (int i = 0; i < list.size(); i++) {
                    Info info = list.get(i);
                    Log.d("ExcelActivity", "" + info.toString());

                    mWritableSheet.addCell(new Label(3, i, info.provider));
                    mWritableSheet.addCell(new Label(4, i, info.card));
                    mWritableSheet.addCell(new Label(5, i, info.number));
                    mWritableSheet.addCell(new Label(6, i, info.information));

                }

                // 写入数据
                mWritableWorkbook.write();
                // 关闭文件
                mWritableWorkbook.close();
            }
        } catch (WriteException | IOException e) {
            e.printStackTrace();
        }

        Log.w("ExcelActivity", "==============写入完毕 ===============");
    }

}
