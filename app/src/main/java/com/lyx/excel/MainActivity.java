package com.lyx.excel;

import android.Manifest;
import android.annotation.SuppressLint;
import android.content.DialogInterface;
import android.os.Environment;
import android.os.Handler;
import android.os.Message;
import android.support.annotation.NonNull;
import android.support.v7.app.AlertDialog;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;
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
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class MainActivity extends AppCompatActivity {
    private TextView noticeTV;
    private TextView tipTV;
    private Button convertBtn;

    private String regex = "[\u4e00-\u9fa5]{1}[A-Z]{1}[A-Z_0-9]{5}";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        noticeTV = findViewById(R.id.tv_notice);
        tipTV = findViewById(R.id.tv_tip);
        convertBtn = findViewById(R.id.btn_convert);

        convertBtn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                tipTV.setText(null);
                scheduleTask();
            }
        });

        noticeTV.append("要求：\n1、必须将文档另保存为.xls格式，不能是.xlsx格式；" +
                "\n2、文档名称必须为：excel.xls；" +
                "\n3、需要在手机存储根目录新建一个文件夹，命名为abc；" +
                "\n4、将excel.xls文档复制到刚才创建的abc目录下面；" +
                "\n5、以上操作无误后即可进行转换操作。");

    }

    @SuppressLint("HandlerLeak")
    private Handler handler = new Handler() {
        @Override
        public void handleMessage(Message msg) {
            super.handleMessage(msg);

            String message = (String) msg.obj;
            tipTV.append(message);

            if (msg.what == 101) {
                new AlertDialog.Builder(MainActivity.this)
                        .setTitle("提示：转换完成")
                        .setMessage("文档转换成功！")
                        .setPositiveButton("确定", new DialogInterface.OnClickListener() {
                            @Override
                            public void onClick(DialogInterface dialog, int which) {
                                dialog.dismiss();
                            }
                        })
                        .create()
                        .show();
            }
        }
    };

    private PermissionManager permissionManager;

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        permissionManager.onPermissionsResult(requestCode, permissions, grantResults);
    }

    private void scheduleTask() {
        tipTV.append("请求转换，检查读写权限\n");
        if (PermissionManager.checkPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE)) {
            tipTV.append("已有读写权限，即将开始转换任务\n");
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
                        tipTV.append("已有读写权限，即将开始转换任务\n");
                        new Thread(new Runnable() {
                            @Override
                            public void run() {
                                convert();
                            }
                        }).start();
                    } else {
                        tipTV.append("没有读写权限，转换失败！\n");
                        Toast.makeText(MainActivity.this, "没有读写文件权限无法操作！", Toast.LENGTH_SHORT).show();
                    }
                }
            }).apply(this);
        }
    }

    private String readFileDir = "abc";
    private String readFileName = "excel.xls";

    private void convert() {
        sendMessage("开始检查转换文件...\n");

        List<ExcelCell> excelList = new ArrayList<>();

        try {
            String path = Environment.getExternalStorageDirectory().getAbsolutePath();
            File file = new File(path + "/" + readFileDir + "/" + readFileName);

            if (!file.exists()) {
                sendMessage(readFileDir + "目录下的" + readFileName + "文件不存在\n");
                return;
            }

            Workbook book = Workbook.getWorkbook(file);

            Sheet[] sheets = book.getSheets();
            if (sheets.length == 0) {
                sendMessage("文件没有工作表可操作\n");
                return;
            }

            Message m = Message.obtain();
            m.obj = "读取文件中...\n";
            handler.sendMessage(m);

            for (Sheet sheet : sheets) {
                ExcelCell excelCell = new ExcelCell();
                excelCell.tableName = sheet.getName();

                sendMessage("正在读取工作表：" + excelCell.tableName + "...\n");

                List<Info> list = new ArrayList<>();

                int row = sheet.getRows();

                for (int i = 1; i < row; i++) {  //i = 1   从2行开始读取
                    Cell cell = sheet.getCell(2, i); //（列，行）2--> 读取第三列（即C列）
                    String data = cell.getContents();
                    String[] content = data.split("装");

                    Info inf = new Info();
                    if (content.length > 1) {
//                    Log.d("MainActivity", "content: " + content[0]);

                        String[] info = content[0].split("，|,|。", 3);  //分为3个数组

                        if (info.length == 3) {
                            String[] pStr = info[0].split("[\\d]+");
                            if (pStr.length > 1) {
                                Matcher matcher = Pattern.compile("[\\d]+").matcher(info[0]);
                                if (matcher.find()) {
                                    inf.row = matcher.group();
                                }

                                inf.provider = pStr[1].substring(1);
                            } else {
                                inf.provider = info[0];
                            }

                            inf.card = info[1];


                            //找出车牌和个人信息
                            String str = info[2];
                            Matcher matcher = Pattern.compile(regex).matcher(str);

                            if (matcher.find()) {
                                String number = matcher.group();
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

                sendMessage("读取完工作表：" + excelCell.tableName + "\n");
            }

            sendMessage("所有工作表已读取完，准备转换写入新文件...\n");

            book.close();

            Log.w("ExcelActivity", "============== 读取完毕 ===============");

            WritableSheet mWritableSheet;
            WritableWorkbook mWritableWorkbook;

            String path1 = Environment.getExternalStorageDirectory().getAbsolutePath();

//        try {
            for (int j = 0; j < excelList.size(); j++) {
                ExcelCell excelCell = excelList.get(j);
                List<Info> list = excelCell.list;

                sendMessage("正在创建工作表：" + excelCell.tableName + "的副文件...\n");

                // 输出Excel的路径
                String filePath = path1 + "/" + readFileDir + "/excel_" + excelCell.tableName + ".xls";
                // 新建一个文件
                OutputStream os = new FileOutputStream(filePath);
                // 创建Excel工作簿
                mWritableWorkbook = Workbook.createWorkbook(os);
                // 创建Sheet表
                mWritableSheet = mWritableWorkbook.createSheet(excelCell.tableName, 0);

                if (null == mWritableSheet) {
                    return;
                }

                sendMessage("已创建工作表：excel_" + excelCell.tableName + ".xls，开始写入...\n");

                for (int i = 0; i < list.size(); i++) {
                    Info info = list.get(i);
                    Log.d("ExcelActivity", "" + info.toString());

                    mWritableSheet.addCell(new Label(2, i, info.row + "、"));
                    mWritableSheet.addCell(new Label(3, i, info.provider));
                    mWritableSheet.addCell(new Label(4, i, info.card));
                    mWritableSheet.addCell(new Label(5, i, info.number));
                    mWritableSheet.addCell(new Label(6, i, info.information));

                }

                sendMessage("工作表：excel_" + excelCell.tableName + ".xls写入成功！\n");

                // 写入数据
                mWritableWorkbook.write();
                // 关闭文件
                mWritableWorkbook.close();
            }

            sendMessage("所有工作表写入成功！\n");

            sendMessage("\n ================ 转换完成 ================\n", 101);

        } catch (WriteException | IOException | BiffException e) {
            e.printStackTrace();
            sendMessage(e.getMessage() + "\n" + e.toString() + "\n转换失败\n");
        }

        Log.w("ExcelActivity", "==============写入完毕 ===============");
    }

    private void sendMessage(String string) {
        sendMessage(string, 0);
    }

    private void sendMessage(String string, int what) {
        Message msg = Message.obtain();
        msg.obj = string;
        msg.what = what;
        handler.sendMessage(msg);
    }
}
