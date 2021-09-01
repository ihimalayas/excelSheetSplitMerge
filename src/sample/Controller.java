package sample;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/*
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
*/


public class Controller {

    @FXML
    private TextArea ProgramStatus;

    @FXML
    private Label MergeStatusInfo;

    @FXML
    private Button chooseFileButton;

    @FXML
    private Button chooseFolderButton;


    //--------------------------------------分割Excel--------------------------------------------------------------------
    @FXML
    void chooseFile(ActionEvent event) throws IOException {

        Date now = new Date();
        SimpleDateFormat programStatusTime = new SimpleDateFormat("HH:mm:ss");
        ProgramStatus.appendText("[" + programStatusTime.format(new Date()) + "]:" + "---------------------------------------------开始分割Excel------------------------------------\n");

        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("选择需要分割sheet的Excel");
        File filechooser = fileChooser.showOpenDialog(new Stage());
        File file = new File(filechooser.getPath());
//        判断选择的文件是否为excel
//        "xlsx", "xls", "xlsm", "xltx", "xlt", "xlam", "csv"
        List<String> excelTypes = Arrays.asList("xls");
        String filename = file.getName();
        int lastIndexOf = filename.lastIndexOf(".");
        String choosenFileSuffix = null;
        String choosenFileName = null;
        if (lastIndexOf == -1) {
            System.out.println("您选择的文件后缀为：" + choosenFileSuffix + "-------------");
        } else {
            choosenFileSuffix = filename.toLowerCase().substring(lastIndexOf + 1);
            choosenFileName = filename.substring(0, lastIndexOf);
            System.out.println("您选择的文件后缀为：" + choosenFileSuffix + "。");
            System.out.println("您选择的文件名称为：" + choosenFileName + "。");
            ProgramStatus.appendText("[" + programStatusTime.format(new Date()) + "]:" + "您选择的文件名称为：" + choosenFileName + "。\n");
            ProgramStatus.appendText("[" + programStatusTime.format(new Date()) + "]:" + "您选择的文件后缀为：" + choosenFileSuffix + "。\n");
        }
        String filePath = null;
        String fileFolderPath = null;
        if (excelTypes.contains(choosenFileSuffix)) {
            filePath = file.getPath();
            fileFolderPath = file.getParent();
            System.out.println(filePath + "\n" + fileFolderPath);
        } else {
            System.out.println("您选择的文件不是有效的excel文件。");
            ProgramStatus.appendText("[" + programStatusTime.format(new Date()) + "]:" + "您选择的文件不是有效的excel文件。" + "\n");
        }
        // 创建分割excel 各个sheet存放的路径，以excel名称命名文件夹的名称，放在同excel相同目录下面
        String fileOutPutPath = fileFolderPath + "\\" + choosenFileName;
        System.out.println("fileOutPutPath" + fileOutPutPath);
        File fileOutPutFolder = new File(fileOutPutPath);
        if (fileOutPutFolder.mkdir()) {
            System.out.println("fileOutPutFolder.getPath()" + fileOutPutFolder.getPath());
        } else {
            System.out.println(fileOutPutFolder.getPath() + "is exits~");
        }

        // beginning ______________________________________

        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(new File(filePath));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        Workbook workbook = null;
        try {
            workbook = Workbook.getWorkbook(fileInputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }


        //将workbook的多个sheet拆分为单个excel文件
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            FileOutputStream fileOutputStream = null;
            WritableWorkbook newworkbook = null;
            Sheet sh = workbook.getSheet(i);
            System.out.println(sh.getName() + " sheet另存为路径为" + fileOutPutFolder.getPath() + "\\" + choosenFileName + "-" + sh.getName() + ".xls");
            ProgramStatus.appendText("[" + programStatusTime.format(new Date()) + "]:【" + sh.getName() + "】sheet的另存为路径为" + fileOutPutFolder.getPath() + "\\" + sh.getName() + ".xls" + "\n");
            try {
                fileOutputStream = new FileOutputStream(new File(fileOutPutFolder.getPath() + "\\" + choosenFileName + "-" + sh.getName() + ".xls"));
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }

            try {
                newworkbook = Workbook.createWorkbook(fileOutputStream);
            } catch (IOException e) {
                e.printStackTrace();
            }



            newworkbook.importSheet(sh.getName(), 0, sh);
            try {
                newworkbook.write();
            } catch (IOException e) {
                e.printStackTrace();
                System.out.println("workbook  write  failed!");
                ProgramStatus.appendText("[" + programStatusTime.format(new Date()) + "]:" + "Excel 写入失败。" + "\n");
            }
            try {
                newworkbook.close();
                this.MergeStatusInfo.setText("恭喜你，已完成分割操作！");
                this.MergeStatusInfo.setStyle("-fx-text-fill:#0c2bda;");
                ProgramStatus.appendText("[" + programStatusTime.format(new Date()) + "]:" + "Excel 成功分割完毕。" + "\n");
            } catch (IOException e) {
                e.printStackTrace();
                System.out.println("workbook  close  failed!");
            } catch (WriteException e) {
                e.printStackTrace();
                System.out.println("workbook  close  failed!");
            }
        }
    }

    //--------------------------------------合并Excel--------------------------------------------------------------------
    @FXML
    void chooseFolder(ActionEvent event) {

        Date now = new Date();
        SimpleDateFormat ft = new SimpleDateFormat("yyyy-MM-dd-HHmmss");
        SimpleDateFormat programStatusTime = new SimpleDateFormat("HH:mm:ss");
        ProgramStatus.appendText("[" + programStatusTime.format(new Date()) + "]:" + "---------------------------------------------开始合并Excel------------------------------------\n");

        //选择需要合并的Excel所在的文件夹，仅合并每个excel的第一个sheet
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle("选择需要合并Excel(.xls,2003版本Excel)文件所在文件夹");
        File file = directoryChooser.showDialog(new Stage());
        File[] toMerge = file.listFiles();
//        String filePath = file.toString();

        // 在选择的文件夹外，即与文件夹同级别的路径下，创建名为汇总.xls的文件

        WritableWorkbook sumWorkbook = null;
//        FileOutputStream fileOutputStream = null;
//        try {
//            fileOutputStream = new FileOutputStream(new File(file.getParent() + "\\" + "汇总.xls"));
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        }

        String sumFileName = "汇总——" + ft.format(now) + ".xls";
        try {
            sumWorkbook = Workbook.createWorkbook(new File(file.getParent() + "\\" + sumFileName));
            ProgramStatus.appendText("[" + programStatusTime.format(new Date()) + "]:" + sumFileName + "创建成功！\n");
            int caretPosition = ProgramStatus.caretPositionProperty().get();
            ProgramStatus.appendText("Here i am appending text to text area" + "\n");
            ProgramStatus.positionCaret(caretPosition);
        } catch (IOException e) {
            e.printStackTrace();
        }

        int numberOfExcelFiles = getExcleFileNumber(file);
        System.out.println("文件夹内Excel的文件数量为： " + numberOfExcelFiles);

        for (int i = 0; i < numberOfExcelFiles; i++) {
            File f = toMerge[i];
            String filepathTamp = f.getPath();
            ProgramStatus.appendText("[" + programStatusTime.format(new Date()) + "]:" + "正在对第" + (i + 1) + "个文件:" + f.getName() + "进行合并······\n");
            if (f.isFile() && getFileSuffix(f.getName()).equals("xls")) {
                // 将要合并的excel转为输入流
                FileInputStream fileInputStream = null;
                try {
                    fileInputStream = new FileInputStream(new File(filepathTamp));
                    System.out.println("fileInputStream  success");
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                }

                // 将workbook获取Excel工作簿
                Workbook workbook = null;
                try {
                    workbook = Workbook.getWorkbook(fileInputStream);
                    System.out.println("workbook  success");
                } catch (IOException e) {
                    e.printStackTrace();
                } catch (BiffException e) {
                    e.printStackTrace();
                }

                //获取workbook工作簿待合并的sheet，保存为sh
//            int numberOfSheets = workbook.getNumberOfSheets();
//            if (numberOfSheets == 1) {
                Sheet sh = workbook.getSheet(0);
                //将sh保存到新的excel，名为：汇总.xls
                sumWorkbook.importSheet(sh.getName(), i, sh);

//            每个workbook只能写一次，否则excle很大但是无法展示
//            比如如果sumWorkbook.write()在这里写入则只会有一个sheet成功写入，其他的均无法成功。
//            try {
//                sumWorkbook.write();
//                System.out.println("sumWorkbook  write  success");
//            } catch (IOException e) {
//                e.printStackTrace();
//                System.out.println("workbook  write  failed!");
//            }

                workbook.close();
                ProgramStatus.appendText("[" + programStatusTime.format(new Date()) + "]:" + "第" + (i + 1) + "个文件:" + f.getName() + "的sheet" + sh.getName() + "成功！！！\n");
            } else {
                System.out.println(f.getName() + "有多个sheet，请确认，并保证只有一个sheet才能进行合并" + "\n");
                ProgramStatus.appendText("[" + programStatusTime.format(new Date()) + "]:" + f.getName() + "有多个sheet，请确认，并保证只有一个sheet才能进行合并" + "\n");
                int caretPosition = ProgramStatus.caretPositionProperty().get();
                ProgramStatus.appendText("Here i am appending text to text area" + "\n");
                ProgramStatus.positionCaret(caretPosition);
            }

        }


        // 将需要合并的sheet写入汇总工作簿
        try {
            sumWorkbook.write();
            System.out.println("需要合并的sheet成功写入汇总工作簿！");
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("workbook  write  failed!");
        }

        // 关闭汇总工作簿
        try {
            sumWorkbook.close();
            this.MergeStatusInfo.setText("恭喜你，已完成合并操作！");
            this.MergeStatusInfo.setStyle("-fx-text-fill:#9a0303;");
            ProgramStatus.appendText("[" + programStatusTime.format(new Date()) + "]:" + "恭喜你，文件夹内Excel已全部完成合并操作！！！\n");
//            TextArea 自动滚动
//            int caretPosition = ProgramStatus.caretPositionProperty().get();
//            ProgramStatus.appendText("Here i am appending text to text area"+"\n");
//            ProgramStatus.positionCaret(caretPosition);

            System.out.println("汇总工作簿成功关闭！");
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("workbook  close  failed!");
        } catch (WriteException e) {
            e.printStackTrace();
            System.out.println("workbook  close  failed!");
        }

    }

    //获取文件后缀名
    public String getFileSuffix(String filename) {
        //获取最后一个.的位置
        int lastIndexOf = filename.lastIndexOf(".");
        if (lastIndexOf == -1) {
            return null;
        }
        //获取文件的后缀名 .jpg
        String suffix = filename.substring(lastIndexOf + 1);
        return suffix;
    }

    public Integer getExcleFileNumber(File folderPath) {
        File[] listFiles = folderPath.listFiles();
        Integer excelFileNumber = 0;
        for (int i = 0; i < listFiles.length; i++) {
            //获取最后一个.的位置
            int lastIndexOf = listFiles[i].getName().lastIndexOf(".");
            if (lastIndexOf != -1) {
                //获取文件的后缀名 .jpg
                String suffix = listFiles[i].getName().substring(lastIndexOf + 1);
                if (suffix.equals("xls")) {
                    excelFileNumber += 1;
                }
            } else {
                System.out.println("请注意[ " + listFiles[i] + " ]没有有效的后缀。");
                continue;
            }
        }
        return excelFileNumber;
    }


}


//       下面的代码报错（jxl.common.AssertionFailed）：解决办法如上采用sheet 而非sheets
//        解决办法参照：http://www.myexceptions.net/ai/1131728.html
//                   https://blog.csdn.net/cheneyfeng3/article/details/6394325
//        Sheet[] sheets = workbook.getSheets();
//
//        for (Sheet sh : sheets) {
//            FileOutputStream fileOutputStream = null;
//            WritableWorkbook newworkbook = null;
//
//            System.out.print(sh.getName() + "\t");
//            System.out.println("sheet 另存为路径为"+ fileOutPutFolder.getPath()+sh.getName()+".xls");
//            try {
//                fileOutputStream = new FileOutputStream(new File(fileOutPutFolder.getPath()+"\\"+choosenFileName+"-"+sh.getName()+".xls"));
//                System.out.println("fileOutputStream creat  suceesss");
//            } catch (FileNotFoundException e) {
//                e.printStackTrace();
//            }
//
//            try {
//                newworkbook = Workbook.createWorkbook(fileOutputStream);
//                System.out.println("newworkbook creat success");
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//
//            newworkbook.importSheet("ss", 0, sh);
//            System.out.println("importSheet success");
//            try {
//                newworkbook.write();
//                System.out.println("write  success");
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//            try {
//                newworkbook.close();
//                System.out.println("close  success");
//            } catch (IOException e) {
//                e.printStackTrace();
//            } catch (WriteException e) {
//                e.printStackTrace();
//            }
//
//        }


//        beginning ending_-------------------------------------------------
//       来自 https://stackoverflow.com/questions/36785425/workbookfactory-createinputstream
//        File customerTemplateFileObj = new File("C:\\Users\\Magnum\\Downloads\\A.xls");
//        FileInputStream inputStream = null;
//        try {
//            inputStream = new FileInputStream(file);
//            System.out.println("inputStream:" + inputStream);
//        } catch (FileNotFoundException e) {
//            System.out.println("没有找到文件。");
//        }

//        try {
//
//            FileInputStream fileInputStream = new FileInputStream(file);
//            FileOutputStream fileOutputStream = new FileOutputStream(new File("d:/x.xls"));
//            Workbook workbook =  Workbook.getWorkbook(fileInputStream);
//
//            System.out.println("success");
//
//            Sheet[] sheets =workbook.getSheets();
//            WritableWorkbook newworkbook =  Workbook.createWorkbook(fileOutputStream);
//            System.out.println("sheets[0].getName()"+sheets[0].getName());
//            newworkbook.importSheet(sheets[0].getName(),0,sheets[0]);
//            newworkbook.write();
//            newworkbook.close();
//
//            int i = 0;
//            for (Sheet sheet:sheets ) {
//                String sheetName = sheet.getName();
//                System.out.println(sheetName);

//                File splitXlsFile = new File(fileFolderPath+"\\拆分文件汇总");
//                System.out.println("getPath:"+splitXlsFile.getPath());
//                System.out.println("getname:"+splitXlsFile.getName());
//                splitXlsFile.mkdir();
//                if(splitXlsFile.exists()){
//                    splitXlsFile.delete();
//                }else{
//                    splitXlsFile.mkdir();
//                    System.out.println("make dir success");
//                }
//                WritableWorkbook newworkbook =  Workbook.createWorkbook(new File("D:\\xx.xls"));
//                newworkbook.importSheet(sheetName,0,sheet);
//                newworkbook.write();
//                try {
//                    newworkbook.close();
//                } catch (WriteException e) {
//                    e.printStackTrace();
//                }
//                i = i+1;


//        } catch (BiffException | WriteException e) {
//            e.printStackTrace();
//        }

//        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
//        Sheet sheet = workbook.getSheetAt(0);
//        System.out.println("success");

//        try {
//            Workbook myWorkBook = WorkbookFactory.create(inputStream);
//        } catch (IOException exception) {
//            exception.printStackTrace();
//        }

//        System.out.println("workbook success。");
//        int totalSheets = myWorkBook.getNumberOfSheets();
//        myWorkBook.setSheetName(0,"text");

//    }
