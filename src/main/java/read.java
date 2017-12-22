import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;

public class read {

    private static String test1 ;

    private static String Excel_path="D:\\result.xls";

    public static void main(String[] args) {

            test1= "D:\\共享目录\\文书";
//            test1="D:\\共享目录\\文书\\1-通用电气石油天然气压力控制（苏州）有限公司与张成坤劳动争议二审民事判决书23287951.doc";

        ArrayList<File> files=getLogFileFromFolder(new File(test1));
        File xlsFile = new File(Excel_path);
        try {
            WritableWorkbook workbook = Workbook.createWorkbook(xlsFile);
            WritableSheet sheet = workbook.createSheet("sheet1", 0);
            int i=1;
            for (File f:files){
                String result=Word.readWord(f.getPath());
                        // 创建一个工作表
                sheet.addCell(new Label(0,i,f.getName()));
                sheet.addCell(new Label(1, i,Word.getPost(result)));
                sheet.addCell(new Label(2, i,Word.getSentenceFact(result)));
                sheet.addCell(new Label(3, i,Word.getSentence(result)));
                i++;
            }
            workbook.write();
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }



    private static ArrayList<File> getLogFileFromFolder(File Targetfile) {
        ArrayList<File> fileList = new ArrayList<File>();
        String name = Targetfile.getName();
       if (Targetfile.isDirectory()) {
            File[] FilesArray = Targetfile.listFiles();
            for (File file : FilesArray) {
                fileList.addAll(getLogFileFromFolder(file));
            }
        }else{
           fileList.add(Targetfile);
       }
        return fileList;
    }


}
