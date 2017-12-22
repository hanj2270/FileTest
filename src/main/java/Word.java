import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by hanj on 2017/12/22.
 */
public class Word {
    public static String readWord(String filePath) {
        String text = "";
        File file = new File(filePath);
        //2003
        if (file.getName().endsWith(".doc")) {
            try {
                FileInputStream stream = new FileInputStream(file);
                WordExtractor word = new WordExtractor(stream);
                text = word.getText();
                //ȥ��word�ĵ��еĶ������
                text = text.replaceAll("(\\r\\n){2,}", "\r\n");
                text = text.replaceAll("(\\n){2,}", "\n");
                stream.close();
            } catch (Exception e) {
                e.printStackTrace();
            }

        } else if (file.getName().endsWith(".docx")) {       //2007
            try {
                OPCPackage oPCPackage = POIXMLDocument.openPackage(filePath);
                XWPFDocument xwpf = new XWPFDocument(oPCPackage);
                POIXMLTextExtractor ex = new XWPFWordExtractor(xwpf);
                text = ex.getText();
                //ȥ��word�ĵ��еĶ������
                text = text.replaceAll("(\\r\\n){2,}", "\r\n");
                text = text.replaceAll("(\\n){2,}", "\n");
                System.out.println(filePath+"ok");
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return text;
    }


    public static String getPost(String plainText) {
        String ret="";
        Pattern p = Pattern.compile("诉称([\\s\\S]*)被告([\\u0391-\\uFFE5]+)辩称");
        Matcher m = p.matcher(plainText);
        while (m.find()) {
            ret+=m.group(1);
        }
        if(ret.equals("")) {
            Pattern p2 = Pattern.compile("检察院指控([\\s\\S]*)没有异议");
            m = p2.matcher(plainText);
            while (m.find()) {
                ret += m.group(1);
            }
        }
        if(ret.equals("")) {
            Pattern p3 = Pattern.compile("检察院指控([\\s\\S]*)原审");
            m = p3.matcher(plainText);
            while (m.find()) {
                ret += m.group(1);
            }
        }
        if(ret.equals("")) {
            Pattern p4 = Pattern.compile("检察院指控([\\s\\S]*)经审理查明");
            m = p4.matcher(plainText);
            while (m.find()) {
                ret += m.group(1);
            }
        }
        if(ret.equals("")) {
            Pattern p6 = Pattern.compile("检察院指控([\\s\\S]*)一审法院");
            m = p6.matcher(plainText);
            while (m.find()) {
                ret += m.group(1);
            }
        }
        System.out.println("原告诉称内容");
        return format(ret);
    }



    public static String getSentenceFact(String plainText){
        String ret="";
        Pattern p = Pattern.compile("本院认为([\\s\\S]*)裁判结果");
        Matcher m = p.matcher(plainText);
        while (m.find()) {
            ret+=m.group(1);
        }
        System.out.println("认定事实内容");
        return format(ret);
    }

    public static String getSentence(String plainText){
        String ret="";
        Pattern p = Pattern.compile("裁判结果([\\s\\S]*)书记员");
        Matcher m = p.matcher(plainText);
        while (m.find()) {
            ret+=m.group(1);
        }
        System.out.println("裁定结果内容");
        return format(ret);
    }

    public static String format(String text){
        if(text.startsWith("：")||text.startsWith("，")){
            text=text.substring(1,text.length());
        }
        if(text.endsWith("代理")){
            text=text.substring(0,text.length()-2);
        }
        return  text;
    }
}
