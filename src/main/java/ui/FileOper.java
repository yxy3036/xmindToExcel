package ui;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.*;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class FileOper {
    static ArrayList<LinkedList<String>> als = new ArrayList<LinkedList<String>>();
    String[] titles = new String[]{"测试集节点名称(必填)","主题(必填)","用例类型(必填)","优先级(必填)","是否自动化","前置条件(必填)","用例步骤(必填)","预期结果(必填)"};

    //读取xml文件内容
    public boolean analysisXML(String strPath){
        SAXReader sax = new SAXReader();
        File xmlFile=new File(strPath);
        if(!xmlFile.exists()){
            System.out.println("没有得到xmind解压的xml文件!");
            return false;
        }
        Document document = null;
        try {
            document = sax.read(xmlFile);
        } catch (DocumentException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }//获取document对象,如果文档无节点，则会抛出Exception提前结束
        Element root=document.getRootElement();//获取根节点
        this.getNodes(root);//从根节点开始遍历所有节点

        //当只存在中心节点时，获取中心节点
        if(als.size() == 0){
            List<Element> node = root.elements();
            for(Element e:node){
                if(e.getName().equals("topic")){
                    List<Element> ee = e.elements();
                    LinkedList<String> ls = new LinkedList<>();
                    ls.add(ee.get(0).getTextTrim());
                    als.add(ls);
                }
            }
        }
        return true;
    }

    //获取最长链表的长度
    public int getMaxNum(){
        int max = 0;
        for(int i = 0; i < als.size(); i++){
            if(als.get(i).size() > max){
                max = als.get(i).size();
            }
        }
        return max;
    }

    //清理临时文件
    public void cleanTemp(File file){
        if(file.isDirectory()){
            File[] fl = file.listFiles();
            if(fl.length != 0){
                for(int i = 0; i< fl.length; i++){
                    cleanTemp(fl[i]);
                }
            }
            file.delete();
        }else{
            file.delete();
        }
    }

    //生成excel文件
    public boolean writeExcel(String sourceFilePath){
        XSSFWorkbook wb= null;

        File xlsxFile = new File(sourceFilePath);
        if(xlsxFile.exists()){
            System.out.println("该文件已存在，无法写入");
            return false;
        }
        Sheet sheet =null;
        wb = new XSSFWorkbook();
        sheet = (Sheet) wb.createSheet("sheet1");

        //XSSFFont font = wb.createFont();
        //font.setFontHeightInPoints((short) 12);
        XSSFCellStyle style = wb.createCellStyle();
        style.setWrapText(true);
        //style.setFont(font);
        Row row = sheet.createRow(0);
        row.setHeight((short) 400);

        int max = titles.length;
        for(int j = 0; j < max; j++){
            sheet.setColumnWidth(j, 30*256);
            Cell cell = row.createCell(j);
            cell.setCellStyle(style);
            cell.setCellValue(titles[j]);
        }

        String preConTag = "&&";
        String priorityTag = "||";
        StringBuffer sb = new StringBuffer();
        for (int j=0;j<priorityTag.length();j++){
            sb.append("\\" + priorityTag.charAt(j));
        }
        String sp = preConTag + "|" + sb.toString();
        //循环写入行数据
        for (int i = 0; i < als.size(); i++) {
            row = (Row) sheet.createRow(i+1);
            row.setHeight((short) 500);

            String setNode="";//测试集节点名称
            String title="";//主题
            String preCondition = "";//前置条件
            String step = "";//用例步骤
            String expResult = "";//预期结果
            String priority = "";//优先级


            List<String> titleList = new ArrayList();
            List<String> stepList = new ArrayList();
            List<String> preConditionList = new ArrayList();
            List<String> priorityList = new ArrayList();

            int size = als.get(i).size();
            for(int j = 0; j < size; j++){
                /*最后一级为预期结果*/
                if (j == size-1){
                    expResult = als.get(i).get(j);
                }
                /*第2级到倒数二级为标题、步骤*/
                if (j < size-1 && j>0){
                    String[] s1 = als.get(i).get(j).split(sp);
                    if (!s1[0].equals("")) {
                        titleList.add(s1[0]);
                        stepList.add(s1[0]);
                    }
                    switch (s1.length){
                        case 2:
                            if (als.get(i).get(j).contains(preConTag)){
                                preConditionList.add(s1[1]);
                            }else {
                                priorityList.add(s1[1]);
                            }
                            break;
                        case 3:
                            if (als.get(i).get(j).indexOf("&&")<als.get(i).get(j).indexOf("||")){
                                preConditionList.add(s1[1]);
                                priorityList.add(s1[2]);
                            }else {
                                preConditionList.add(s1[2]);
                                priorityList.add(s1[1]);
                            }
                            break;
                    }
                }
            }
            //指定测试集节点名
            setNode = "房产直销渠道项目/PC/房产CRM/工作区/KP列表/负责业务";
            //拼接标题
            title = String.join("-",titleList);

            //拼步骤
            for (int m=1; m<=stepList.size(); m++)
                if (stepList.get(m-1) != "") {
                    step += m + "." + stepList.get(m - 1) + "\n";
                }
            //拼前置条件
            if (preConditionList.size() != 0) {
                for (int n = 1; n <= preConditionList.size(); n++) {
                    preCondition += n + "." + preConditionList.get(n - 1) + "\n";
                }
            }else {
                preCondition = "无";
            }
            //取优先级
            if (priorityList.size() != 0){
                priority = priorityList.get(priorityList.size()-1);
            }
            /*excel各列数据*/
            for(int j = 0; j < max; j++){
                Cell cell = row.createCell(j);
                cell.setCellStyle(style);
                switch(j){
                    case 0 :
                        cell.setCellValue(setNode);
                        break;
                    case 1 :
                        cell.setCellValue(title);
                        break;
                    case 2 :
                        cell.setCellValue("功能测试");
                        break;
                    case 3 :
                        cell.setCellValue(priority);
                        break;
                    case 4 :
                        cell.setCellValue("否");
                        break;
                    case 5 :
                        cell.setCellValue(preCondition);
                        break;
                    case 6 :
                        cell.setCellValue(step);
                        break;
                    case 7 :
                        cell.setCellValue(expResult);
                        break;
                }
            }
        }

        OutputStream stream = null;
        try {
            stream = new FileOutputStream(xlsxFile);
            //写入数据

            wb.write(stream);

        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }finally {
            try {
                stream.close();
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }
        return true;
    }

    //递归遍历节点
    public void getNodes(Element node){
        List<Element> listElement=node.elements(); //递归遍历当前节点所有的子节点

        //遍历所有一级子节点
        for(Element e:listElement){
            List<Element> ee = e.elements();

            //找到最后一个节点
            if(e.getName().equals("topic") && ee.size() == 1){
                LinkedList<String> ls = new LinkedList<>();
                ls.add(ee.get(0).getTextTrim());
                als.add(ls);

                Element nodePa = e.getParent();

                //回溯路径给list赋值
                while(!nodePa.getName().equals("sheet")){
                    if(nodePa.getName().equals("topic")){
                        List<Element> eee = nodePa.elements();
                        als.get(als.size()-1).add(0, eee.get(0).getTextTrim());
                    }
                    nodePa = nodePa.getParent();
                }
            }
            this.getNodes(e);//递归遍历
        }
    }

    //解压压缩文件
    public  void unZipFiles(String zipFile){
        long startTime=System.currentTimeMillis();
        try {
            ZipInputStream Zin=new ZipInputStream(new FileInputStream(
                    zipFile));//输入源zip路径
            BufferedInputStream Bin=new BufferedInputStream(Zin);
            String Parent = new File(zipFile).getParent() + "\\tm"; //输出路径（文件夹目录）
            File Fout=null;
            ZipEntry entry;
            try {
                while((entry = Zin.getNextEntry())!=null && !entry.isDirectory()){
                    Fout=new File(Parent,entry.getName());
                    if(!Fout.exists()){
                        (new File(Fout.getParent())).mkdirs();
                    }
                    FileOutputStream out=new FileOutputStream(Fout);
                    BufferedOutputStream Bout=new BufferedOutputStream(out);
                    int b;
                    while((b=Bin.read())!=-1){
                        Bout.write(b);
                    }
                    Bout.close();
                    out.close();
                }

            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }finally {
                try {
                    Bin.close();
                    Zin.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        long endTime=System.currentTimeMillis();
        System.out.println("耗费时间： "+(endTime-startTime)+" ms");
    }
}


