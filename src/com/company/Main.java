package com.company;

import com.sun.deploy.util.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Index;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {


    public static void main(String[] args) throws IOException {
//        String path = "d:/";
        String path = "./";
//        String fileName = "test";


        BufferedReader bufferedReader = null;
        try {
            bufferedReader = new BufferedReader(new InputStreamReader(System.in));
            System.out.println("输入路径：");
            String read = "E:\\shanghaojia\\2021.05.xls"; //bufferedReader.readLine();
            System.out.println("路径：" + read);
            read(read);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (bufferedReader != null) {
                bufferedReader.close();
            }
        }
        String fileName = "测试数据";
        String fileType = "xlsx";

      //  writer(path, fileName, fileType);
        //       read(path, fileName, fileType);

    }


    private static void writer(List<Score> newScoreList) {
        //创建工作文档对象
        Workbook wb = null;
        String fileType="xls";
        try {
            if (fileType.equals("xls")) {
                wb = new HSSFWorkbook();

            } else if (fileType.equals("xlsx")) {
                wb = new XSSFWorkbook();
            } else {
                System.out.println("您的文档格式不正确！");
            }
            //创建sheet对象
            Title title=new Title();
            Sheet sheet1 = (Sheet) wb.createSheet("sheet1");
            //循环写入行数据
            Row row = (Row) sheet1.createRow(0);

            Cell cell0 = row.createCell(0);
            cell0.setCellValue(title.getNO());
            Cell cell11 = row.createCell(1);
            cell11.setCellValue(title.getName());

            Cell cell22 = row.createCell(2);
            cell22.setCellValue(title.getPartment());

            Cell cell33 = row.createCell(3);
            cell33.setCellValue(title.getFuZhuRen());

            Cell cell44 = row.createCell(4);
            cell44.setCellValue(title.getZhuRen());

            Cell cell55 = row.createCell(5);
            cell55.setCellValue(title.getLingDao());

            Cell cell66 = row.createCell(6);
            cell66.setCellValue(title.getLastChengji());

            Cell cell77 = row.createCell(7);
            cell77.setCellValue(title.getLevel());


            Cell cell88 = row.createCell(8);
            cell88.setCellValue(title.getBeizhu());


            for (int i = 0; i < newScoreList.size(); i++) {
                Row row2 = (Row) sheet1.createRow(i+1);
                Score score=newScoreList.get(i);
                Cell cell = row2.createCell(0);
                cell.setCellValue(i+1);
                Cell cell1 = row2.createCell(1);
                cell1.setCellValue(score.getName());

                Cell cell2 = row2.createCell(2);
                cell2.setCellValue(score.getPartment());

                Cell cell3 = row2.createCell(3);
                cell3.setCellValue(score.getFuZhuRen());

                Cell cell4 = row2.createCell(4);
                cell4.setCellValue(score.getZhuRen());

                Cell cell5 = row2.createCell(5);
                if(score.isTwo==false){
                    cell5.setCellValue(score.getLingDao());
                }else {
                    cell5.setCellValue("");
                }
                Cell cell6 = row2.createCell(6);
                cell6.setCellValue(score.getLastChengji());

            }
            //创建文件流
            OutputStream stream = new FileOutputStream("E:/shanghaojia/panghang.xls");
            //写入数据
            wb.write(stream);
            //关闭文件流
            stream.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (wb != null) {
                try {
                    wb.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }


    public static void read(String path) {
        Workbook wb = null;
        PrintWriter fileWriter = null;
        List<Score> scoreList = new ArrayList<>();
        try {
            InputStream stream = new FileInputStream(path);
            fileWriter = new PrintWriter(new File("all_cities.txt"));
            if (path.endsWith("xls")) {
                wb = new HSSFWorkbook(stream);

            } else if (path.endsWith("xlsx")) {
                wb = new XSSFWorkbook(stream);

            } else {
                System.out.println("您输入的excel格式不正确");

            }

            Sheet sheet1 = wb.getSheetAt(0);
            int rowCount = 0;
            Title title = new Title();
            for (Row row : sheet1) {
                StringBuilder builder = new StringBuilder();
                rowCount++;
                if (rowCount == 1) continue;
                Score score = new Score();
                for (Cell cell : row) {
                    String s = null;
                    Integer index = cell.getColumnIndex();
                    if (index > 5) {
                        continue;
                    }
                    //0序号	1部门	2姓名	3部门副主任评分	4部门主任评分	5分管领导评分
                    switch (cell.getCellType()) {
                        case 0:
                            s = String.valueOf((int) cell.getNumericCellValue()).trim();
                            break;
//                            return "numeric";
                        case 1:
                            s = cell.getStringCellValue().trim();
                            break;
//                            return "text";
                        case 2:
                            s = cell.getCellFormula().trim();
                            break;
//                            return "formula";
                        case 3:
                            s = " ";
                            break;
//                            return "blank";
                        case 4:
                            s = String.valueOf(cell.getBooleanCellValue());
                            break;
//                            return "boolean";
                        case 5:
                            s = String.valueOf(cell.getErrorCellValue());
                            break;
//                            return "error";
                    }

                    switch (index) {
                        case 0:
                            score.setNO(s);
                            break;
//                            return "numeric";
                        case 1:
                            score.setPartment(s);
                            break;
//                            return "text";
                        case 2:
                            score.setName(s);
                            break;
//                            return "formula";
                        case 3:
                            score.setFuZhuRen(s);
                            break;
//                            return "blank";
                        case 4:
                            score.setZhuRen(s);
                            break;
//                            return "boolean";
                        case 5:
                            score.setLingDao(s);
                            break;
//                            return "error";
                    }

                }
                scoreList.add(score);
            }
            System.out.println();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (wb != null) {
                try {
                    wb.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (fileWriter != null) {
                fileWriter.close();
            }
        }

//        System.out.println();
//        for(int i=0;i<scoreList.size();i++){
//            Score score=scoreList.get(i);
//            //System.out.println("i="+i+"  "+score.name+"  "+score.getPartment()+"  "+score.getFuZhuRen());
//        }

        List<Score> newScoreList = new ArrayList<>();
        for (int i = 0; i < scoreList.size(); i += 2) {
            //  System.out.println("i="+i+"  I+1="+(i+1));
            Score scoreOne = scoreList.get(i);
            Score scoreTwo = scoreList.get(i + 1);


            scoreOne.setFuZhuRen(stringTrimAll(scoreOne.getFuZhuRen()));
            scoreOne.setZhuRen(stringTrimAll(scoreOne.getZhuRen()));
            scoreOne.setLingDao(stringTrimAll(scoreOne.getLingDao()));

            scoreTwo.setFuZhuRen(stringTrimAll(scoreTwo.getFuZhuRen()));
            scoreTwo.setZhuRen(stringTrimAll(scoreTwo.getZhuRen()));
            scoreTwo.setLingDao(stringTrimAll(scoreTwo.getLingDao()));


            double fuzhuren1 = Double.valueOf(scoreOne.getFuZhuRen());
            double fuzhuren2 = Double.valueOf(scoreTwo.getFuZhuRen());
            //  System.out.println("fuzhuren1="+fuzhuren1+" fuzhuren2="+fuzhuren2+"   value="+(fuzhuren1+fuzhuren2));

            scoreOne.setFuZhuRen(String.valueOf(fuzhuren1*0.8 + fuzhuren2*0.2));
            scoreOne.setZhuRen(String.valueOf(0.8*Double.valueOf(scoreOne.getZhuRen()) + 0.2*Double.valueOf(scoreTwo.getZhuRen())));
            //    System.out.println("11 ");
            if (scoreOne.getLingDao() != null && scoreOne.getLingDao().trim() != null && scoreOne.getLingDao().length() > 0) {
                scoreOne.setLingDao(String.valueOf(0.8*Double.valueOf(scoreOne.getLingDao()) + 0.2*Double.valueOf(scoreTwo.getLingDao())));
                //  System.out.println("22 scoreOne.getLingDao()=" + scoreOne.getLingDao() + "  scoreTwo.getLingDao()=" + scoreTwo.getLingDao());
                scoreOne.setTwo(false);
            }
            newScoreList.add(scoreOne);
        }
        System.out.println();
        for (int i = 0; i < newScoreList.size(); i++) {
            Score score = newScoreList.get(i);
            double one = Double.valueOf(score.getFuZhuRen());
            String  oneStr= new DecimalFormat("#.00").format(one);
            score.setFuZhuRen(oneStr);

            double two = Double.valueOf(score.getZhuRen());
            String  twoStr= new DecimalFormat("#.00").format(two);
            score.setZhuRen(twoStr);

            double three = 0.0;
            if (score.isTwo) {
                double chengji = one * 0.5 + two * 0.5;
                String  chengjiStr= new DecimalFormat("#.00").format(chengji);
                chengji=Double.valueOf(chengjiStr);
                score.setLastChengji(String.valueOf(chengji));
                score.setChengji(chengji);
            } else {
                three = Double.valueOf(score.getLingDao());
                String  threeStr= new DecimalFormat("#.00").format(three);
                score.setLingDao(threeStr);


                double chengji = one * 0.3 + two * 0.5 + three * 0.2;
                String  chengjiStr= new DecimalFormat("#.00").format(chengji);
                chengji=Double.valueOf(chengjiStr);
                score.setLastChengji(String.valueOf(chengji));
                score.setChengji(chengji);
            }

        }
        Collections.sort(newScoreList, new Comparator<Score>() {
            @Override
            public int compare(Score u1, Score u2) {
                double diff = u1.getChengji() - u2.getChengji();
                if (diff > 0) {
                    return -1;
                } else if (diff < 0) {
                    return 1;
                }
                return 0; //相等为0
            }

        }); // 按年龄排序
        for (int i = 0; i < newScoreList.size(); i++) {
            Score score = newScoreList.get(i);
            System.out.println("i=" + (i + 1) + "  " + score.name + "  " + score.getPartment() + "  " + score.getChengji());
        }
        writer(newScoreList);
    }

    static public String stringTrimAll(String input) {
        input = input.trim();
        if (null == input)
            return "";
        // 正则匹配{空格/换行/回车/制表符/换页符}
        final String regx = "\\s*|\t|\r|\n";
        Pattern patt = Pattern.compile(regx);
        Matcher m = patt.matcher(input);
        return m.replaceAll("");
    }

    public static void read(String path, String fileName, String fileType) {
        Workbook wb = null;
        PrintWriter fileWriter = null;
        try {
            InputStream stream = new FileInputStream(path + fileName + "." + fileType);
            fileWriter = new PrintWriter(new File("all_cities.txt"));
            if (fileType.equals("xls")) {
                wb = new HSSFWorkbook(stream);

            } else if (fileType.equals("xlsx")) {
                wb = new XSSFWorkbook(stream);

            } else {
                System.out.println("您输入的excel格式不正确");

            }
            Sheet sheet1 = wb.getSheetAt(0);
            for (Row row : sheet1) {
                StringBuilder builder = new StringBuilder();
                for (Cell cell : row) {
                    String s = null;
                    if (cell.getColumnIndex() > 5) {
                        continue;
                    }
                    switch (cell.getCellType()) {
                        case 0:
                            s = String.valueOf((int) cell.getNumericCellValue()).trim();
                            break;
//                            return "numeric";
                        case 1:
                            s = cell.getStringCellValue().trim();
                            if (s.equals("√")) {
                                s = "1";
                            }
                            break;
//                            return "text";
                        case 2:
                            s = cell.getCellFormula().trim();
                            break;
//                            return "formula";
                        case 3:
                            s = " ";
                            break;
//                            return "blank";
                        case 4:
                            s = String.valueOf(cell.getBooleanCellValue());
                            break;
//                            return "boolean";
                        case 5:
                            s = String.valueOf(cell.getErrorCellValue());
                            break;
//                            return "error";
                    }
                    if (cell.getColumnIndex() != 5) {
                        s = String.format("%s&", s);
                    }
                    builder.append(s);
//                    System.out.print(s);

                }
                fileWriter.println(builder.toString());
//                builder.append("/n");
//                System.out.println();

            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (wb != null) {
                try {
                    wb.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (fileWriter != null) {
                fileWriter.close();
            }
        }

    }

}
