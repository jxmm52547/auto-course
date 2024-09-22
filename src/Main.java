import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class Main {
    public static void main(String[] args) {
        Workbook writeBook = new XSSFWorkbook();
        try (FileInputStream fis = new FileInputStream("C:/Users/12508/Desktop/课表.xlsx");
             Workbook workbook = new XSSFWorkbook(fis)) {

            // 获取第一个工作表
            Sheet sheet = workbook.getSheetAt(0);

            //创建基础sheet
            createBaseSheet(sheet, writeBook);

            // 在指定的工作表中写入课程信息
            writeCourse(sheet, writeBook);

            // 将workbook中的所有单元格设置为自动换行
            setAutoWrapText(writeBook);

            // 将课程信息写入到工作簿中
            write(writeBook);


        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void setAutoWrapText(Workbook workbook){
        // 定义文件路径
        String filePath = "D:/Project/test/src/课表Output.xlsx";
        String outputFilePath = "C:/Users/12508/Desktop/课表_修改后.xlsx";

        // 创建单元格样式并启用自动换行
        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);  // 启用自动换行

        // 遍历工作簿中的每个工作表、每行和每个单元格，并应用自动换行样式
        for (Sheet sheet : workbook) {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    cell.setCellStyle(style);
                }
            }
        }
    }

    public static void writeCourse(Sheet sheet, Workbook writeBook){
        HashMap<String,Integer> nameToIndex = new HashMap<>();
        nameToIndex.put("教学班名称",5);
        nameToIndex.put("授课教师",8);
        nameToIndex.put("星期",9);
        nameToIndex.put("节次",10);
        nameToIndex.put("上课地点",11);
        nameToIndex.put("课程名称",4);

        HashMap<String,Integer> weekToIndex = new HashMap<>();
        weekToIndex.put("1",1);
        weekToIndex.put("2",2);
        weekToIndex.put("3",3);
        weekToIndex.put("4",4);
        weekToIndex.put("5",5);
        weekToIndex.put("6",6);
        weekToIndex.put("7",7);

        HashMap<String,Integer> sectionToIndex = new HashMap<>();
        sectionToIndex.put("0102",1);
        sectionToIndex.put("0304",2);
        sectionToIndex.put("0506",3);
        sectionToIndex.put("0708",4);
        sectionToIndex.put("0910",5);


        for (Row row : sheet){
            if (row.getRowNum() == 0) continue;
            String className = row.getCell(nameToIndex.get("教学班名称")).getStringCellValue();
            String teacherName = row.getCell(nameToIndex.get("授课教师")).getStringCellValue();
            String week = row.getCell(nameToIndex.get("星期")).getStringCellValue();
            String section = row.getCell(nameToIndex.get("节次")).getStringCellValue();
            String location = row.getCell(nameToIndex.get("上课地点")).getStringCellValue();
            String courseName = row.getCell(nameToIndex.get("课程名称")).getStringCellValue();

            if(writeBook.getSheet(className) != null){
                Sheet classSheet = writeBook.getSheet(className);
                int line = sectionToIndex.get(section);
                int column = weekToIndex.get(week);
                classSheet.getRow(line).createCell(column).setCellValue(courseName + " (" + teacherName + ")" + " @" +location);
            }


        }
    }

    public static void createBaseSheet(Sheet sheet, Workbook writeBook){
        List<String> list = new ArrayList<>();
        // 获取标题行
        Row title = sheet.getRow(0);

        for (Cell cell : title) {
            if (cell.getStringCellValue().equals("教学班名称")){
                int index =cell.getColumnIndex();
                //获取 cell 单元所在这一列的所有值
                for (int i = 1; i <= sheet.getLastRowNum(); i++){
                    Cell classCell = sheet.getRow(i).getCell(index);

                    String className = classCell.getStringCellValue();
                    if (!list.contains(className)){
                        list.add(className);
                        Sheet classSheet = writeBook.createSheet(className);
                        classSheet = defaultValue(classSheet);

                    }


                }

            }

        }
    }

    public static Sheet defaultValue(Sheet sheet){
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("星期/节次");
        for (int i = 1; i < 8; i++) {
            row.createCell(i).setCellValue("周" + i);
        }
        for (int i = 1; i < 6; i++){
            List<String> list = new ArrayList<>();
            list.add("0102");
            list.add("0304");
            list.add("0506");
            list.add("0708");
            list.add("0910");
            Row row1 = sheet.createRow(i);
            row1.createCell(0).setCellValue(list.get(i-1));
        }

        return sheet;
    }

    public static void printSheet(Sheet sheet){
        for (Row row1 : sheet) {
            // 遍历每个单元格
            Iterator<Cell> cellIterator = row1.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch (cell.getCellType()) {
                    case STRING:
                        System.out.print(cell.getStringCellValue() + "\t");
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            System.out.print(cell.getDateCellValue() + "\t");
                        } else {
                            System.out.print(cell.getNumericCellValue() + "\t");
                        }
                        break;
                    default:
                        System.out.print("\t");
                        break;
                }
            }
            System.out.println();
        }
    }
    public static void write(Workbook workbook){
        try {
            workbook.write(new FileOutputStream("D:/Project/test/src/课表Output.xlsx"));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


}