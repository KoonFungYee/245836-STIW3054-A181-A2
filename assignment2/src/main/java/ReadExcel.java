import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;

public class ReadExcel {
    public static ArrayList<newData> excelData = new ArrayList<newData>();

    public static void readExcel(){
        System.out.println("Reading the excel file...");
        try {
            FileInputStream fileInputStream=new FileInputStream("C:\\Users\\Koon Fung Yee\\Desktop\\Starting rank.xlsx");
            XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet=workbook.getSheet("Ranking");
            Iterator<Row>rows=sheet.iterator();

            String newData1,newData2,newData3,newData4,newData5,newData6;
            while (rows.hasNext()){
                Row row = rows.next();
                Cell newColoumn1=row.getCell(0);
                Cell newColoumn2 = row.getCell(2);
                Cell newColoumn3 = row.getCell(3);
                Cell newColoumn4 = row.getCell(4);
                Cell newColoumn5 = row.getCell(5);
                Cell newColoumn6 = row.getCell(6);

                newData1=newColoumn1.toString();
                newData2=newColoumn2.toString();
                newData3=newColoumn3.toString();
                newData4=newColoumn4.toString();
                newData5=newColoumn5.toString();
                newData6=newColoumn6.toString();

                excelData.add(new newData(newData1,newData2,newData3,newData4,newData5,newData6));

            }
            workbook.close();
            fileInputStream.close();

            for (int i=0;i<excelData.size();i++) {
                System.out.printf("%-5s", excelData.get(i).getData1());
                System.out.printf("%-40s", excelData.get(i).getData2());
                System.out.printf("%-8s", excelData.get(i).getData3());
                System.out.printf("%-5s", excelData.get(i).getData4());
                System.out.printf("%-5s", excelData.get(i).getData5());
                System.out.printf("%-5s", excelData.get(i).getData6());
                System.out.println("");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
