import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class scrap2Excel {
    public static ArrayList<Data> data=new ArrayList<Data>();


    public static void scrapData(){
        try{
            System.out.println("Accessing...");

            Document source= Jsoup.connect("http://chess-results.com/tnr380806.aspx?lan=1&zeilen=99999").get();
            Element table=source.select("table.CRs1").get(0);
            Elements rows=table.select("tr");

            for (Element row:rows){
                Element data1=row.select("td").get(0);
                Element data2=row.select("td").get(1);
                Element data3=row.select("td").get(2);
                Element data4=row.select("td").get(3);
                Element data5=row.select("td").get(4);
                Element data6=row.select("td").get(5);
                Element data7=row.select("td").get(6);
                String coloumn1=data1.text();
                String coloumn2=data2.text();
                String coloumn3=data3.text();
                String coloumn4=data4.text();
                String coloumn5=data5.text();
                String coloumn6=data6.text();
                String coloumn7=data7.text();

                data.add(new Data(coloumn1,coloumn2,coloumn3,coloumn4,coloumn5,coloumn6,coloumn7));
            }

            /*for (int i=0;i<data.size();i++){
                System.out.printf("%3s",data.get(i).getData1());
                System.out.printf("%-1s",data.get(i).getData2());
                System.out.printf("%-40s",data.get(i).getData3());
                System.out.printf("%-10s",data.get(i).getData4());
                System.out.printf("%-5s",data.get(i).getData5());
                System.out.printf("%-5s",data.get(i).getData6());
                System.out.println(data.get(i).getData7());
            }*/


        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    public static void write2Excel(){

        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("Ranking");
        System.out.println("Writing to excel file...");

        try {
            for (int i=0;i<data.size();i++){
                XSSFRow row=sheet.createRow(i);

                XSSFCell cell=row.createCell(0);
                cell.setCellValue(data.get(i).getData1());
                XSSFCell cell2=row.createCell(1);
                cell2.setCellValue(data.get(i).getData2());
                XSSFCell cell3=row.createCell(2);
                cell3.setCellValue(data.get(i).getData3());
                XSSFCell cell4=row.createCell(3);
                cell4.setCellValue(data.get(i).getData4());
                XSSFCell cell5=row.createCell(4);
                cell5.setCellValue(data.get(i).getData5());
                XSSFCell cell6=row.createCell(5);
                cell6.setCellValue(data.get(i).getData6());
                XSSFCell cell7=row.createCell(6);
                cell7.setCellValue(data.get(i).getData7());
            }
            FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\Koon Fung Yee\\Desktop\\Starting rank.xlsx");
            workbook.write(fileOutputStream);
            workbook.close();
            System.out.println("done");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

