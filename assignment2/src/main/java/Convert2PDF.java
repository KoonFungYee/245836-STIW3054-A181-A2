import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import java.io.FileOutputStream;

public class Convert2PDF {

    public static void write(){
        try {
            Document document = new Document();
            PdfPTable table = new PdfPTable(new float[]{2, 12, 3, 3, 2, 8});


            try {
                for (int i = 0; i < ReadExcel.excelData.size(); i++) {
                    table.addCell(ReadExcel.excelData.get(i).getData1());
                    table.addCell(ReadExcel.excelData.get(i).getData2());
                    table.addCell(ReadExcel.excelData.get(i).getData3());
                    table.addCell(ReadExcel.excelData.get(i).getData4());
                    table.addCell(ReadExcel.excelData.get(i).getData5());
                    table.addCell(ReadExcel.excelData.get(i).getData6());
                }

                System.out.println("Writing to PDF file...");
                PdfWriter.getInstance(document, new FileOutputStream("C:\\Users\\Koon Fung Yee\\Desktop\\Starting rank.pdf"));
                document.open();
                document.add(table);
                document.close();

                System.out.println("Successful.");
            } catch (Exception e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
