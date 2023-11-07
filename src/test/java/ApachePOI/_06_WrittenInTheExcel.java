package ApachePOI;

import org.apache.poi.ss.usermodel.*;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class _06_WrittenInTheExcel {
    public static void main(String[] args) throws IOException {



    String  path ="src/test/java/ApachePOI/resource/WriteInTheExcelFile.xlsx";

        FileInputStream inputStream=new FileInputStream(path);//okuma yonunde acildi
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet=workbook.getSheetAt(0);

        //hafizada yazma islemlerine basliyorum
        int sonSatirIndex=sheet.getPhysicalNumberOfRows();

        Row yeniSatir=sheet.createRow(sonSatirIndex);
        Cell yeniHucre= yeniSatir.createCell(0);//ilk hucre olusturuldu
        yeniHucre.setCellValue("Merhaba Dunya");
        //yazma isi bitti
        inputStream.close();
        //dosyaya kaydetmeye geciyorum

        //bunun icin dosyayi yazma yonunde ac
        FileOutputStream outputStream=new FileOutputStream(path);
        workbook.write(outputStream);
        workbook.close();//hafizayi bosalt
        outputStream.close();//yazma kanalini kapat

        System.out.println("islem bitti");
}
}