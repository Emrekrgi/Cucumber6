package ApachePOI;
/*
   Main den bir metod çağırmak suretiyle, path i ve sheetName i verilen excelden
   istenilen sütun kadar veriyi okuyup bir List e atınız.
   Bu soruda kaynak Excel için : ApacheExcel2.xlsx  in 2.sheet ini kullanabilirsiniz.
 */

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class _11_SoruHoca {
    public static void main(String[] args) {
        String path = "src/test/java/ApachePOI/resource/ApacheExcel2.xlsx";
        String sheetName = "testCitizen";
        int sutunSayisi = 6;

        List<List<String>> data = getData(path, sheetName, sutunSayisi);


        for (List<String> row : data) {
            for (String cell : row) {
                System.out.print(cell + "\t");
            }
            System.out.println();
        }
    }

    public static List<List<String>> getData(String path, String sheetName, int sutunSayisi) {
        List<List<String>> data = new ArrayList<>();

        try {
            FileInputStream excelFile = new FileInputStream(path);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet(sheetName);

            for (Row row : sheet) {
                List<String> rowData = new ArrayList<>();
                for (int i = 0; i < sutunSayisi; i++) {
                    Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cell.setCellType(CellType.STRING); // Hücreyi metin olarak al
                    String cellData = cell.getStringCellValue();
                    rowData.add(cellData);
                }
                data.add(rowData);
            }

            workbook.close();
            excelFile.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        return data;
    }
}