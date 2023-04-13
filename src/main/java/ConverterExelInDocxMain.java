import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import java.io.*;
import java.util.List;
import java.util.Map;

public class ConverterExelInDocxMain {
    private static final String FILE_NAME = "src/main/resources/exeal.xlsx";
    private static final int[] NUMBERING = {0, 0, 0};

    public static void main(String[] args) throws IOException {

        XWPFDocument newDocFile = new XWPFDocument();
        XWPFParagraph paragraph = newDocFile.createParagraph();
        XWPFRun run = paragraph.createRun();

        try {
            FileInputStream excelFile = new FileInputStream(FILE_NAME);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheetOne = workbook.getSheet("Модуль 1");


            List<Integer> filledСolumnsSheetOne = SampleDocx.roundCellAndRow(sheetOne).get(1);
            Map<Integer, List<String>> excelMap =
                    SampleDocx.createMapFromExcel(SampleDocx.copyDataFromExcel(sheetOne));
            SampleDocx.creatSample(filledСolumnsSheetOne, excelMap, NUMBERING, run);

        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        FileOutputStream outputStream = new FileOutputStream("test1.docx");
        newDocFile.write(outputStream);
        outputStream.close();
    }

}



