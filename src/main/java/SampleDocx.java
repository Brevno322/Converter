import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import java.util.*;

public class SampleDocx {

    public static List<List<String>> copyDataFromExcel(Sheet sheet) {
        List<String> oneCell = new ArrayList<>();
        List<String> twoCell = new ArrayList<>();
        List<String> threeCell = new ArrayList<>();
        List<String> fourCell = new ArrayList<>();

        int[] arrayIdCell = createArrayID();

        addDataFromCell(oneCell, sheet, arrayIdCell[0]);
        addDataFromCell(twoCell, sheet, arrayIdCell[1]);
        addDataFromCell(threeCell, sheet, arrayIdCell[2]);
        addDataFromCell(fourCell, sheet, arrayIdCell[3]);

        roundCellAndRow(sheet);

        List<List<String>> allCellForExcel = new ArrayList<>();
        allCellForExcel.add(oneCell);
        allCellForExcel.add(twoCell);
        allCellForExcel.add(threeCell);
        allCellForExcel.add(fourCell);


        return allCellForExcel;
    }


    public static int[] createArrayID() {
        int[] a = {0, 1, 2, 3};
        return a;
    }

    public static void creatSample(List<Integer> filledСolumn,
                                   Map<Integer, List<String>> excelMap, int[] numbering, XWPFRun run) {
        for (Integer idCell : filledСolumn) {
            List<String> listRequiredCell = excelMap.get(idCell);
            for (String data : listRequiredCell) {
                switch (idCell) {
                    case 0:
                        numbering[0] += 1;
                        numbering[1] = 0;
                        numbering[2] = 0;
                        run.setText(data + ". ");
                        break;
                    case 1:
                        run.setText(data);
                        run.addBreak();
                        break;
                    case 2:
                        numbering[1] += 1;
                        numbering[2] = 0;
                        if (numbering[1] == 10) {
                            numbering[1] = 0;
                            numbering[0] += 1;
                        }

                        run.addTab();
                        run.setText(numbering[0] + "." + numbering[1] + "." + " " + data);
                        run.addBreak();
                        break;
                    case 3:
                        numbering[2] += 1;
                        if (numbering[2] == 10) {
                            numbering[2] = 0;
                            numbering[1] += 1;
                        }
                        run.addTab();
                        run.setText(numbering[0] + "." + numbering[1] + "." + numbering[2] + " " + data);
                        run.addBreak();
                        break;
                }
                listRequiredCell.remove(data);
                break;
            }
        }
    }

    public static List<String> addDataFromCell(List<String> listCell, Sheet sheet, int idCell) {
        for (Row currentRow : sheet) {

            if (currentRow.getRowNum() > 0) {

                for (Cell currentCell : currentRow) {

                    if (currentCell.getColumnIndex() == idCell) {

                        if (currentCell.getCellType() == CellType.STRING) {

                            listCell.add(currentCell.getStringCellValue());

                        } else if (currentCell.getCellType() == CellType.NUMERIC) {

                            listCell.add(String.valueOf(currentCell.getNumericCellValue()));
                        }
                    }
                }
            }
        }
        return listCell;

    }


    public static List<List<Integer>> roundCellAndRow(Sheet sheet) {
        List<Integer> listRow = new ArrayList<>();
        List<Integer> listCel = new ArrayList<>();
        List<List<Integer>> allData = new ArrayList<>();

        for (Row row : sheet) {

            if (row.getRowNum() > 0) {

                for (Cell cell : row) {

                    if (cell.getColumnIndex() < 4) {

                        if (cell.getCellType() == CellType.STRING) {

                            listRow.add(row.getRowNum());
                            listCel.add(cell.getColumnIndex());

                        } else if (cell.getCellType() == CellType.NUMERIC) {

                            listRow.add(row.getRowNum());
                            listCel.add(cell.getColumnIndex());

                        }
                    }
                }
            }
        }

        allData.add(listRow);
        allData.add(listCel);

        return allData;
    }

    public static Map<Integer, List<String>> createMapFromExcel(List<List<String>> allCell) {

        Map<Integer, List<String>> map = new HashMap<>();

        map.put(0, allCell.get(0));
        map.put(1, allCell.get(1));
        map.put(2, allCell.get(2));
        map.put(3, allCell.get(3));

        return map;
    }

}



