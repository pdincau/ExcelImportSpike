import java.io.File;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Spike {
    public static void main(String[] args) {
        if (args.length > 0) {
            String filePath = args[0];
            String sheetName = args[1];
            int rowIndex = Integer.parseInt(args[2]);
            int colIndex = Integer.parseInt(args[3]);
            readFile(new File(filePath), sheetName, rowIndex, colIndex);
        } else {
            System.out.println("Missing command line arguments!");
            return;
        }

    }

    private static void readFile(File file, String sheetName, int rowIndex, int colIndex) {
        try {
            Workbook workbook = WorkbookFactory.create(file);

            Sheet sheet = workbook.getSheet(sheetName);
            Row row = sheet.getRow(rowIndex);
            Cell cell = row.getCell(colIndex);

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            CellValue cellValue = evaluator.evaluate(cell);

            System.out.println("=== Sheet " + sheet.getSheetName());
            System.out.println(cell.toString());
            System.out.println(cellValue.getStringValue());
        } catch (Exception e) {
            System.err.println("Generated exception: " + e);
        }

    }
}

