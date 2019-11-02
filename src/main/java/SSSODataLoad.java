import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class SSSODataLoad {
    public static void main(String[] args) throws JSONException, FileNotFoundException, IOException, InvalidFormatException {
        // You can specify your excel file path.
        String file = "C:\\Users\\Priya\\ProjectC\\data exports\\SSSO\\SSSO_Events-Nov 2019.xlsx";

        FileInputStream inp = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(inp);

        // Get the first Sheet.
        Sheet sheet = workbook.getSheetAt(0);

        // Start constructing JSON.
        JSONObject json = new JSONObject();

        // Iterate through the rows.
        JSONArray rows = new JSONArray();
        for (Iterator<Row> rowsIT = sheet.rowIterator(); rowsIT.hasNext(); ) {
            Row row = rowsIT.next();
            JSONObject jRow = new JSONObject();

            // Iterate through the cells.
            JSONArray cells = new JSONArray();
            for (Iterator<Cell> cellsIT = row.cellIterator(); cellsIT.hasNext(); ) {
                Cell cell = cellsIT.next();
                try {
                    cells.put(cell.getStringCellValue());
                } catch(IllegalStateException e) {
                    cells.put(cell.getNumericCellValue());
                }
            }
            jRow.put("cell", cells);
            rows.put(jRow);
        }

        // Create the JSON.
        json.put("rows", rows);

        System.out.println("JSON : " + json);

        // Get the JSON text.
       // return json.toString();
    }

}