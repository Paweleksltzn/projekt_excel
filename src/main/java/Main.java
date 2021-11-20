import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.util.Iterator;
public class Main {

    /* This is my first java program.
     * This will print 'Hello World' as the output
     */
    static int liczba_booli = 0;
    static int liczba_stringow = 0;
    static int liczba_numerycznych = 0;

    public static void main(String []args) {
        // Creating a Workbook from an Excel file (.xls or .xlsx)

        try {
            Workbook workbook = WorkbookFactory.create(new File("C:\\Users\\Student\\Downloads\\Financial Sample (1).xlsx"));
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while (sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();
                System.out.println("=> " + sheet.getSheetName());
                sheet.forEach(row -> {
                    row.forEach(cell -> {
                        zliczaWartosci(cell);
                    });
                    System.out.println();
                });
            }
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }


        System.out.println("Liczba stringow: " + liczba_stringow);
        System.out.println("Liczba numerycznych: " + liczba_numerycznych);
        System.out.println("Czy bool parzysty: " + (liczba_booli % 2 == 0 ? "Tak" : "Nie"));
    }



     private static void zliczaWartosci(Cell cell) {
        switch (cell.getCellType()) {
            case BOOLEAN:
                liczba_booli++;
                break;
            case STRING:
                liczba_stringow++;
                break;
            case NUMERIC:
                liczba_numerycznych++;
                break;
        }

        System.out.print("\t");
    }

}