package sk.ptacin.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by Michal Ptacin (michal.ptacin@icz.sk) on 6. 7. 2016.
 */
@Service
public class TestService {

    private static final Logger log = LoggerFactory.getLogger(TestService.class);

    public void startCopying() throws IOException {

        FileInputStream fileMap = null;
        fileMap = new FileInputStream("c:\\Users\\michalp\\Desktop\\zosit_date-time_graf.xlsx");


        XSSFWorkbook mapBook = null;
        try{
            mapBook = new XSSFWorkbook(fileMap);
        } catch (IllegalStateException ex)
        {
            ex.printStackTrace();
        }

        CreationHelper createHelper = mapBook.getCreationHelper();

        XSSFSheet sheetMap = mapBook.getSheet("toto");
        CellReference crSource = new CellReference("A1");
        CellReference crTarget = new CellReference("B1");

        Row rowSource = sheetMap.getRow(crSource.getRow());
        Cell cellSource = rowSource.getCell(crSource.getCol());

        Row rowTarget = sheetMap.getRow(crTarget.getRow());
        Cell cellTarget = rowTarget.getCell(crTarget.getCol());


        CellReference a2 = new CellReference("A2");
        CellReference b2 = new CellReference("B2");

        Row rowSource2 = sheetMap.getRow(a2.getRow());
        Cell cellSource2 = rowSource2.getCell(a2.getCol());

        Row rowTarget2 = sheetMap.getRow(b2.getRow());
        Cell cellTarget2 = rowTarget2.getCell(b2.getCol());

        // we style the second cell as a date (and time).  It is important to
        // create a new cell style from the workbook otherwise you can end up
        // modifying the built in style and effecting not only this cell but other cells.
        CellStyle cellStyle = mapBook.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("DD.MM.YYYY h:mm"));
        cellSource2.setCellValue(new Date());
        cellSource2.setCellStyle(cellStyle);


        CellStyle style = mapBook.createCellStyle();
        DataFormat df = mapBook.createDataFormat();
        style.setDataFormat(df.getFormat("HH:mm:ss"));

        cellTarget2.setCellFormula("TIME(9,30,00)");
        cellTarget2.setCellType(Cell.CELL_TYPE_FORMULA);
        cellTarget2.setCellStyle(style);
        createHelper.createFormulaEvaluator().evaluateFormulaCell(cellTarget2);


        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("c:\\Users\\michalp\\Desktop\\workbook7.xlsx");
        mapBook.write(fileOut);
        fileOut.close();

        System.out.println("Len datum=" + getCellValueAsString(cellSource2) );
        System.out.println("Len cas=" + getCellValueAsString(cellTarget2) );
    }

    private String getCellValueAsString(Cell poiCell){

        if (poiCell.getCellType()==Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(poiCell)) {
            //get date
            Date date = poiCell.getDateCellValue();

            //set up formatters that will be used below
            SimpleDateFormat formatTime = new SimpleDateFormat("HH:mm:ss");
            SimpleDateFormat formatYearOnly = new SimpleDateFormat("yyyy");

        /*get date year.
        *"Time-only" values have date set to 31-Dec-1899 so if year is "1899"
        * you can assume it is a "time-only" value
        */
            String dateStamp = formatYearOnly.format(date);

            if (dateStamp.equals("1899")){
                //Return "Time-only" value as String HH:mm:ss
                return formatTime.format(date);
            } else {
                //here you may have a date-only or date-time value

                //get time as String HH:mm:ss
                String timeStamp =formatTime.format(date);

                if (timeStamp.equals("00:00:00")){
                    //if time is 00:00:00 you can assume it is a date only value (but it could be midnight)
                    //In this case I'm fine with the default Cell.toString method (returning dd-MMM-yyyy in case of a date value)
                    return poiCell.toString();
                } else {
                    //return date-time value as "dd-MMM-yyyy HH:mm:ss"
                    return poiCell.toString()+" "+timeStamp;
                }
            }
        }

        //use the default Cell.toString method (returning "dd-MMM-yyyy" in case of a date value)
        return poiCell.toString();
    }
}
