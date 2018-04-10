package sk.ptacin.excel;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;
import sk.ptacin.excel.model.CopyPastePath;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URLDecoder;
import java.security.CodeSource;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

/**
 * Created by Michal on 1.6.2016.
 */
@Component
public class CopyProcessor {

    private static final Logger log = LoggerFactory.getLogger(CopyProcessor.class);

    @Value("${startingDir}")
    private String startingDir;

    @Value("${sourceCell}")
    private String sourceCellName;

    @Value("${targetCell}")
    private String tagetCellName;

    @Value("${mapName}")
    private String mapName;

    @Value("${cisloRiadkuZaciatku}")
    private int cisloRiadkuZaciatku;

    @Value("${mapTableName}")
    private String mapTableName;

    @Value("${idColumnName}")
    private String idColumnName;

    @Value("${amountsColumnName}")
    private String amountsColumnName;

    @Value("${ignoreTargetSheet}")
    private String ignoreTargetSheet;


    public void startCopying() throws Exception {
        log.info(" -------------  ZACIATOK spracovania suborov  -------------------");
        if (startingDir.isEmpty()) {
            try {
                startingDir = getJarContainingFolder(CopyProcessor.class) + "\\";
            } catch (Exception e) {
                log.error("Pri nacitani parent adresara doslo k chybe", e);
                throw e;
            }
        }
        log.info("Adresar v ktorom hladam mapovaciu tabulku, zdrojovy a cielovy subror {}", startingDir);

        try {
            writeXLSXFile();
        } catch (Exception e) {
            log.error("Nastala chyba pri kopirovani", e);
            throw e;
        }
        log.info(" ------------- Kopirovanie excelu bolo uspesne UKONCENE -------------------");

    }


    public void writeXLSXFile() throws Exception {

        //FileInputStream fileMap = new FileInputStream("d:\\_Personal\\Roofart\\kopiruj_excel\\mapovacia_tabulka.xlsx");
        //FileInputStream fileMap = new FileInputStream("e:\\Roofart\\kopiruj_excelmapovacia_tabulka.xlsx");
        // File fileMap = new File(dir+"mapovacia_tabulka.xlsx");
        //File fileMap = new File(ExcelHandler.class.getClassLoader().getResource("mapovacia_tabulka.xlsx").getFile());

        //Pridane kvoli chybe pri naplnani tabulky - https://stackoverflow.com/questions/44897500/using-apache-poi-zip-bomb-detected
        ZipSecureFile.setMinInflateRatio(0.0009);

        log.info("------------- START - citanie mapovacej tabulky");
        XSSFWorkbook mapBook = null;
        boolean ignoreTargetSheetValue = false;
        try {
            mapBook = new XSSFWorkbook(startingDir + mapTableName);
        } catch (IllegalStateException ex) {
            log.error("Subor={} nebol najdeny! Skontrolujte ci naozaj existuje v danom umiestneni", startingDir + mapTableName);
        }

        //TODO:zatial beriem len jednu mapu s properties - neskor dorobit aby prechadzalo vsetky mapy
        XSSFSheet sheetMap = mapBook.getSheet(mapName);
        log.info("Zdojova bunka z properties=" + sourceCellName);
        CellReference crSource = new CellReference(sourceCellName);
        log.info("Cielova bunka z properties=" + tagetCellName);
        CellReference crTarget = new CellReference(tagetCellName);
        log.info("Bunka ignorovania cieloveho zosita=" + tagetCellName);
        CellReference ignoreTargetSheetRef = new CellReference(ignoreTargetSheet);
        if (crSource == null) {
            log.error("Na bunke {} nebol najdeny nazov zdrojoveho suboru", sourceCellName);
            throw new Exception("Na bunke B1 nebol najdeny nazov zdrojoveho suboru");
        }
        if (crTarget == null) {
            log.error("Na bunke {} nebol najdeny nazov cieloveho suboru", tagetCellName);
            throw new Exception("Na bunke B2 nebol najdeny nazov cieloveho suboru");
        }
        if (ignoreTargetSheetRef == null) {
            log.info("Na bunke {} nebola najdena hodnota", ignoreTargetSheet);
        }
        Row rowSource = sheetMap.getRow(crSource.getRow());
        Cell cellSource = rowSource.getCell(crSource.getCol());
        String sourceFileName = cellSource.getStringCellValue();

        Row rowTarget = sheetMap.getRow(crTarget.getRow());
        Cell cellTarget = rowTarget.getCell(crTarget.getCol());
        String targetFileName = cellTarget.getStringCellValue();

        Row rowIgnoreTarget = sheetMap.getRow(ignoreTargetSheetRef.getRow());
        Cell cellIgnoreTarget = rowIgnoreTarget.getCell(ignoreTargetSheetRef.getCol());
        ignoreTargetSheetValue = cellIgnoreTarget.getBooleanCellValue();
        if (ignoreTargetSheetValue) {
            log.info("Je zapnute ignorovanie cieloveho harku a v cielovom harku sa bude hladat nazov podla prave prebiehajuceho mesiaca tzn: {}", String.format("%02d", LocalDate.now().getMonthValue()));
        }

        XSSFWorkbook sourceBook = null;
        XSSFWorkbook targetBook = null;
        try {
            sourceBook = new XSSFWorkbook(startingDir + sourceFileName);
        } catch (IllegalStateException ex) {
            log.error("Subor={} nebol najdeny! Skontrolujte ci naozaj existuje v danom umiestneni", startingDir + sourceFileName);
            throw new Exception("Subor=" + startingDir + sourceFileName + " nebol najdeny! Skontrolujte ci naozaj existuje v danom umiestneni");
        }

        try {
            targetBook = new XSSFWorkbook(startingDir + targetFileName);
        } catch (IllegalStateException ex) {

            log.error("Subor={} nebol najdeny! Skontrolujte ci naozaj existuje v danom umiestneni", startingDir + targetFileName);
            throw new Exception("Subor=" + startingDir + targetFileName + " nebol najdeny! Skontrolujte ci naozaj existuje v danom umiestneni");
        }

        if (sourceBook == null || targetBook == null) {
            log.error("Subory {} a {} neboli najdene, skontrolujte ci naozaj existuje v danom umiestneni", startingDir + targetFileName, startingDir + sourceFileName);
            throw new Exception("Subor nebol najdeny! Skontrolujte ci naozaj existuje v danom umiestneni");
        }
        boolean nastavitCielovySheetPodlaCislaMesiaca = ignoreTargetSheet.toUpperCase().equals("TRUE");
        log.info("Subory {} a {} boli uspesne nacitane", startingDir + sourceFileName, startingDir + targetFileName);
        log.info("----- START - nacitanie poloziek mapy ");
        log.info("Nacitavam od riadku {}, ktory je definovany v properties", cisloRiadkuZaciatku);
        List<CopyPastePath> pathsList = new ArrayList<>();
        Iterator rowIterator = sheetMap.rowIterator();
        boolean globalError = false;
        while (rowIterator.hasNext()) {
            //TODO: brat do uvahy aj komentare na riadkoch
            XSSFRow row = (XSSFRow) rowIterator.next();
            // Odpocitavam dva kvoli tomu ze pocitanie riadkov v excel zacina od 0
            int cisloRiadku = row.getRowNum();
            if (cisloRiadku > cisloRiadkuZaciatku - 2) {
                boolean chyba = false;
                CopyPastePath item = new CopyPastePath();
                XSSFCell cellSheetSource = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                XSSFCell cellMenoProduktu = row.getCell(1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                XSSFCell cellId = row.getCell(2, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                XSSFCell cellSheetTarget = row.getCell(3, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                XSSFCell cellCellTaget = row.getCell(4, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cellSheetSource != null) {
                    item.setSourceSheet(cellSheetSource.getStringCellValue());
                } else {
                    log.warn("Na riadku {} nie je vyplneny stlpec {}", cisloRiadku, "A");
                    chyba = true;
                }
                if (cellId != null) {
                    item.setSourceId(cellId.getStringCellValue());
                } else {
                    log.warn("Na riadku {} nie je vyplneny stlpec {}", cisloRiadku, "B");
                    chyba = true;
                }
                if (cellSheetTarget != null) {
                    if (ignoreTargetSheetValue) {
                        item.setTargetSheet(getNameOfSheetBasedOnMonth());
                    } else {
                        item.setTargetSheet(cellSheetTarget.getStringCellValue());
                    }

                } else {
                    if (ignoreTargetSheetValue) {
                        item.setTargetSheet(getNameOfSheetBasedOnMonth());
                    } else {
                        log.warn("Na riadku {} nie je vyplneny stlpec {}", cisloRiadku, "C");
                        chyba = true;
                    }

                }
                if (cellCellTaget != null) {
                    item.setTagetCell(cellCellTaget.getStringCellValue());
                } else {
                    log.warn("Na riadku {} nie je vyplneny stlpec {}", cisloRiadku, "D");
                    chyba = true;
                }

                if (chyba) {
                    log.warn("Riadok {} nebude fungovat kedze nema vyplene vsetky povinne stlpce - vid vyssie", cisloRiadku);
                    //Nastav aj globalnu chybu aby zapisalo do noveho suboru
                    globalError = true;
                } else {
                    pathsList.add(item);
                }

            }

        }
        if (globalError) {
            //String fileName = startingDir+"OZNACENE_CHYBAJUCE_BUNKY-"+ mapTableName;
            //writeWb(mapBook,fileName);
            //log.warn("Zapisujem subor s oznacenymu chybajucimi bunkami {}" , fileName);
        }
        log.info("----- END - nacitanie poloziek mapy ");
        log.info("----- START - spracovanie poloziek mapy ");
        //System.out.println(pathsList);
        //osetrenie ak sa v cielovej bunke vyskytne rovnaka hodnota aby sa vo vysledku neprepisali ale zratali
        HashMap<String, List<Double>> mapaDvojitychHodnot = new HashMap<>();

        for (CopyPastePath cpp : pathsList) {
            log.info("Spracovanie polozky={}", cpp.getSourceId());
            //Nacitanie hodnoty zo zdroja
            XSSFSheet sourceSheet = sourceBook.getSheet(cpp.getSourceSheet());
            //IDX stlpcov by sa dalo zapamatavat do globalnejsej premenenj a najprv by to slo tam a ak nenajde tak oapt vyhlada
            XSSFRow rowBaseOnId = getRowBasedOnCellId(sourceSheet, cpp.getSourceId());
            if (rowBaseOnId == null) {
                log.warn("V zdrojovom subore sa nenachadza c. karta {}", cpp.getSourceId());
            } else {
                int idxOfAmountColumn = getIdxOfColumnName(sourceSheet, amountsColumnName);
                XSSFCell amountCell = rowBaseOnId.getCell(idxOfAmountColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                //TODO: zabezpecit pretypovanie podla cielovej bunky. Ak je blank tak co v tom pripade ?
                int cellType = amountCell.getCellType();
                if (amountCell.getCellType() == Cell.CELL_TYPE_STRING) {
                    String cellString = amountCell.getStringCellValue();
                    cellString = cellString.replace(",", ".");
                    if (!cellString.isEmpty()) {
                        double amountValue = Double.valueOf(cellString);
                        List<Double> listHodnot = mapaDvojitychHodnot.get(cpp.getTagetCell());
                        if (listHodnot == null) {
                            listHodnot = new ArrayList<>();
                            listHodnot.add(amountValue);
                            mapaDvojitychHodnot.put(cpp.getTagetCell() + cpp.getTargetSheet(), listHodnot);
                        } else {
                            listHodnot.add(amountValue);
                        }

                    }
                } else if (amountCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    List<Double> listHodnot = mapaDvojitychHodnot.get(cpp.getTagetCell() + cpp.getTargetSheet());
                    if (listHodnot == null) {
                        listHodnot = new ArrayList<>();
                        listHodnot.add(amountCell.getNumericCellValue());
                        mapaDvojitychHodnot.put(cpp.getTagetCell() + cpp.getTargetSheet(), listHodnot);
                    } else {
                        listHodnot.add(amountCell.getNumericCellValue());
                    }

                }

                //Nacitanie ciela a zapis hodnoty
                XSSFSheet targetSheet = targetBook.getSheet(cpp.getTargetSheet());
                CellReference cr = new CellReference(cpp.getTagetCell());
                XSSFRow rowTarget2 = targetSheet.getRow(cr.getRow());
                XSSFCell cellTarget2 = rowTarget2.getCell(cr.getCol());
                cellTarget2.setCellType(Cell.CELL_TYPE_NUMERIC);
                //Vytiahnutie hodnoty z mapy
                List<Double> viacnosobneHodnoty = mapaDvojitychHodnot.get(cpp.getTagetCell() + cpp.getTargetSheet());
                double resultAmount = 0;
                for (double ra : viacnosobneHodnoty) {
                    resultAmount += ra;
                }
                cellTarget2.setCellValue(resultAmount);
                log.info("Uspesne prekopirovana hodnota {} do bunky {}", resultAmount, cellTarget2.getReference());
                //Nastavenie stylu - zoberiem stary styl a len zmenim farbu pozadia
                getInsertCellStyle(targetBook, cellTarget2);
            }

        }
        //Riadok preratava vsetky vzorce v subore lebo excel si ich kesuje  http://poi.apache.org/spreadsheet/eval.html
        targetBook.getCreationHelper().createFormulaEvaluator().evaluateAll();

        //Zapisanie vysledku
        String targetFileNameToWrite = startingDir + getNameOfFileBasedOnMonthAndYear() +"-"+ targetFileName;
        FileOutputStream targetFile = new FileOutputStream(new File(targetFileNameToWrite));
        targetBook.write(targetFile);
        targetFile.close();

        //Kedze potrebujem mat zapisany kazdy mesiac tak musim prepisat aj zdrojovy subor
        XSSFWorkbook targetBook2 = null;
        try {
            targetBook2 = new XSSFWorkbook(targetFileNameToWrite);
            log.info("Prepisanie povodneho suboru aby obsahoval aj tohto mesacne veci bolo uspesen");
        } catch (IllegalStateException ex) {
            log.error("Nepodarilo sa prepisanie povodneho suboru aby obsahoval aj tohto mesacne veci!");
        }
        FileOutputStream origFile = new FileOutputStream(new File(startingDir + targetFileName));
        targetBook2.write(origFile);
        origFile.close();

    }

    private XSSFRow getRowBasedOnCellId(XSSFSheet sourceSheet, String sourceId) {
        Iterator rowIter = sourceSheet.rowIterator();
        int idxOfIdColumn = getIdxOfColumnName(sourceSheet, idColumnName);
        while (rowIter.hasNext()) {
            XSSFRow row = (XSSFRow) rowIter.next();
            if (row.getRowNum() != 0) { //Overit ci som nasiel index stlpca v ktorom sa nachadzaju IDcka
                if (idxOfIdColumn < 0) {
                    log.error("V zdrojovom subore sa nenasiel stlpec s menom {} , ktory je pouzity ako referencny voci mapovacej tabulke a kopirovanie nemoze pokracovat", idColumnName);
                } else {
                    XSSFCell cell = row.getCell(idxOfIdColumn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (cell != null) {
                        if (cell.getStringCellValue().equals(sourceId)) {
                            return row;
                        }
                    }
                }
            }
        }
        return null;

    }

    private String getNameOfSheetBasedOnMonth(){
        LocalDate dnesnyDatum = LocalDate.now().minusMonths(1);
        return String.format("%02d", dnesnyDatum.getMonthValue());
    }

    private String getNameOfFileBasedOnMonthAndYear(){
        LocalDate dnesnyDatum = LocalDate.now();
        if (getNameOfSheetBasedOnMonth().equals("12")) {
            return getNameOfSheetBasedOnMonth()+ dnesnyDatum.minusYears(1).getYear();
        }
        return getNameOfSheetBasedOnMonth()+ dnesnyDatum.getYear();
    }

    //TODO: zabezpeci aby bolo mozne dat aj znaky s diakritikou
    private int getIdxOfColumnName(XSSFSheet sourceSheet, String columnName) {
        Iterator rowIter = sourceSheet.rowIterator();
        while (rowIter.hasNext()) {
            XSSFRow row = (XSSFRow) rowIter.next();
            if (row.getRowNum() == 0) {

                short minColIx = row.getFirstCellNum();
                short maxColIx = row.getLastCellNum();
                for (short colIx = minColIx; colIx < maxColIx; colIx++) {
                    XSSFCell cell = row.getCell(colIx);
                    if (cell == null) {
                        continue;
                    } else {
                        if (cell.getStringCellValue().trim().equalsIgnoreCase(columnName)) {
                            return colIx;
                        }
                    }
                }
            }
        }
        return -1;

    }


    public void getErrorCellStyle(XSSFWorkbook wb, XSSFCell cell) {
        CellUtil.setCellStyleProperty(cell, wb, CellUtil.FILL_FOREGROUND_COLOR, IndexedColors.RED.getIndex());
        CellUtil.setCellStyleProperty(cell, wb, CellUtil.FILL_PATTERN, CellStyle.SOLID_FOREGROUND);
    }

    public void getInsertCellStyle(XSSFWorkbook wb, XSSFCell cell) {
        CellUtil.setCellStyleProperty(cell, wb, CellUtil.FILL_FOREGROUND_COLOR, IndexedColors.LIGHT_YELLOW.getIndex());
        CellUtil.setCellStyleProperty(cell, wb, CellUtil.FILL_PATTERN, CellStyle.SOLID_FOREGROUND);
        CellStyle style = wb.createCellStyle();
    }

    public void writeWb(XSSFWorkbook wb, String path) throws IOException {
        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(path);
        wb.write(fileOut);
        fileOut.close();
    }

    public String getJarContainingFolder(Class aclass) throws Exception {
        CodeSource codeSource = aclass.getProtectionDomain().getCodeSource();

        File jarFile;

        if (codeSource.getLocation() != null) {
            jarFile = new File(codeSource.getLocation().toURI());
        } else {
            String path = aclass.getResource(aclass.getSimpleName() + ".class").getPath();
            String jarFilePath = path.substring(path.indexOf(":") + 1, path.indexOf("!"));
            jarFilePath = URLDecoder.decode(jarFilePath, "UTF-8");
            jarFile = new File(jarFilePath);
        }
        //Samotna aplikacia bude v subore
        File parentFile = jarFile.getParentFile();
        String returnPath = parentFile.getAbsolutePath();
        if (parentFile.getParentFile() != null) {
            return parentFile.getParentFile().getAbsolutePath();
        }
        return returnPath;
    }
}
