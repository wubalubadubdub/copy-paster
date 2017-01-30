import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Set;

/**
 * Created by bearg on 1/27/2017.
 */
public class ReadDocFile {

    private static final String WORD_DOC_PATH_PREFIX = "C:\\Users\\bearg\\OneDrive\\Documents\\transcriptions\\";
    private static final String TEMPLATE_PATH_PREFIX = "C:\\Users\\bearg\\OneDrive\\Documents\\transcriptions\\";
    private static final String ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    private static String wordDocumentName;
    private static HWPFDocument wordDocument;
    private static HashMap<String, Integer> lettersToNumbers;
    private static HSSFWorkbook template;
    private static FileOutputStream stream;
    private static int rowNumber;
    private static String fontName;
    private static int fontSize;
    private static Set<Integer> sheetNumbersUsed;



    public static void main(String[] args) throws IOException {

        if (args.length < 2) {
            System.out.println("Must supply word filename as an argument and row # (0-based) from the Excel sheet" +
                    "that text should be pasted into");
            System.exit(0);
        }

        try {
            rowNumber = Integer.parseInt(args[1]);
            wordDocumentName = args[0];
            File wordDocFile = new File(WORD_DOC_PATH_PREFIX + wordDocumentName);
            FileInputStream fis = new FileInputStream(wordDocFile);
            wordDocument = new HWPFDocument(fis);

            // not sure yet why fontSize needs to be divided by 2, but the font size was being detected as 22 when
            // it should've been 11
            fontSize = (wordDocument.getRange().getParagraph(0).getCharacterRun(0).getFontSize()) / 2;
            fontName =  wordDocument.getRange().getParagraph(0).getCharacterRun(0).getFontName();
            associateColumnLettersWithNumbers();

            final String excelDocumentName = TEMPLATE_PATH_PREFIX + wordDocumentName.replace(".doc", ".xls");
            template = readFile(excelDocumentName);
            stream = new FileOutputStream(excelDocumentName);

            paragraphLoop();
            putXsInBlankCells();

            template.write(stream);
            stream.close();
            template.close();


        }

        catch (IOException e){
            e.printStackTrace();
        }



    }

    private static void paragraphLoop() throws IOException {
        int paragraphNumber = 0;
        int loopCounter = 0;
        Paragraph currentParagraph;
        sheetNumbersUsed = new HashSet<>();

        while (true) {
            currentParagraph = getCurrentParagraph(paragraphNumber);
            if (currentParagraph == null) {
                break;
            }
            paragraphNumber++;
            String plainTextParagraph = currentParagraph.text();

            if (plainTextParagraph.equals("\r")) {
                continue;
            }

            System.out.println(plainTextParagraph);


            int columnNumber = getColumnNumberFromParagraph(plainTextParagraph);
            HSSFRichTextString rts = getBoldParagraph(currentParagraph);

            int sheetNumber = getSheetNumberFromParagraph(plainTextParagraph);
            sheetNumbersUsed.add(sheetNumber);

            // must call before stripping identifier
            HSSFSheet currentSheet = template.getSheetAt(sheetNumber);

            HSSFCell currentCell = getCell(currentSheet, rowNumber, columnNumber);
            pasteTextIntoCell(currentCell, rts);
            loopCounter++;
            System.out.println("Text pasted into cell " + loopCounter + " times");



        }

    }

    private static int getColumnNumberFromParagraph(String plainTextParagraph) {

        // get first four characters from the paragraph, i.e. the column identifier and colon,
        // e.g. B: or D2: or AC: or AD2:
        String columnIdentifier = plainTextParagraph.substring(0, 4);

        if (columnIdentifier.charAt(1) == ':') { // e.g. A:
            columnIdentifier = ""+ columnIdentifier.charAt(0);
        }

        else if (columnIdentifier.charAt(2) == ':') { // e.g. A1: or AB:
            String determinant = "" + columnIdentifier.charAt(1);

            if (!ALPHABET.contains(determinant)) { // this is the A1: case
                columnIdentifier = ""+ columnIdentifier.charAt(0);
            }

            else { // this is the AB: case
                columnIdentifier = columnIdentifier.substring(0, 2);
            }
        }

        else if (columnIdentifier.charAt(3) == ':') { // e.g. AD2:
            columnIdentifier = columnIdentifier.substring(0, 2);
        }

        return lettersToNumbers.get(columnIdentifier);
    }

    private static int getSheetNumberFromParagraph(String plainTextParagraph) {
        String sheetIdentifier = plainTextParagraph.substring(0, 4);
        if (sheetIdentifier.charAt(1) == ':') { // we should be pasting into sheet 0
            return 0;
        }
        if (sheetIdentifier.charAt(2) == ':') { // e.g. D2: or BB:
            // we should return 0 if charAt(1), aka determinant, is a letter (something like BB:)
            // and return determinant if determinant is a number (something like D2:)
            String determinant = "" + sheetIdentifier.charAt(1);
            return ALPHABET.contains(determinant)? 0:Integer.valueOf(determinant);
        }
        if (sheetIdentifier.charAt(3) == ':') { // e.g. AA1:
            return Integer.parseInt("" + sheetIdentifier.charAt(2));
        }

        throw new IllegalStateException("Could not get sheet number from paragraph. None of the available options" +
                "were matched");
    }

    private static HSSFFont getBoldText() {
        HSSFCellStyle cellStyle = template.createCellStyle();
        HSSFFont boldFont = template.createFont();
        boldFont.setBold(true);
        boldFont.setFontName(fontName);
        boldFont.setFontHeightInPoints((short) fontSize);
        cellStyle.setFont(boldFont);
        return boldFont;

    }

    private static HSSFRichTextString getBoldParagraph(Paragraph currentParagraph) {

        if (currentParagraph.text().equals("\r")) {
            return new HSSFRichTextString(currentParagraph.text());
        }

        String plainTextWithoutIdentifier = stripIdentifier(currentParagraph.text());
        HSSFRichTextString rts = new HSSFRichTextString(plainTextWithoutIdentifier);


        for (int i=0; i < currentParagraph.numCharacterRuns(); i++) {

            // get character runs one at a time
            CharacterRun characterRun = currentParagraph.getCharacterRun(i);

            // if the run of characters is bolded
            if (characterRun.isBold()) {

                // find that substring in the original
                String textToMatch = characterRun.text().trim();

                // TODO add check here for multiple occurences of same bolded string

                int startBold = plainTextWithoutIdentifier.indexOf(textToMatch);
                int endBold = startBold + textToMatch.length();


                // apply bold font to that substring
                rts.applyFont(startBold, endBold, getBoldText());

            }
        }
        return rts;
    }

    private static String stripIdentifier(String plainText) {
        if (plainText.charAt(1) == ':') { // single-letter identifier; three chars to strip (e.g. D:\s)
           return plainText.substring(3);
        }

        if (plainText.charAt(2) == ':') { // e.g. A1:\s or AB:\s
            return plainText.substring(4);
        }

        if (plainText.charAt(3) == ':') { // e.g. BC2:\s
            return plainText.substring(5);
        }

       throw new IllegalStateException("Could not strip identifier. None of the options were matched.");

    }


    private static void pasteTextIntoCell(HSSFCell cell, HSSFRichTextString rts){

        cell.setCellValue(rts);

    }

    private static int getLastFilledColumnNumber(HSSFSheet currentSheet) {

        // want to put Xs in cell range from first column, 0, in rowNumber until the cell directly
        // above is blank, but only if those cells are blank after pasting in text of all paragraphs

        // get the row above the one we're pasting text into
        HSSFRow row = currentSheet.getRow(rowNumber-1);
        int endColumnNumber = 0;
        while (true) { // check cell contents of row above column number
            if (row.getCell(endColumnNumber) == null) { // the cell above is empty
                break;
            }

            endColumnNumber++;
        }

        return endColumnNumber; // actually gives us 1 more than the column number we want to stop at,
        // but this will be dealt with in the loop

    }

    // needs to be called after methods that place text into cells
    private static void putXsInBlankCells() {

        for (Integer sheetNumber : sheetNumbersUsed) {
            HSSFSheet currentSheet = template.getSheetAt(sheetNumber);
            HSSFRow row = currentSheet.getRow(rowNumber);
            for (int column = 0; column < getLastFilledColumnNumber(currentSheet); column++) {
                HSSFCell cell = row.getCell(column);
                if (cell.getStringCellValue().equals("")) { // if the cell is empty
                    cell.setCellValue("X");
                }
                
            }

        }

    }


    // get a paragraph from the word document
    private static Paragraph getCurrentParagraph(int paragraphNumber) throws IOException {

        Range range = wordDocument.getRange();

        int lastParagraphNumber = range.numParagraphs() - 1;
        if (paragraphNumber > lastParagraphNumber) {
            return null;
        }
        return range.getParagraph(paragraphNumber);

    }


    private static HSSFCell getCell(HSSFSheet sheet, int rowNumber, int columnNumber) {
        System.out.println("Getting cell at row " + rowNumber + " , column " + columnNumber);
        HSSFRow row = sheet.getRow(rowNumber);
        if (row == null) {
            throw new IllegalStateException("Trying to get the row in this sheet returned null." +
                    "Is the entire sheet blank? If so, type some text into any cell in that row and " +
                    "run the program again.");
        }
        return row.getCell(columnNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);


    }


    private static HSSFWorkbook readFile(String filename) throws IOException {
        FileInputStream fis = new FileInputStream(filename);
        try {
            return new HSSFWorkbook(fis);
        }
        finally {
            fis.close();
        }

    }

    private static void associateColumnLettersWithNumbers() {
        lettersToNumbers = new HashMap<>();


        for (int i = 0; i < 6; i++) {

            if (i==0) { // single-letter column
                for (int j = 0; j < 26; j++) {
                    lettersToNumbers.put(String.valueOf(ALPHABET.charAt(j)), j);

                }
            }

            else {
                for (int j = 0; j < 26; j++) {
                    lettersToNumbers.put(
                            String.valueOf(ALPHABET.charAt(i-1)) + String.valueOf(ALPHABET.charAt(j)), 26*i+j);

                    // when i = 1, column #s 26-51, or (26*1)+j
                    // i = 2, column #s 52-77, or (26*2)+j
                    // etc

                }

            }

        }

    }
}
