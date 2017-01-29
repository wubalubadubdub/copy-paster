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

/**
 * Created by bearg on 1/27/2017.
 */
public class ReadDocFile {

    private static final String WORD_DOC_PATH_PREFIX = "C:\\Users\\bearg\\OneDrive\\Documents\\transcriptions\\";
    private static final String TEMPLATE_PATH_PREFIX = "C:\\Users\\bearg\\OneDrive\\Documents\\transcriptions\\";
    private static String wordDocumentName;
    private static HWPFDocument wordDocument;
    private static HashMap<String, Integer> lettersToNumbers;
    private static HSSFWorkbook template;
    private static FileOutputStream stream;
    private static HSSFSheet sheet;
    private static int rowNumber;

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
            associateColumnLettersWithNumbers();

            final String excelDocumentName = TEMPLATE_PATH_PREFIX + wordDocumentName.replace(".doc", ".xls");
            template = readFile(excelDocumentName);
            stream = new FileOutputStream(excelDocumentName);
            sheet = getSheet();

            paragraphLoop();

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
            HSSFCell currentCell = getCell(sheet, rowNumber, columnNumber);
            pasteTextIntoCell(currentCell, rts);
            loopCounter++;
            System.out.println("Text pasted into cell " + loopCounter + " times");



        }

    }

    private static int getColumnNumberFromParagraph(String plainTextParagraph) {

        // get first two characters from the paragraph, i.e. the column identifier
        String columnIdentifier = plainTextParagraph.substring(0, 2);

        // if 2nd char is ":", we have a single-letter column identifier
        if (columnIdentifier.charAt(1) == ':') {
            columnIdentifier = String.valueOf(columnIdentifier.charAt(0));
        }

        // System.out.println(columnIdentifier + " mapped to " + lettersToNumbers.get(columnIdentifier));
        return lettersToNumbers.get(columnIdentifier);
    }

    private static HSSFFont getBoldText() {
        HSSFFont boldFont = template.createFont();
        boldFont.setBold(true);
        return boldFont;

    }

    private static HSSFCellStyle getCellStyle() {
        HSSFCellStyle cellStyle = template.createCellStyle();
        HSSFFont cellFont = template.createFont();
        cellFont.setFontName("Calibri");
        cellFont.setFontHeightInPoints((short) 10);
        cellStyle.setFont(cellFont);
        return cellStyle;

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

            // TODO use these to detect the font name and size automatically
            final String fontName = characterRun.getFontName();
            final int fontSize = characterRun.getFontSize();

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

        // two-letter identifier; four chars to strip (e.g. AB:\s)
        return plainText.substring(4);

    }




    private static void pasteTextIntoCell(HSSFCell cell, HSSFRichTextString rts){

        cell.setCellStyle(getCellStyle());
        cell.setCellValue(rts);

    }


    // get a paragraph from the word document
    private static Paragraph getCurrentParagraph(int paragraphNumber) throws IOException {

        Range range = wordDocument.getRange();
       // System.out.println("Number of paragraphs is " + range.numParagraphs());
      //  System.out.println("Currently on paragraph number " + paragraphNumber);

        int lastParagraphNumber = range.numParagraphs() - 1;
        if (paragraphNumber > lastParagraphNumber) {
            return null;
        }
        return range.getParagraph(paragraphNumber);

    }

    private static HSSFSheet getSheet() {
        return template.getSheetAt(0);

    }

    private static HSSFCell getCell(HSSFSheet sheet, int rowNumber, int columnNumber) {
        System.out.println("Getting cell at row " + rowNumber + " , column " + columnNumber);
        HSSFRow row = sheet.getRow(rowNumber);
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
        final String alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        for (int i = 0; i < 6; i++) {

            if (i==0) { // single-letter column
                for (int j = 0; j < 26; j++) {
                    lettersToNumbers.put(String.valueOf(alphabet.charAt(j)), j);

                }
            }

            else {
                for (int j = 0; j < 26; j++) {
                    lettersToNumbers.put(
                            String.valueOf(alphabet.charAt(i-1)) + String.valueOf(alphabet.charAt(j)), 26*i+j);

                    // when i = 1, column #s 26-51, or (26*1)+j
                    // i = 2, column #s 52-77, or (26*2)+j
                    // etc

                }

            }

        }




    }
}
