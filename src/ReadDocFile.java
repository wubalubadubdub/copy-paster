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
import java.util.Map;

/**
 * Created by bearg on 1/27/2017.
 */
public class ReadDocFile {

    private static final String WORD_DOC_PATH_PREFIX = "C:\\Users\\bearg\\OneDrive\\Documents\\transcriptions\\";
    private static final String TEMPLATE_PATH_PREFIX = "C:\\Users\\bearg\\OneDrive\\Documents\\transcript_docs\\";
    private static String wordDocumentName;
    private static HWPFDocument wordDocument;

    public static void main(String[] args) throws IOException {

        if (args.length < 2) {
            System.out.println("Must supply word filename as an argument and row # (0-based) from the Excel sheet" +
                    "that text should be pasted into");
            System.exit(0);
        }

        try {
            int rowNumber = Integer.parseInt(args[1]);
            wordDocumentName = args[0];
            File wordDocFile = new File(WORD_DOC_PATH_PREFIX + wordDocumentName);
            FileInputStream fis = new FileInputStream(wordDocFile);
            wordDocument = new HWPFDocument(fis);

            final String excelDocumentName = TEMPLATE_PATH_PREFIX + wordDocumentName.replace(".doc", ".xls");
            HSSFWorkbook template = readFile(excelDocumentName);
            FileOutputStream stream = new FileOutputStream(excelDocumentName);
            HSSFSheet sheet = getSheet(template);

        }

        catch (IOException e){
            e.printStackTrace();
        }



    }

    private static void paragraphLoop() throws IOException {
        int paragraphNumber = 0;
        Paragraph currentParagraph;
        while (getParagraph(paragraphNumber) != null) {
            currentParagraph = getParagraph(paragraphNumber);
            String plainTextParagraph = currentParagraph.text();
            HSSFRichTextString rts = new HSSFRichTextString(plainTextParagraph);


        }

    }

    private static HSSFFont getBoldFont(HSSFWorkbook template) {
        HSSFFont boldFont = template.createFont();
        boldFont.setBold(true);
        return boldFont;

    }


        // getCell takes a 0-based param called cellnum that represents a column.
        // e.g. column A is 0, B is 1, etc.
        HSSFCell cell = row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);



        // we will bold a range of characters that need to be bolded in the paragraph
        // with each run through the loop

        for (int i = 0; i < firstParagraph.numCharacterRuns(); i++) {

            // get character runs one at a time
            CharacterRun characterRun = firstParagraph.getCharacterRun(i);


            // if the run of characters is bolded
            if (characterRun.isBold()) {

                // find that substring in the original
                String textToMatch = characterRun.text();

                // TODO add check here for multiple occurences of same bolded string

                int startBold = plaintextParagraph.indexOf(textToMatch);
                int endBold = startBold + textToMatch.length();


                // apply bold font to that substring
                rts.applyFont(startBold, endBold, font);
            }


        }

        cell.setCellValue(rts);

        template.write(stream);
        stream.close();
        template.close();


    }


    // get a paragraph from the word document
    private static Paragraph getParagraph(int paragraphNumber) throws IOException {

        Range range = wordDocument.getRange();

        if (paragraphNumber >= range.numParagraphs()) {
            return null;
        }
        return range.getParagraph(paragraphNumber);

    }

    private static HSSFSheet getSheet(HSSFWorkbook template) {
        return template.getSheetAt(0);

    }

    private static HSSFCell getCell(HSSFSheet sheet, int rowNumber, int columnNumber) {
        HSSFRow row = sheet.getRow(rowNumber);
        return row.getCell(columnNumber);


    }

    private static int getColumnNumber(String plainTextParagraph) {

    }

    private static void pasteTextIntoCell(HSSFCell cell, HSSFRichTextString rts){
        cell.setCellValue(rts);

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
        Map<String, Integer> lettersToNumbers = new HashMap<>();
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
