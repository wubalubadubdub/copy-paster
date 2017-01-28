import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.ArrayList;

/**
 * Created by bearg on 1/27/2017.
 */
public class ReadDocFile {

    private static final String WORD_DOC_PATH_PREFIX = "C:\\Users\\bearg\\OneDrive\\Documents\\transcriptions\\";
    private static final String TEMPLATE_PATH_PREFIX = "C:\\Users\\bearg\\OneDrive\\Documents\\transcript_docs\\";

    public static void main(String[] args) throws IOException {

        Paragraph firstParagraph = getDocumentText();
        HSSFWorkbook template = readFile(TEMPLATE_PATH_PREFIX + "G-1613 AFP Patient Immersion Capture Sheet 1-11-17.xls");
        FileOutputStream stream = new FileOutputStream(TEMPLATE_PATH_PREFIX + "G-1613 AFP Patient Immersion Capture Sheet 1-11-17.xls");
        HSSFSheet sheet = template.getSheetAt(0);

        // takes a 0-based param. if we want row n in the spreadsheet, this param should be n-1
        HSSFRow row = sheet.getRow(3);

        // getCell takes a 0-based param called cellnum that represents a column.
        // e.g. column A is 0, B is 1, etc.
        HSSFCell cell = row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

        HSSFFont font = template.createFont();
        font.setBold(true);

        String formattedText = "";
        HSSFRichTextString rts;

        for (int i = 0; i < firstParagraph.numCharacterRuns(); i++) {

            // get character runs one at a time
            CharacterRun characterRun = firstParagraph.getCharacterRun(i);
            String characterRunString = characterRun.text();
            rts = new HSSFRichTextString(characterRunString);


            // if the run of characters is bolded
            if (characterRun.isBold()) {

                // apply bold font to that string
                rts.applyFont(font);
            }


            formattedText += rts.toString();


        }


        cell.setCellValue(formattedText);

        template.write(stream);
        stream.close();
        template.close();

    }

    private static Paragraph getDocumentText() throws IOException {
        final File file;
        FileInputStream fis = null;

        try {
            file = new File(WORD_DOC_PATH_PREFIX + "31 G-1613 AFP Patient Immersion - Phase 2 122016 12pm BC.doc");
            fis = new FileInputStream(file);
            HWPFDocument document = new HWPFDocument(fis);
            Range range = document.getRange();
            Paragraph firstParagraph = range.getParagraph(0);
           /* CharacterRun characterRun = paragraph.getCharacterRun(2);
            System.out.println(characterRun.text());
            System.out.println(characterRun.isBold());*/

           return firstParagraph;


        } catch (java.io.IOException e) {
            e.printStackTrace();
        }

        finally {
            fis.close();
        }

        return null;
    }

    private static void pasteTextIntoSpreadsheet(){

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
}
