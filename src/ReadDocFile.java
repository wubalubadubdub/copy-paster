import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by bearg on 1/27/2017.
 */
public class ReadDocFile {

    private static final String WORD_DOC_PATH_PREFIX = "C:\\Users\\bearg\\OneDrive\\Documents\\transcriptions\\";
    private static final String TEMPLATE_PATH_PREFIX = "C:\\Users\\bearg\\OneDrive\\Documents\\transcript_docs\\";

    public static void main(String[] args) throws IOException {

        String firstParagraph = getDocumentText();
        HSSFWorkbook template = readFile(TEMPLATE_PATH_PREFIX + "G-1613 AFP Patient Immersion Capture Sheet 1-11-17.xls");
        FileOutputStream stream = new FileOutputStream(TEMPLATE_PATH_PREFIX + "G-1613 AFP Patient Immersion Capture Sheet 1-11-17.xls");
        HSSFSheet sheet = template.getSheetAt(0);

        // takes a 0-based param. if we want row n in the spreadsheet, this param should be n-1
        HSSFRow row = sheet.getRow(3);

        // getCell takes a 0-based param called cellnum that represents a column.
        // e.g. column A is 0, B is 1, etc.
        HSSFCell cell = row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); // returning null for some reason
        cell.setCellValue(firstParagraph);

        HSSFFont font = template.createFont();
        font.setBold(true);
        HSSFRichTextString paragraphWithCorrectBolding = new HSSFRichTextString(firstParagraph);
        paragraphWithCorrectBolding.applyFont(0, 10, font);
        cell.setCellValue(paragraphWithCorrectBolding);

        template.write(stream);
        stream.close();
        template.close();

    }

    private static String getDocumentText() throws IOException {
        final File file;
        final WordExtractor extractor;
        FileInputStream fis = null;

        try {
            file = new File(WORD_DOC_PATH_PREFIX + "31 G-1613 AFP Patient Immersion - Phase 2 122016 12pm BC.doc");
            fis = new FileInputStream(file);
            HWPFDocument document = new HWPFDocument(fis);
            extractor = new WordExtractor(document);
            String[] paragraphs = extractor.getParagraphText();
            return paragraphs[0];


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
