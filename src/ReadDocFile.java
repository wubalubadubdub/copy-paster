import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * Created by bearg on 1/27/2017.
 */
public class ReadDocFile {

    private static final String PATH_PREFIX = "C:\\Users\\bearg\\OneDrive\\Documents\\transcriptions\\";

    public static void main(String[] args) {

        getDocumentText();
    }

    private static void getDocumentText(){
        final File file;
        final WordExtractor extractor;
        FileInputStream fis = null;

        try {
            file = new File(PATH_PREFIX + "31 G-1613 AFP Patient Immersion - Phase 2 122016 12pm BC.doc");
            fis = new FileInputStream(file);
            HWPFDocument document = new HWPFDocument(fis);
            extractor = new WordExtractor(document);
            String[] paragraphs = extractor.getParagraphText();
            for (int i = 0; i < paragraphs.length; i++) {
                if (paragraphs[i] != null) {
                    System.out.println(paragraphs[i]);
                }
            }


        } catch (java.io.IOException e) {
            e.printStackTrace();
        }

        finally {
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }

    private static HSSFWorkbook readFile(String filename) throws IOException {
        FileInputStream fis = new FileInputStream(filename);
        return null;

    }
}
