import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;

import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by bearg on 1/30/2017.
 */
public class DocAnalyzer {

    private static final String REGEX = "[A-Z]+([0-9]).*";
    static final String PATH_PREFIX = "C:\\Users\\bearg\\OneDrive\\Documents\\transcriptions\\";

    private static FileInputStream inputStream;
    static HWPFDocument wordDocument;
    static String wordDocumentName;
    static int fontSize;
    static String fontName;
    static int rowNumber;
    private static boolean isRowTemplate;


    public static void main(String[] args) {
        if (args.length != 1) {
            System.out.println("Must provide the Word document filename as an argument");
            System.exit(0);
        }

        wordDocumentName = args[0];
        analyzeDoc();
    }

    private static void analyzeDoc() {
        try {
            File wordDocFile = new File(PATH_PREFIX + wordDocumentName);
            inputStream = new FileInputStream(wordDocFile);
            wordDocument = new HWPFDocument(inputStream);

        } catch (IOException e) {
            e.printStackTrace();
        }

        Range range = wordDocument.getRange(); // an object that contains the entire text of the document
        isRowTemplate = checkIfRowTemplate(range); // true if row template, false if column template
        Paragraph firstParagraph = range.getParagraph(0); // first paragraph used to determine font size and type

        // not sure yet why fontSize needs to be divided by 2, but the font size was being detected as 22 when
        // it should've been 11.
        fontSize = (firstParagraph.getCharacterRun(0).getFontSize()) / 2;
        fontName =  firstParagraph.getCharacterRun(0).getFontName();

        if (isRowTemplate) {
            RowCopyPaster.run();
        }

        else {
            throw new IllegalStateException("Row template was not detected correctly");
        }


    }

    private static String getIdentifier(Paragraph paragraph) {
        String plaintext = paragraph.text();
        int colonIndex = plaintext.indexOf(':');
        return plaintext.substring(0, colonIndex + 1); // e.g. N4(1):
    }

    private static boolean checkIfRowTemplate(Range range) {
        String firstParagraphIdentifier = getIdentifier(range.getParagraph(0));
        int i = 1;
        while (range.getParagraph(i).text().equals("\r")) {
            i++;
        }
        String secondParagraphIdentifier = getIdentifier(range.getParagraph(i));

        Pattern pattern = Pattern.compile(REGEX);
        Matcher first = pattern.matcher(firstParagraphIdentifier);
        Matcher second = pattern.matcher(secondParagraphIdentifier);

        if (!(first.find() && second.find())) {
            throw new IllegalStateException("The regex did not match the identifier(s). Check the document for malformed" +
                    " identifiers");
        }

        String firstMatch = first.group(1); // group 1 refers to the digit captured with () used in the REGEX
        String secondMatch = second.group(1);

        if (Integer.parseInt(firstMatch) == Integer.parseInt(secondMatch)) { // e.g. A4(0), B4(0); Z3(0), AA3(0); AC4(1), AD4(1)
            rowNumber = Integer.parseInt(firstMatch) - 1; // using the same row number in the document as in the template
            // makes it easier to do the hand-in copy of the doc file in many cases (only need to delete the (#) part
            return true;
        }

        return false;


    }
}
