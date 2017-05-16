
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by bearg on 1/27/2017.
 * Answers go horizontally in row workbook
 */
public class RowCopyPaster {

    private static Workbook workbook;
    private static FileOutputStream outputStream;
    private static Set<Integer> sheetNumbersUsed;


   static void run()

    {

        final String excelDocumentName = DocAnalyzer.PATH_PREFIX + DocAnalyzer.wordDocumentName.replace(".doc", ".xlsx"); // save as .xls file after program has run if needed
        try {
            workbook = readFile(excelDocumentName);
            outputStream = new FileOutputStream(excelDocumentName);
            System.out.println("Number of cell styles is " + workbook.getNumCellStyles());
            paragraphLoop();
            putXsInBlankCells();
            tearDown();

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private static void tearDown() throws IOException {
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
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
            RichTextString rts = getBoldParagraph(currentParagraph);

            int sheetNumber = getSheetNumberFromParagraph(plainTextParagraph);
            sheetNumbersUsed.add(sheetNumber);

            // must call before stripping identifier
            Sheet currentSheet;
            if (DocAnalyzer.defaultRowAndSheetSet) {
                currentSheet = workbook.getSheetAt(0);
            }

            else {
                currentSheet = workbook.getSheetAt(sheetNumber);
            }

            currentSheet.setDisplayGridlines(true);
            Cell currentCell = getCell(currentSheet, DocAnalyzer.rowNumber, columnNumber);
            pasteTextIntoCell(currentCell, rts);
            loopCounter++;
            System.out.println("Text pasted into cell " + loopCounter + " times");

        }

    }

    private static int getColumnNumberFromParagraph(String plainTextParagraph) {

        Pattern pattern = Pattern.compile(DocAnalyzer.COLUMN_REGEX);
        Matcher columnMatcher = pattern.matcher(plainTextParagraph.substring(0,6).trim()); // e.g. AA3(1): or AA: The
        // when command line option D is given
        if (!columnMatcher.find()) {
            throw new IllegalStateException("Couldn't match the COLUMN_REGEX to get the column number");
        }
        String columnIdentifier = columnMatcher.group(1);
        return DocAnalyzer.lettersToNumbers.get(columnIdentifier);
    }


    private static int getSheetNumberFromParagraph(String plainTextParagraph) {

        // below code can be bypassed with command line arg "D" for the common case where we are pasting into only one sheet
       if (DocAnalyzer.defaultRowAndSheetSet) {
           return 0;
       }

        Pattern pattern = Pattern.compile(DocAnalyzer.SHEET_REGEX);
        String identifier = plainTextParagraph.substring(0, 6).trim(); // e.g. AB4(0):
        Matcher identifierMatcher = pattern.matcher(identifier);
        if (!identifierMatcher.find()) {
            throw new IllegalStateException("Could not get sheet number from paragraph. The regex was not matched.");
        }

        String matched = identifierMatcher.group(1);
        return Integer.parseInt(matched);


    }

    private static Font getBold() {
        CellStyle cellStyle = workbook.createCellStyle();
        Font bold = workbook.createFont();
        bold.setBold(true);
        bold.setFontName(DocAnalyzer.fontName);
        bold.setFontHeightInPoints((short) DocAnalyzer.fontSize);
        cellStyle.setFont(bold);
        return bold;

    }

    private static Font getFontWithCorrectNameAndSize() {
       Font font = workbook.createFont();
       font.setFontName(DocAnalyzer.fontName);
       font.setFontHeightInPoints((short) DocAnalyzer.fontSize);
       return font;
    }

    private static RichTextString getBoldParagraph(Paragraph currentParagraph) {
        CreationHelper helper = workbook.getCreationHelper();
        if (currentParagraph.text().equals("\r")) {

            return helper.createRichTextString(currentParagraph.text());
        }

        String plainTextNoIdentifier = stripIdentifier(currentParagraph.text()).trim();
        RichTextString rts = helper.createRichTextString(plainTextNoIdentifier);
        ArrayList<Integer> indicesUsed = new ArrayList<>();


        Matcher identifierMatcher = DocAnalyzer.identifierPattern.matcher(currentParagraph.text());
        if (!identifierMatcher.find()) {
            throw new IllegalStateException("Identifier was not found within this paragraph");
        }
        String identifier = identifierMatcher.group(0);
        currentParagraph.replaceText(identifier, "");
        int numCharacterRuns = currentParagraph.numCharacterRuns();

        for (int i=0; i < numCharacterRuns; i++) {

            // get character runs one at a time
            CharacterRun characterRun = currentParagraph.getCharacterRun(i);

            // find that character run substring in the entire text
            // ignore character runs until after the identifier
            String textToMatch = characterRun.text().trim();

            if (textToMatch.isEmpty() | textToMatch.equals("\r")) {
                continue;
            }

            // TODO add check here for multiple occurences of same bolded string
            // need to use "fromIndex" 2nd param of indexOf method

            int startIndex = plainTextNoIdentifier.indexOf(textToMatch);
            if (indicesUsed.contains(startIndex)) {
                int fromIndex = indicesUsed.get(indicesUsed.size() - 1) + 1;
                startIndex = plainTextNoIdentifier.indexOf(textToMatch, fromIndex);
            }

            indicesUsed.add(startIndex);
            int endIndex = startIndex + textToMatch.length();

            // if the run of characters is bolded
            if (characterRun.isBold()) {

                // apply bold font to that substring
                rts.applyFont(startIndex, endIndex, getBold()); // this is applying the font only to the bold portion
                // of text, from startBoldIndex to endBoldIndex. we need to apply the font size and name to the entire
                // rts string, though

                // use getbold, which has correct font name and size, and apply that to only the substrings that
                // should be bolded. use getfont, which has correct font name and size but no bolding, and apply that
                // to the rest of the text

            } else {

                rts.applyFont(startIndex, endIndex, getFontWithCorrectNameAndSize());
            }

        }



        return rts;
    }


    private static String stripIdentifier(String plainText) {

        Pattern identifierPattern = Pattern.compile(DocAnalyzer.IDENTIFIER_REGEX);
        Matcher identifierMatcher = identifierPattern.matcher(plainText);

        if (!identifierMatcher.find()) {
            throw new IllegalStateException("Couldn't strip identifier from paragraph. Did not match regex");
        }

        plainText = plainText.replaceFirst(DocAnalyzer.IDENTIFIER_REGEX, "").trim();
        return plainText;

    }


    private static void pasteTextIntoCell(Cell cell, RichTextString rts){

        cell.setCellValue(rts);
        setCellStyle(cell);
    }

    private static void setCellStyle(Cell cell) {

        CellStyle style = workbook.createCellStyle();

        // default alignment is top left, but this can be manually given as argument TL
        if (DocAnalyzer.cellAlignment == null) {
            style.setVerticalAlignment(VerticalAlignment.TOP);
            style.setAlignment(HorizontalAlignment.LEFT);
        }

        else { // should be one of TL, TC, TR, ML, MC, MR, BL, BC, BR
            switch (DocAnalyzer.cellAlignment) {
                case "TL":
                    style.setVerticalAlignment(VerticalAlignment.TOP);
                    style.setAlignment(HorizontalAlignment.LEFT);
                    break;

                case "TC":
                    style.setVerticalAlignment(VerticalAlignment.TOP);
                    style.setAlignment(HorizontalAlignment.CENTER);
                    break;

                case "TR":
                    style.setVerticalAlignment(VerticalAlignment.TOP);
                    style.setAlignment(HorizontalAlignment.RIGHT);
                    break;

                case "ML":
                    style.setVerticalAlignment(VerticalAlignment.CENTER);
                    style.setAlignment(HorizontalAlignment.LEFT);
                    break;

                case "MC":
                    style.setVerticalAlignment(VerticalAlignment.CENTER);
                    style.setAlignment(HorizontalAlignment.CENTER);
                    break;

                case "MR":
                    style.setVerticalAlignment(VerticalAlignment.CENTER);
                    style.setAlignment(HorizontalAlignment.RIGHT);
                    break;

                case "BL":
                    style.setVerticalAlignment(VerticalAlignment.BOTTOM);
                    style.setAlignment(HorizontalAlignment.LEFT);
                    break;

                case "BC":
                    style.setVerticalAlignment(VerticalAlignment.BOTTOM);
                    style.setAlignment(HorizontalAlignment.CENTER);
                    break;

                case "BR":
                    style.setVerticalAlignment(VerticalAlignment.BOTTOM);
                    style.setAlignment(HorizontalAlignment.RIGHT);
                    break;

                default:
                    throw new IllegalArgumentException("Invalid cell alignment argument. Must be one of " +
                            "TL, TC, TR, ML, MC, MR, BL, BC, BR");
            }
        }

        style.setWrapText(true);
        style.setFillForegroundColor(IndexedColors.AUTOMATIC.getIndex());
        cell.setCellStyle(style);

    }

    private static int getLastFilledColumnNumber(Sheet currentSheet) {

        // want to put Xs in cell range from first column, 0, in rowNumber until the cell directly
        // above is blank, but only if those cells are blank after pasting in text of all paragraphs

        // get the row above the one we're pasting text into
        int rowIndex = 0;
        Row row = currentSheet.getRow(rowIndex);
        int endColumnNumber = 0;
        while (true) { // check cell contents of row above column number
            // check if the cell above is empty or null
            if (row.getCell(endColumnNumber) == null)
            {
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
            Sheet currentSheet = workbook.getSheetAt(sheetNumber);
            Row row = currentSheet.getRow(DocAnalyzer.rowNumber);
            int endColumnNumber = getLastFilledColumnNumber(currentSheet);
            for (int column = 0; column < endColumnNumber; column++) {
                Cell cell = row.getCell(column, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (cell.getStringCellValue().equals("")) { // if the cell is empty
                    cell.setCellValue("X");
                    setCellStyle(cell);
                }
                
            }

        }

    }


    // get a paragraph from the word document
    private static Paragraph getCurrentParagraph(int paragraphNumber) throws IOException {

        Range range = DocAnalyzer.wordDocument.getRange();

        int lastParagraphNumber = range.numParagraphs() - 1;
        if (paragraphNumber > lastParagraphNumber) {
            return null;
        }
        return range.getParagraph(paragraphNumber);

    }


    private static Cell getCell(Sheet sheet, int rowNumber, int columnNumber) {
        System.out.println("Getting cell at row " + rowNumber + " , column " + columnNumber);
        Row row = sheet.getRow(rowNumber);
        if (row == null) {
            throw new IllegalStateException("Trying to get the row in this sheet returned null." +
                    " Is the entire sheet blank? If so, type some text into any cell in that row and " +
                    "run the program again. Also check to see if the first couple columns/rows of the workbook" +
                    " are blank, and if so, put a placeholder letter in them before running program.");
        }
        return row.getCell(columnNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);


    }


    private static Workbook readFile(String filename) throws IOException {
        FileInputStream fis = new FileInputStream(filename);
        try {
            return WorkbookFactory.create(fis);
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } finally {
            fis.close();
        }

        return null;
    }

}
