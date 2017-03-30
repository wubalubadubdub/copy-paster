import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
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
 * Created by bearg on 1/31/2017.
 * Answers go vertically in column template
 */
public class ColumnCopyPaster {

    private static HSSFWorkbook template;
    private static FileOutputStream outputStream;
    private static Set<Integer> sheetNumbersUsed;
    private static int highestRowNumber;

    static void run() {
        final String excelDocumentName = DocAnalyzer.PATH_PREFIX + DocAnalyzer.wordDocumentName.replace(".doc", ".xls");
        try {
            template = readFile(excelDocumentName);
            outputStream = new FileOutputStream(excelDocumentName);
            paragraphLoop();
            putXsInBlankCells();
            tearDown();
        }

        catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void tearDown() throws IOException {
        template.write(outputStream);
        outputStream.close();
        template.close();
    }

    private static void paragraphLoop() {
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

            int rowNumber = getRowNumberFromParagraph(plainTextParagraph);
            if (rowNumber > highestRowNumber) {
                highestRowNumber = rowNumber;
            }
            HSSFRichTextString rts = getBoldParagraph(currentParagraph);

            int sheetNumber = getSheetNumberFromParagraph(plainTextParagraph);
            sheetNumbersUsed.add(sheetNumber);

            // must call before stripping identifier
            HSSFSheet currentSheet = template.getSheetAt(sheetNumber);

            HSSFCell currentCell = getCell(currentSheet, rowNumber, DocAnalyzer.columnNumber);
            pasteTextIntoCell(currentCell, rts);
            loopCounter++;
            System.out.println("Text pasted into cell " + loopCounter + " times");

        }


    }

    private static Paragraph getCurrentParagraph(int paragraphNumber) {

        Range range = DocAnalyzer.wordDocument.getRange();
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

    private static void pasteTextIntoCell(HSSFCell cell, HSSFRichTextString rts){

        cell.setCellValue(rts);
        setCellStyle(cell);
    }

    private static void setCellStyle(HSSFCell cell) {

        CellStyle style = template.createCellStyle();

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
        style.setFillForegroundColor(HSSFColor.WHITE.index);
        cell.setCellStyle(style);
    }

    private static void putXsInBlankCells() {

        for (Integer sheetNumber : sheetNumbersUsed) {
            HSSFSheet currentSheet = template.getSheetAt(sheetNumber);
            int endRowNumber = getLastFilledRowNumber(currentSheet);
            for (int rowNumber = 0; rowNumber < endRowNumber; rowNumber++) {
                HSSFRow currentRow = currentSheet.getRow(rowNumber);
                HSSFCell cell = currentRow.getCell(DocAnalyzer.columnNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (cell.getStringCellValue().equals("")) { // if the cell is empty
                    cell.setCellValue("X");
                    setCellStyle(cell);
                }

            }

        }

    }

    private static int getLastFilledRowNumber(HSSFSheet currentSheet) {

        // start looking for empty rows at the next one down from the lowest one we pasted text into
        int rowIndex = highestRowNumber + 1;
        while (true) { // check for empty rows

            HSSFRow row = currentSheet.getRow(rowIndex);
            if (row == null) {
                break;
            }

            rowIndex++;

        }

        return rowIndex;
    }

    private static int getRowNumberFromParagraph(String plainTextParagraph) {

        Pattern rowPattern = Pattern.compile(DocAnalyzer.ROW_REGEX);
        String identifier = plainTextParagraph.substring(0, 6).trim(); // e.g. AB4(0):
        Matcher rowMatcher = rowPattern.matcher(identifier);
        if (!rowMatcher.find()) {
            throw new IllegalStateException("Couldn't match the COLUMN_REGEX to get the column number");
        }
        String rowIdentifier = rowMatcher.group(1); // e.g. the "10" from J10(0):
        return Integer.parseInt(rowIdentifier) - 1; // need to -1 since program is 0-based
    }

    private static HSSFFont getBold() {
        HSSFCellStyle cellStyle = template.createCellStyle();
        HSSFFont bold = template.createFont();
        bold.setBold(true);
        bold.setFontName(DocAnalyzer.fontName);
        bold.setFontHeightInPoints((short) DocAnalyzer.fontSize);
        cellStyle.setFont(bold);
        return bold;

    }

    private static HSSFFont getFontWithCorrectNameAndSize() {
        HSSFFont font = template.createFont();
        font.setFontName(DocAnalyzer.fontName);
        font.setFontHeightInPoints((short) DocAnalyzer.fontSize);
        return font;
    }

    private static HSSFRichTextString getBoldParagraph(Paragraph currentParagraph) {

        if (currentParagraph.text().equals("\r")) {
            return new HSSFRichTextString(currentParagraph.text());
        }

        String plainTextNoIdentifier = stripIdentifier(currentParagraph.text()).trim();
        HSSFRichTextString rts = new HSSFRichTextString(plainTextNoIdentifier);
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

    private static int getSheetNumberFromParagraph(String plainTextParagraph) {

        Pattern pattern = Pattern.compile(DocAnalyzer.SHEET_REGEX);
        String identifier = plainTextParagraph.substring(0, 6).trim(); // e.g. AB4(0):
        Matcher identifierMatcher = pattern.matcher(identifier);
        if (!identifierMatcher.find()) {
            throw new IllegalStateException("Could not get sheet number from paragraph. The regex was not matched.");
        }

        String matched = identifierMatcher.group(1);
        return Integer.parseInt(matched);


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
