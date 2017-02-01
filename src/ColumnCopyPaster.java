import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by bearg on 1/31/2017.
 */
public class ColumnCopyPaster {

    private static HSSFWorkbook template;
    private static FileOutputStream outputStream;
    private static Set<Integer> sheetNumbersUsed;

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
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setAlignment(HorizontalAlignment.LEFT);
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

        // get the column left of the one we're pasting into
        int columnIndex = DocAnalyzer.columnNumber - 1;
        int rowIndex = 0;
        while (true) { // check cell contents of cells left of column number
            // keep looking until we have found three contiguous empty cells in a column
            HSSFRow rowOne = currentSheet.getRow(rowIndex);
            HSSFRow rowTwo = currentSheet.getRow(rowIndex + 1);
            HSSFRow rowThree = currentSheet.getRow(rowIndex + 2);

            if (rowOne == null) {
                throw new IllegalStateException("Row " + rowIndex + " is null");
            }


            if (rowTwo == null) {
                throw new IllegalStateException("Row " + rowIndex + 1 +  " is null");
            }


            if (rowThree == null) {
                throw new IllegalStateException("Row " + rowIndex + 2 + " is null");
            }

            HSSFCell cellOne = rowOne.getCell(columnIndex);
            HSSFCell cellTwo = rowTwo.getCell(columnIndex);
            HSSFCell cellThree = rowThree.getCell(columnIndex);

            boolean cellOneEmpty = isCellEmpty(cellOne);
            boolean cellTwoEmpty = isCellEmpty(cellTwo);
            boolean cellThreeEmpty = isCellEmpty(cellThree);


           /* if (cellOneEmpty && cellTwoEmpty && cellThreeEmpty) {
                break;
            }
            if (cellOneEmpty && cellTwoEmpty) { // only first 2 of 3 cells were null. can advance counter by 3
                rowIndex += 3;
            }
            else if (!cellOneEmpty && cellThreeEmpty) { // last of 3 cells was null. can advance counter by 2
                rowIndex += 2;
            }

            else {
                rowIndex++;
            }*/

           if (cellOneEmpty) {
               break;
           }

           rowIndex++;
        }

        return rowIndex;
    }

    private static boolean isCellEmpty(HSSFCell cell) {
        if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
            return true;
        }

        if (cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().isEmpty()) {
            return true;
        }

        return false;
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

    private static HSSFFont getBoldText() {
        HSSFCellStyle cellStyle = template.createCellStyle();
        HSSFFont boldFont = template.createFont();
        boldFont.setBold(true);
        boldFont.setFontName(DocAnalyzer.fontName);
        boldFont.setFontHeightInPoints((short) DocAnalyzer.fontSize);
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
