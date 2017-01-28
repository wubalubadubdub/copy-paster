import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;

/**
 * Created by bearg on 1/27/2017.
 */

public class TextCopier {

    private static final String TEST_PARAGRAPH = "So to get started, while I’m trying to get the joinme up—just to begin—and again, keeping names of institutions and stuff out of our discussion, for confidentiality purposes—to begin, what’s your role within your institution, as it relates to COPD population? And then, tell me a little bit about the institution itself, as well. It’s a large, academic medical center health system. 18 affiliated hospitals of varying size [CROSSTALK] location, and a few—certainly, good-sized hospitals, and several smaller, critical access hospitals in a regional area. And so, I serve as the Pharmacotherapy Director. I sit on a system-wide formulary—committee—that advises on COPD and other areas—respiratory diseases.";

    StringSelection selection = new StringSelection(TEST_PARAGRAPH);
    Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();

    public void setClipboardContents() {
        clipboard.setContents(selection, selection);
    }

    /*public static void main(String[] args) {
        TextCopier textCopier = new TextCopier();
        textCopier.setClipboardContents();
    }*/

}
