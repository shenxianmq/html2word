import com.aspose.words.DocumentVisitor;
import com.aspose.words.FieldEnd;
import com.aspose.words.FieldSeparator;
import com.aspose.words.FieldStart;
import com.aspose.words.Font;
import com.aspose.words.Footnote;
import com.aspose.words.FormField;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import com.aspose.words.SpecialChar;
import com.aspose.words.VisitorAction;

public class FontChanger extends DocumentVisitor
{
    ///
/// Called when a FieldEnd node is encountered in the document.
///
    public int VisitFieldEnd(final FieldEnd fieldEnd)
    {
        //Simply change font name
        ResetFont(fieldEnd.getFont());
        return VisitorAction.CONTINUE;
    }

    ///
/// Called when a FieldSeparator node is encountered in the document.
///
    public int VisitFieldSeparator(final FieldSeparator fieldSeparator)
    {
        ResetFont(fieldSeparator.getFont());
        return VisitorAction.CONTINUE;
    }

    ///
/// Called when a FieldStart node is encountered in the document.
///
    public int VisitFieldStart(final FieldStart fieldStart)
    {
        ResetFont(fieldStart.getFont());
        return VisitorAction.CONTINUE;
    }

    ///
/// Called when a Footnote end is encountered in the document.
///
    public int VisitFootnoteEnd(final Footnote footnote)
    {
        ResetFont(footnote.getFont());
        return VisitorAction.CONTINUE;
    }

    ///
/// Called when a FormField node is encountered in the document.
///
    public int VisitFormField(final FormField formField)
    {
        ResetFont(formField.getFont());
        return VisitorAction.CONTINUE;
    }

    ///
/// Called when a Paragraph end is encountered in the document.
///
    public int VisitParagraphEnd(final Paragraph paragraph)
    {
        ResetFont(paragraph.getParagraphBreakFont());
        return VisitorAction.CONTINUE;
    }

    ///
/// Called when a Run node is encountered in the document.
///
    public int visitRun(final Run run)
    {
        ResetFont(run.getFont());
        return VisitorAction.CONTINUE;
    }

    ///
/// Called when a SpecialChar is encountered in the document.
///
    public int VisitSpecialChar(final SpecialChar specialChar)
    {
        ResetFont(specialChar.getFont());
        return VisitorAction.CONTINUE;
    }

    private void ResetFont(Font font)
    {
        font.setName(mNewFontName);
        font.setSize(mNewFontSize);

    }

    private String mNewFontName = "宋体";
    private double mNewFontSize = 10.5;
}
