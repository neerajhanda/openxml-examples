using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SearchAndCommentText;

class Program
{
    static void Main(string[] args)
    {
        string filePath = "sample.docx";
        string searchTerm = "Online Video";

        WordprocessingDocument wordprocessingDocument =
            WordprocessingDocument.Open(filePath, true);

        Body? body = wordprocessingDocument.MainDocumentPart?.Document.Body;
        foreach (var para in body?.Descendants<Paragraph>()!)
        {
            foreach (var run in para?.Descendants<Run>()!)
            {
                var text = run.GetFirstChild<Text>();
                if (text != null)
                {
                    int index = text.Text.IndexOf(searchTerm, StringComparison.InvariantCultureIgnoreCase);
                    if (index >= 0)
                    {
                        var searchTermRun = SplitRun(run, index, searchTerm.Length);
                        Comments? comments;
                        comments = GetCommentsPart(wordprocessingDocument);
                        int id = GetNextId(comments);

                        string comment = $"Found {searchTerm}";
                        InsertComment(comment, id, "OpenXml.Examples", "OE", comments, para, searchTermRun);
                    }
                }
            }
        }
        wordprocessingDocument.Save();
    }

    private static void InsertComment(string comment, int id, string author, string initials, Comments? comments,
        Paragraph para, Run run)
    {
        Paragraph p = new Paragraph(new Run(new Text(comment)));
        string idString = id.ToString();
        Comment cmt = new Comment()
        {
            Id = idString,
            Author = author, Initials = initials, Date = DateTime.Now
        };
        cmt.AppendChild(p);
        comments?.AppendChild(cmt);
        comments?.Save();

        para.InsertBefore(new CommentRangeStart()
            { Id = idString }, run);

        var cmtEnd = para.InsertAfter(new CommentRangeEnd()
            { Id = idString }, run);

        para.InsertAfter(new Run(new CommentReference() { Id = idString }), cmtEnd);
    }

    private static int GetNextId(Comments? comments)
    {
        if (comments is { HasChildren: true })
        {
            return comments.Descendants<Comment>().Select(e => int.Parse(e.Id?.Value)).Max() + 1;
        }

        return 1;
    }

    private static Comments? GetCommentsPart(WordprocessingDocument wordprocessingDocument)
    {
        Comments? comments = null;
        if (wordprocessingDocument.MainDocumentPart != null && wordprocessingDocument.MainDocumentPart.GetPartsOfType<WordprocessingCommentsPart>().Any())
        {
            comments =
                wordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart?.Comments;
        }
        else
        {
            WordprocessingCommentsPart? commentPart =
                wordprocessingDocument.MainDocumentPart?.AddNewPart<WordprocessingCommentsPart>();
            if (commentPart != null)
            {
                commentPart.Comments = new Comments();
                comments = commentPart.Comments;
            }
        }

        return comments;
    }

    static Run SplitRun(Run runToSplit, int indexToSplitAt, int lengthOfSplitTerm)
    {
        var text = runToSplit.GetFirstChild<Text>();
        if (text == null)
        {
            throw new ArgumentException("Run is empty.");
        }

        var beforeText = text.Text.Substring(0, indexToSplitAt);
        var actualText = text.Text.Substring(indexToSplitAt, lengthOfSplitTerm);
        var afterText = text.Text.Substring(indexToSplitAt + lengthOfSplitTerm,
            text.Text.Length - (indexToSplitAt + lengthOfSplitTerm));
        var beforeTextRun = new Run(new Text(beforeText) { Space = SpaceProcessingModeValues.Preserve });

        runToSplit.InsertBeforeSelf(beforeTextRun);
        runToSplit.Remove();
        var actualTextRun = new Run(new Text(actualText) { Space = SpaceProcessingModeValues.Preserve });
        beforeTextRun.InsertAfterSelf(actualTextRun);
        var afterTextRun = new Run(new Text(afterText) { Space = SpaceProcessingModeValues.Preserve });
        actualTextRun.InsertAfterSelf(afterTextRun);
        return actualTextRun;
    }
}