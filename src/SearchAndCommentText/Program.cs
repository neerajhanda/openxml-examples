using System.Text.RegularExpressions;
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
            var paraInnerText = para.InnerText;
            var matches = Regex.Matches(paraInnerText, searchTerm);
            foreach (Match match in matches)
            {
                var startIndex = match.Index;
                var endIndex = startIndex + match.Length;
                var runs = para.Descendants<Run>()!.ToArray();

                int i = 0;
                for (; i < runs.Length; i++)
                {
                    int runLength = runs[i].InnerText.Length;
                    if (startIndex < runLength)
                        break;
                    startIndex -= runLength;
                    endIndex -= runLength;
                }

                var searchTermRun = new Run(new Text(searchTerm) { Space = SpaceProcessingModeValues.Preserve });

                var currentRun = runs[i];
                currentRun.InsertBeforeSelf(searchTermRun);
                
                if (startIndex > 0)
                {
                    var beforeText = currentRun.InnerText.Substring(0, startIndex);
                    var beforeTextRun = new Run(new Text(beforeText)
                        { Space = SpaceProcessingModeValues.Preserve });
                    searchTermRun.InsertBeforeSelf(beforeTextRun);
                }

                if (startIndex + searchTerm.Length < currentRun.InnerText.Length)
                {
                    var afterText = currentRun.InnerText.Substring(startIndex + searchTerm.Length,
                        currentRun.InnerText.Length - (startIndex + searchTerm.Length));
                    var afterTextRun = new Run(new Text(afterText) { Space = SpaceProcessingModeValues.Preserve });
                    searchTermRun.InsertAfterSelf(afterTextRun);
                    currentRun.Remove();
                }
                else
                {
                    while (endIndex > currentRun.InnerText.Length)
                    {
                        endIndex -= currentRun.InnerText.Length;
                        currentRun.Remove();
                        currentRun = runs[++i];
                    }
                    var afterText = currentRun.InnerText.Substring(endIndex,
                        currentRun.InnerText.Length - endIndex);
                    var afterTextRun = new Run(new Text(afterText) { Space = SpaceProcessingModeValues.Preserve });
                    searchTermRun.InsertAfterSelf(afterTextRun);
                    currentRun.Remove();
                }
                
                Comments? comments;
                comments = GetCommentsPart(wordprocessingDocument);
                int id = GetNextId(comments);

                string comment = $"Found {searchTerm}";
                InsertComment(comment, id, "OpenXml.Examples", "OE", comments, para, searchTermRun);
            }

            // foreach (var run in para?.Descendants<Run>()!)
            // {
            //     var text = run.GetFirstChild<Text>();
            //     if (text != null)
            //     {
            //         int index = text.Text.IndexOf(searchTerm, StringComparison.InvariantCultureIgnoreCase);
            //         if (index >= 0)
            //         {
            //             var searchTermRun = SplitRun(run, index, searchTerm.Length);
            //             Comments? comments;
            //             comments = GetCommentsPart(wordprocessingDocument);
            //             int id = GetNextId(comments);
            //
            //             string comment = $"Found {searchTerm}";
            //             InsertComment(comment, id, "OpenXml.Examples", "OE", comments, para, searchTermRun);
            //         }
            //     }
            // }
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
        if (wordprocessingDocument.MainDocumentPart != null && wordprocessingDocument.MainDocumentPart
                .GetPartsOfType<WordprocessingCommentsPart>().Any())
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