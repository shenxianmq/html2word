import com.aspose.words.*;
import java.io.File;

public class html2word {
    public static void main(String[] args) throws Exception
    {
        if (args.length == 5) {
            SaveAs(args[0], args[1], args[2], args[3],args[4]);
        }
        else{
            SaveAs(args[0], args[1], args[2], args[3],"沈闲的教室");
        }
//        SaveAs("/Users/shenxian/Desktop/test","test.html","/Users/shenxian/Desktop/new.docx","我是沈闲","none");
    }
    public static void SaveAs(String inPutHtmlDir, String inPutHtmlName, String outPutDocPath, String title,String header) throws Exception
    {
        ToWord(inPutHtmlDir, inPutHtmlName, outPutDocPath);
        Document doc = new Document(outPutDocPath);
        FindReplaceOptions options = new FindReplaceOptions();
        doc.getRange().replace(" B．","\t\tB．",options);
        doc.getRange().replace(" C．","\t\tC．",options);
        doc.getRange().replace(" D．","\t\tD．",options);
        FontChanger changer = new FontChanger();
        doc.accept(changer);
        DocumentBuilder builder = new DocumentBuilder(doc);
        Font font = builder.getFont();
        font.setSize(15);
        font.setName("微软雅黑");
        font.setBold(true);
        builder.getCurrentParagraph().getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.writeln(title);
        builder.getCurrentParagraph().getParagraphFormat().setAlignment(ParagraphAlignment.LEFT);
//        直接变回左对齐就行了嘛
        font.setSize(10.5);
        font.setName("宋体");
        font.setBold(false);
        if (!header.equals("none")) {
            builder.getParagraphFormat().setAlignment(ParagraphAlignment.LEFT);
//        builder.getCurrentStory().getFirstParagraph().getParagraphBreakFont().setSize(30);
//        上面之所以这样写是为了将标题下面的文本对齐方式改回左对齐
            builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
            builder.writeln("\n\t\t\t\t\t\t\t\t\t\t" + header);
        }
        for (Section section : doc.getSections()){
            section.getPageSetup().setPaperSize(PaperSize.A4);
        }
        int vLineHeight = 5;
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
        for (Shape shape : (Iterable<Shape>) shapes)
        {
            shape.getFont().setPosition(-(shape.getHeight() / 2 - vLineHeight));
        }

        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FOOTER_PRIMARY);
        doc.getFirstSection().getHeadersFooters().add(footer);
        //页脚段落
        Paragraph footerpara = new Paragraph(doc);
        footerpara.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        Run footerparaRun = new Run(doc);
        footerparaRun.getFont().setName("宋体");
        footerparaRun.getFont().setSize(10.5);//小5号字体
        footerpara.appendChild(footerparaRun);
        footerpara.appendField(FieldType.FIELD_PAGE, true);//当前页码
        footerpara.appendChild(footerparaRun);
        footer.appendChild(footerpara);
        doc.save(outPutDocPath);
//        File.separator是分隔符
        String done_doc = inPutHtmlDir + File.separator + "done.docx";
        Document doc_done = new Document();
        doc_done.save(done_doc, SaveFormat.DOCX);
    }
    static void ToWord(String dataDir, String dataName, String outPutDocPath) throws Exception
    {
        // Load the document from disk.
        String htmlPath = dataDir + "/" + dataName;
        Document doc = new Document(htmlPath);
        doc.save(outPutDocPath, SaveFormat.DOCX); //保存为docx
    }

}