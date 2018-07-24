import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

import javax.xml.bind.JAXBException;

import org.docx4j.Docx4J;
import org.docx4j.Docx4jProperties;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class StreamBasedWordConvertor {

	public static void wordCOnvertor() throws Docx4JException, IOException, JAXBException {
		String xhtml = "<table border=\"1\" cellpadding=\"1\" cellspacing=\"1\" style=\"width:100%;\"><tbody><tr><td>test</td><td>test</td></tr><tr><td>test</td><td>test</td></tr><tr><td>test</td><td>test</td></tr></tbody></table>";

	      //  File sourceFile = new File("D:/Compunnel-13354/eclipse workspace/HtmlToPdfConverter/sparks_mail_attachment.html");

	       // FileInputStream fis = new FileInputStream(sourceFile);
	        
		// File destinationFile = new File("D:/Compunnel-13354/eclipse workspace/HtmlToPdfConverter/outputTest.pdf");

	       // FileOutputStream fos = new FileOutputStream(destinationFile);
	        
	        byte[] encoded = Files.readAllBytes(Paths.get("D:/Sparx/templates/deposit-invoice-pdf-html/deposit_invoice.html"));
	        String html =  new String(encoded, "UTF-8");
	        
		/*File sourceFile = new File("D:/Compunnel-13354/eclipse workspace/HtmlToPdfConverter/sparks_mail_attachment.html");
		FileInputStream fis = new FileInputStream(sourceFile);
		byte[] sources = fis*/
		
		// To docx, with content controls
		/*WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

		XHTMLImporterImpl XHTMLImporter = new XHTMLImporterImpl(wordMLPackage);
		// XHTMLImporter.setDivHandler(new DivToSdt());

		wordMLPackage.getMainDocumentPart().getContent().addAll(XHTMLImporter.convert(html, null));

		System.out.println(XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getJaxbElement(), true, true));

		 wordMLPackage.save(new java.io.File("D:/Compunnel-13354/eclipse workspace/HtmlToPdfConverter" + "/OUT_from_XHTML.docx"));

		// Back to XHTML

		HTMLSettings htmlSettings = Docx4J.createHTMLSettings();
		htmlSettings.setWmlPackage(wordMLPackage);

		// output to an OutputStream.
		OutputStream os = new ByteArrayOutputStream();

		// If you want XHTML output
		Docx4jProperties.setProperty("docx4j.Convert.Out.HTML.OutputMethodXML", true);
		Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);

		System.out.println(((ByteArrayOutputStream) os).toString());*/
		
		/*WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
		AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(new PartName("/hw.html"));
		afiPart.setBinaryData(html.getBytes());
		afiPart.setContentType(new ContentType("text/html"));
		Relationship altChunkRel = wordMLPackage.getMainDocumentPart().addTargetPart(afiPart);
		 
		// .. the bit in document body
		CTAltChunk ac = Context.getWmlObjectFactory().createCTAltChunk();
		ac.setId(altChunkRel.getId() );
		wordMLPackage.getMainDocumentPart().addObject(ac);
		 
		// .. content type
		wordMLPackage.getContentTypeManager().addDefaultContentType("html", "text/html");*/
	        
	        File dir = new File ("D:/Compunnel-13354/eclipse workspace/HtmlToPdfConverter");
	        File actualFile = new File (dir,  "/OUT_from_XHTML_TEST.docx");

	        WordprocessingMLPackage wordMLPackage = null;
	        try
	        {
	            wordMLPackage = WordprocessingMLPackage.createPackage();
	        }
	        catch (InvalidFormatException e)
	        {
	            e.printStackTrace();
	        }


	        XHTMLImporterImpl XHTMLImporter = new XHTMLImporterImpl(wordMLPackage);

	        OutputStream fos = null;
	        try
	        {
	            //fos = new FileOutputStream("D:/Compunnel-13354/eclipse workspace/HtmlToPdfConverter/OUT_from_XHTML_TEST.docx");

	            System.out.println(XmlUtils.marshaltoString(wordMLPackage
	                    .getMainDocumentPart().getJaxbElement(), true, true));

	                       /* HTMLSettings htmlSettings = Docx4J.createHTMLSettings();
	            htmlSettings.setWmlPackage(wordMLPackage);
	  Docx4jProperties.setProperty("docx4j.Convert.Out.HTML.OutputMethodXML",
	                    true);
	            Docx4J.toHTML(htmlSettings, fos, Docx4J.FLAG_EXPORT_PREFER_XSL);*/
	            wordMLPackage.save(actualFile); 
	        }
	        catch (Docx4JException e)
	        {
	            // TODO Auto-generated catch block
	            e.printStackTrace();
	        }
	        finally{
	            try {
	                fos.close();
	            } catch (IOException e) {
	                e.printStackTrace();
	            }
	        }
		//wordMLPackage.save(new java.io.File("D:/Compunnel-13354/eclipse workspace/HtmlToPdfConverter" + "/OUT_from_XHTML.docx"));

		System.out.println("done");
		
	}
}
