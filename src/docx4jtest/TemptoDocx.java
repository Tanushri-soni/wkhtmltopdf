package docx4jtest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;

public class TemptoDocx {

	private WordprocessingMLPackage getTemplate(String name) throws Docx4JException, FileNotFoundException {
		WordprocessingMLPackage template = WordprocessingMLPackage.load(new FileInputStream(new File(name)));
		return template;
	}

	private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
		List<Object> result = new ArrayList<Object>();
		if (obj instanceof JAXBElement)
			obj = ((JAXBElement<?>) obj).getValue();

		if (obj.getClass().equals(toSearch))
			result.add(obj);
		else if (obj instanceof ContentAccessor) {
			List<?> children = ((ContentAccessor) obj).getContent();
			for (Object child : children) {
				result.addAll(getAllElementFromObject(child, toSearch));
			}
		}
		return result;
	}

	private void replacePlaceholder(WordprocessingMLPackage template, String name, String placeholder) {
		List<Object> texts = getAllElementFromObject(template.getMainDocumentPart(), Text.class);

		for (Object text : texts) {
			Text textElement = (Text) text;
			if (textElement.getValue().equals(placeholder)) {
				textElement.setValue(name);
			}
		}
	}

	private void writeDocxToStream(WordprocessingMLPackage template, String target)
			throws IOException, Docx4JException {
		File f = new File(target);
		template.save(f);
	}

	public static void main(String[] args) {
		try {
			TemptoDocx doctodocx = new TemptoDocx();
			WordprocessingMLPackage wordprocessingMLPackage = doctodocx
					.getTemplate("D://Sparx/Word Export POC/SalesProposal_Table.docx");
			doctodocx.replacePlaceholder(wordprocessingMLPackage, "ACG 2009-20x30-SanDiegoCA", "&Account");

			Map<String, String> accountName = new HashMap<String, String>();
			accountName.put("AccountName", "ACG 2009");

			Map<String, String> accountAddress = new HashMap<String, String>();
			accountAddress.put("AccountAddress", "San Diego, CA");

			Map<String, String> projectName = new HashMap<String, String>();
			projectName.put("ProjectName", "20 x 30");

			List<Map<String, String>> list = new ArrayList<>();
			list.add(accountName);
			list.add(accountAddress);
			list.add(projectName);

			for (int i = 0; i < 2; i++) {
				Map<String, String> sectionTitle = new HashMap<String, String>();
				sectionTitle.put("SectionTitle", "Section Title");
				list.add(sectionTitle);

				Map<String, String> repl1 = new HashMap<String, String>();
				repl1.put("SNO", "1");
				repl1.put("ComponentTitle", "title1");
				repl1.put("Amount", "$1205");
				list.add(repl1);

				Map<String, String> repl1_1 = new HashMap<String, String>();
				repl1_1.put("ComponentDescription", "desc1");
				list.add(repl1_1);

				Map<String, String> repl2 = new HashMap<String, String>();
				repl2.put("SNO", "2");
				repl2.put("ComponentTitle", "title2");
				repl2.put("Amount", "$1200");
				list.add(repl2);

				Map<String, String> repl2_1 = new HashMap<String, String>();
				repl2_1.put("ComponentDescription", "desc2");
				list.add(repl2_1);

				Map<String, String> repl3 = new HashMap<String, String>();
				repl3.put("SNO", "3");
				repl3.put("ComponentTitle", "title3");
				repl3.put("Amount", "$1300");
				list.add(repl3);

				Map<String, String> repl3_1 = new HashMap<String, String>();
				repl3_1.put("ComponentDescription",
						"desc3desc3desc3desc3desc3desc3desc3desc3desc3desc3desc3desc3desc3desc3desc3desc3desc3desc3desc3");
				list.add(repl3_1);
			}

			Map<String, String> note = new HashMap<String, String>();
			note.put("NOTE", "NOTE:");

			Map<String, String> noteDesc = new HashMap<String, String>();
			noteDesc.put("NoteDescription",
					"The above estimate does not include labor and service arrangements, graphic production, refurbishment or cleaning of existing fabric panels.  This property transships from NACFC.");

			Map<String, String> projectTotal = new HashMap<String, String>();
			projectTotal.put("ProjectTotal", "PROJECT TOTAL:");
			projectTotal.put("TotalAmount", "$5000");

			list.add(note);
			list.add(noteDesc);
			list.add(projectTotal);

			replaceTable(new String[] { "SNO", "ComponentTitle", "ComponentDescription" }, list, wordprocessingMLPackage);
			//wordprocessingMLPackage.sav

			// wordprocessingMLPackage.save(file);
			doctodocx.writeDocxToStream(wordprocessingMLPackage, "D://Sparx/Word Export POC/SalesProposal_Output.docx");

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private static void replaceTable(String[] placeholders, List<Map<String, String>> textToAdd,
			WordprocessingMLPackage template) throws Docx4JException, JAXBException {
		List<Object> tables = getAllElementFromObject(template.getMainDocumentPart(), Tbl.class);

		// 1. find the table
		Tbl tempTable = getTemplateTable(tables, placeholders[0]);
		List<Object> rows = getAllElementFromObject(tempTable, Tr.class);

		// first row is header, second row is content
		if (rows.size() > 0) {
			// this is our template row
			Tr templateRowAccName = (Tr) rows.get(0);
			Tr templateRowAccAddress = (Tr) rows.get(1);
			Tr templateRowProjName = (Tr) rows.get(2);

			Tr templateRowSection = (Tr) rows.get(4);

			Tr templateRowCompTitle = (Tr) rows.get(6);
			Tr templateRowCompDesc = (Tr) rows.get(7);

			Tr templateRowNote = (Tr) rows.get(9);
			Tr templateRowNoteDesc = (Tr) rows.get(10);
			Tr templateRowProjTot = (Tr) rows.get(11);

			for (Map<String, String> replacements : textToAdd) {
				// 2 and 3 are done in this method
				if (replacements.containsKey("AccountName"))
					addRowToTable(tempTable, templateRowAccName, replacements);
				else if (replacements.containsKey("AccountAddress"))
					addRowToTable(tempTable, templateRowAccAddress, replacements);
				else if (replacements.containsKey("ProjectName"))
					addRowToTable(tempTable, templateRowProjName, replacements);
				else if (replacements.containsKey("SectionTitle"))
					addRowToTable(tempTable, templateRowSection, replacements);
				else if (replacements.containsKey("SNO"))
					addRowToTable(tempTable, templateRowCompTitle, replacements);
				else if (replacements.containsKey("ComponentDescription"))
					addRowToTable(tempTable, templateRowCompDesc, replacements);
				else if (replacements.containsKey("NOTE"))
					addRowToTable(tempTable, templateRowNote, replacements);
				else if (replacements.containsKey("NoteDescription"))
					addRowToTable(tempTable, templateRowNoteDesc, replacements);
				else if (replacements.containsKey("ProjectTotal"))
					addRowToTable(tempTable, templateRowProjTot, replacements);
			}

			// 4. remove the template row
			rows.forEach(row -> tempTable.getContent().remove(row));
			/*
			 * tempTable.getContent().remove(templateRowAccName);
			 * tempTable.getContent().remove(templateRowAccAddress);
			 * tempTable.getContent().remove(templateRowProjName);
			 * tempTable.getContent().remove(templateRowCompTitle);
			 * tempTable.getContent().remove(templateRowCompDesc);
			 * tempTable.getContent().remove(templateRowAccName);
			 * tempTable.getContent().remove(templateRowNote);
			 * tempTable.getContent().remove(templateRowNoteDesc);
			 * tempTable.getContent().remove(templateRowProjTot);
			 */
		}
	}

	private static Tbl getTemplateTable(List<Object> tables, String templateKey) throws Docx4JException, JAXBException {
		for (Iterator<Object> iterator = tables.iterator(); iterator.hasNext();) {
			Object tbl = iterator.next();
			List<?> textElements = getAllElementFromObject(tbl, Text.class);
			for (Object text : textElements) {
				Text textElement = (Text) text;
				if (textElement.getValue() != null && textElement.getValue().equals(templateKey))
					return (Tbl) tbl;
			}
		}
		return null;
	}

	private static void addRowToTable(Tbl reviewtable, Tr templateRow, Map<String, String> replacements) {
		Tr workingRow = (Tr) XmlUtils.deepCopy(templateRow);
		List<?> textElements = getAllElementFromObject(workingRow, Text.class);
		for (Object object : textElements) {
			Text text = (Text) object;
			String replacementValue = (String) replacements.get(text.getValue());
			if (replacementValue != null)
				text.setValue(replacementValue);
		}

		reviewtable.getContent().add(workingRow);
	}

}
