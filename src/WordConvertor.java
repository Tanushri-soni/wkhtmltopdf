import java.io.IOException;

import javax.xml.bind.JAXBException;

import org.docx4j.openpackaging.exceptions.Docx4JException;

public class WordConvertor {

	public static void main(String[] args) throws InterruptedException, IOException, Docx4JException, JAXBException {

		StreamBasedWordConvertor.wordCOnvertor();
	}

}
