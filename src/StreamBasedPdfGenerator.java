import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.io.IOUtils;


public class StreamBasedPdfGenerator {
  public static void makeAPdfWithStreams() throws InterruptedException, IOException {
        Process wkhtml; // Create uninitialized process

        // Start by setting up file streams
        File destinationFile = new File("D:/POCs/HtmlToPdfConverter/outputTest1.pdf");
        File sourceFile = new File("D:/POCs/HtmlToPdfConverter/sparks_mail_attachment.html");

        FileInputStream fis = new FileInputStream(sourceFile);
        FileOutputStream fos = new FileOutputStream(destinationFile);

        String command = "C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf - -"; // Desired command

        wkhtml = Runtime.getRuntime().exec(command); // Start process

        Thread errThread = new Thread(() -> {
            try {
                IOUtils.copy(wkhtml.getErrorStream(), System.err);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        });
        Thread htmlReadThread = new Thread(() -> {
            try {
                IOUtils.copy(fis, wkhtml.getOutputStream());
                wkhtml.getOutputStream().flush();
                wkhtml.getOutputStream().close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        });
        Thread pdfWriteThread = new Thread(() -> {
            try {
                IOUtils.copy(wkhtml.getInputStream(), fos);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        });

        // Do NOT use Run... it should be clear why, you want them to all be going at the same time.
        errThread.start();
        pdfWriteThread.start();
        htmlReadThread.start();

         // Connect HTML Source Stream to wkhtmltopdf
         // Connect PDF Source Stream from wkhtmltopdf to the Destination file steam

        wkhtml.waitFor(); // Allow process to run
        
        System.out.println("testing");
    }
}