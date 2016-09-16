package au.gov.aims.html2word;

import org.apache.log4j.Logger;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.LinkedList;
import java.util.List;

/**
 * Created by gcoleman on 17/06/2016.
 */
public class Main {


    private static Logger log = Logger.getLogger(Main.class);
    public static void main(String[] argv) throws IOException, Docx4JException {

        Boolean error = false;
        try {

            if (argv.length != 3) {
                log.error("USAGE: java -jar html2word.jar file <input.html> <output.docx>\n or" +
                        "java -jar html2word.jar folder <inputfolder> <outputfolder>");
                System.exit(1);
            }

            if (argv[0].equalsIgnoreCase("folder")) {
                File folder = new File(argv[1]);
                File outputFolder= new File(argv[2]);
                if (!folder.isDirectory()) {
                    throw new RuntimeException("input folder " + argv[1] + " does not exist.");
                }
                if (!outputFolder.isDirectory()) {
                    throw new RuntimeException("output folder " + argv[2] + " does not exist.");
                }

                for (File html: listHtmlFilesRecursive(null, folder)) {
                    try {
                        File docx = new File(outputFolder, removeExtension(html.getName()) + ".docx");
                        generate(html, docx);
                    } catch (Exception e) {
                        log.error(e);
                        error = true;
                    }
                }


            } else {
                File html = new File(argv[1]);
                if (!html.exists()) {
                    throw new RuntimeException(argv[0] + " does not exist.");
                }
                File docx = new File(argv[2]);
                Main.generate(html, docx);
            }

            log.info("SUCCESS!");
            if (error) {
                System.exit(1);
            } else {
                System.exit(0);
            }
        } catch (Throwable e) {
            e.printStackTrace();
            log.error(e);
            System.exit(1);
        }
    }

    public static List<File> listHtmlFilesRecursive(List<File> files, File dir)
    {
        if (files == null)
            files = new LinkedList<File>();

        if (!dir.isDirectory())
        {
            if ((dir.getName().toLowerCase().endsWith(".html")) || (dir.getName().toLowerCase().endsWith(".xhtml"))) {
                files.add(dir);
            }
            return files;
        }

        for (File file : dir.listFiles())
            listHtmlFilesRecursive(files, file);
        return files;
    }

    private static String removeExtension(String name) {
        return name.replaceFirst("[.][^.]+$", "");
    }

    public static void generate(File inputFile, File outputFile) {
        InputStream templateStream = null;
        try {
            // Get the template input stream from the application resources.
            final URL resource = inputFile.toURI().toURL();

            // Instanciate the Docx4j objects.
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
            XHTMLImporterImpl XHTMLImporter = new XHTMLImporterImpl(wordMLPackage);

            // Load the XHTML document.
            wordMLPackage.getMainDocumentPart().getContent().addAll(XHTMLImporter.convert(resource));

            // Save it as a DOCX document on disc.
            wordMLPackage.save(outputFile);
//            Desktop.getDesktop().open(outputFile);

        } catch (Exception e) {
            throw new RuntimeException("Error converting file " + inputFile, e);

        } finally {
            if (templateStream != null) {
                try {
                    templateStream.close();
                } catch (Exception ex) {
                    log.error("Can not close the input stream.", ex);

                }
            }
        }
    }


}
