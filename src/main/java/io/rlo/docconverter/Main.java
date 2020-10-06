package io.rlo.docconverter;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.net.URI;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class Main {

    public static void main(String[] args) {
        List<String> officeFileExtensions = new ArrayList<>();
        officeFileExtensions.add(".docx");
        officeFileExtensions.add(".xlsx");
        officeFileExtensions.add(".pptx");

        // Make sure both args are present
        if(args.length < 2) {
            System.out.println("Usage: docconverter <input file> <output dir>");
            System.exit(1);
        }

        // Get the input file and output directory, replacing separators for interoperability
        Path inputFile = Paths.get(args[0]);
        Path outputDir = Paths.get(args[1]);

        // Get the input file
        File file = new File(inputFile.toUri());

        // Get the name of the input file
        String fileName = file.getName();

        // Make sure the file is convertible type
        if(!officeFileExtensions.contains(fileName.substring(fileName.lastIndexOf('.')))) {
            System.out.print("Bad file, must be ");
            System.out.println(officeFileExtensions);
            System.exit(2);
        }

        // Remove the extension from the input file name
        fileName = fileName.substring(0, fileName.lastIndexOf('.'));

        // Set the output file to the output directory and the same filename ending in .pdf
        URI outputFile = Paths.get(outputDir + "/" + fileName + ".pdf").toUri();

        System.out.println("Converting file " + fileName);

        // Attempt the conversion
        // Get the document in try-with-resources
        try(XWPFDocument document = new XWPFDocument(new FileInputStream(file))) {
            OutputStream out = new FileOutputStream(new File(outputFile));
            PdfConverter.getInstance().convert(document, out, null);
            System.out.println("PDF file created and output to " + outputFile);
        } catch (IOException x) {
            System.out.println("Failed to convert file: " + x.getMessage());
            x.printStackTrace();
        }

    }

}
