package org.doctribute.tools;

import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class FOPToPowerPointConverterApp {

    public static void main(String[] args) throws Exception {

        if (args.length != 1 || !args[0].endsWith(".xml")) {
            System.out.println("Usage:\n");
            System.out.println("java -jar fop-to-powerpoint-converter.jar <fop-intermediate-format.xml>\n");

        } else {
            Path fopIFPath = Paths.get(args[0]);
            Path pptxPath = Paths.get(args[0].replace(".xml", ".pptx"));

            try (
                    InputStream inputStream = Files.newInputStream(fopIFPath);
                    OutputStream outputStream = Files.newOutputStream(pptxPath)) {

                FOPToPowerPointConverter.convert(inputStream, outputStream);
            }
        }
    }

}
