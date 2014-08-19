import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.math.BigInteger;
import javax.xml.bind.JAXBException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Style;

public class Generator {

    public static void main(String[] args) {
        if (args.length==1 && (args[0].equals("-h") || args[0].equals("--help"))) {
            printHelp();
            System.exit();
        } else if (args.length != 2) {
            System.out.println("Invalid arguments.");
            printHelp();
            System.exit(0);
        }
        
        String fileIn = args[0];
        String fileOut = args[1];
        
        Generator g = new Generator();
        try {
            g.generate(fileIn, fileOut);
        } catch (JAXBException | Docx4JException e) {
            throw new RuntimeException(e);
        } catch (FileNotFoundException e) {
            System.out.println("File not found " + fileIn);
        } catch (IOException e) {
            System.out.println("Failed to read file: " + e);
        }
    }

    static void printHelp() {
        System.out.println("Usage:");
        System.out.println("    text2docx [input file] [output file]");
        System.out.println("    text2docx --help");
        System.out.println();
    }
    
    WordprocessingMLPackage wordDoc;
    
    void generate(String in, String out) throws JAXBException, Docx4JException, IOException {
        BufferedReader r = new BufferedReader(new FileReader(in));
        
        wordDoc = WordprocessingMLPackage.createPackage();
        
        wordDoc.getMainDocumentPart().getStyleDefinitionsPart().getJaxbElement().getStyle().stream().forEach((Style s) -> {
            if (s.getStyleId().equals("Normal"))
                setStyleMLA(s);
        });
        
        MainDocumentPart mdp = wordDoc.getMainDocumentPart();
        
        String line;
        while((line = r.readLine()) != null)
            mdp.addParagraphOfText(line);
        
        wordDoc.save(new File(out));
        System.out.println("Saved " + out);

    }
    
    void setStyleMLA(Style style) {
        RPr rpr = new RPr();
        changeFont(rpr, "Times New Roman");
        changeFontSize(rpr, 12*2);
        style.setRPr(rpr);
    }
    
    RPr removeInfo(Style style) {
        RPr rpr = style.getRPr();
        rpr.getRFonts().setAsciiTheme(null);
        rpr.getRFonts().setHAnsiTheme(null);
        return rpr;
    }
    
    void changeFont(RPr rp, String font) {
        RFonts rf = new RFonts();
        rf.setAscii(font);
        rf.setHAnsi(font);
        rp.setRFonts(rf);
    }
    
    void changeFontSize(RPr rp, int fSize) {
        HpsMeasure size = new HpsMeasure();
        size.setVal(BigInteger.valueOf(fSize));
        rp.setSz(size);
    }
}