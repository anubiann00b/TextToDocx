import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.math.BigInteger;
import javax.xml.bind.JAXBException;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.PPr;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Style;

public class Generator {

    public static void main(String[] args) {
        if (args.length==1 && (args[0].equals("-h") || args[0].equals("--help"))) {
            printHelp();
            System.exit(0);
        } else if (args.length == 6) {
            
        } else if (args.length == 2) {
        } else {
            System.out.println("Invalid arguments.");
            printHelp();
            System.exit(0);
        }
        
        String fileIn = args[0];
        String fileOut = args[1];
        
        Generator g = new Generator(new MLA("Shreyas", "Mr. Becker", "AP World History p. 5", "24 February 2014", "Title"));
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
    MLA mla = null;
    
    Generator() { }
    
    Generator(MLA m) {
        mla = m;
    }
    
    void generate(String in, String out) throws JAXBException, Docx4JException, IOException {
        BufferedReader r = new BufferedReader(new FileReader(in));
        
        wordDoc = WordprocessingMLPackage.createPackage();
        
        MainDocumentPart mdp = wordDoc.getMainDocumentPart();
        
        if (mla != null) {
            mdp.addParagraphOfText(mla.name);
            mdp.addParagraphOfText(mla.teacher);
            mdp.addParagraphOfText(mla.classInfo);
            mdp.addParagraphOfText(mla.date);
            mdp.addParagraphOfText(mla.title);
        }
        
        String line;
        while((line = r.readLine()) != null) {
            mdp.addParagraphOfText(line);
        }
        
        mdp.getStyleDefinitionsPart().getJaxbElement().getStyle().stream().forEach((Style s) -> {
            if (s.getStyleId().equals("Normal"))
                setStyleMLA(s);
        });
        
        wordDoc.save(new File(out));
        System.out.println("Saved " + out);

    }
    
    void setStyleMLA(Style style) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        PPr paragraphProperties = factory.createPPr();
        Jc justification = factory.createJc();
        justification.setVal(JcEnumeration.BOTH);
        paragraphProperties.setJc(justification);
        
        style.setPPr(paragraphProperties);
        
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

class MLA {
    
    String name;
    String teacher;
    String classInfo;
    String date;
    String title;
    
    public MLA(String n, String t, String c, String d, String h) {
        name = n;
        teacher = t;
        classInfo = c;
        date = d;
        title = h;
    }
}