package dhans;


import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;

import com.aspose.words.Document;
import com.aspose.words.FontSettings;
import com.aspose.words.SaveFormat;

public class Mainmethod {

  public static void main(String[] args) throws Exception {

    FontSettings.getDefaultInstance().setFontsFolder("fonts", true);

    // pdfToWordConvertor();
    wordToPdfConvertor();

    System.out.println("created");
  }

  private static void wordToPdfConvertor() throws Exception {
    Set<String> keys = new LinkedHashSet<>();
    keys.add("name");
    keys.add("father_name");
    keys.add("address");
    List<String> values = new LinkedList<>();
    values.add("dhans");
    values.add("srini");
    values.add("namakkal");
    
    Document doc = new Document("inori.docx");
    doc.getMailMerge().execute(keys.toArray(new String[0]), values.toArray());
    doc.save("outtamil.pdf", SaveFormat.PDF);

  }

  private static void pdfToWordConvertor() {
    com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document("intamil.pdf");
    pdfDocument.save("outtamil.doc", com.aspose.pdf.SaveFormat.Doc);
  }

}
