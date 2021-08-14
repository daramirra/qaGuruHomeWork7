package guru.qa;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.jupiter.api.Test;
import java.io.IOException;
import java.net.URISyntaxException;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class ReadFilesTest {

    @Test
    public void readDocxFileTest() throws IOException, InvalidFormatException {
        assertEquals("Тестовые данные", Utils.readFromDocx("doc_file.docx"));
    }

    @Test
    public void readXlsxFile() throws IOException {
        assertEquals("Тестовые данные", Utils.readFirstCellFromXlsx("xls_file.xlsx"));
    }

    @Test
    public void readPdfFile() throws IOException {
        assertEquals("A Simple PDF File", Utils.readFirstLineFromPdf("pdf_file.pdf"));
    }

    @Test
    public void readZipFile() throws URISyntaxException, IOException {
        assertEquals("Тестовые данные", Utils.readFromZip("zip_file.zip", "txt_file.txt", "zip123"));
    }

}