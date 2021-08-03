package guru.qa;

import net.lingala.zip4j.ZipFile;
import net.lingala.zip4j.model.FileHeader;
import org.apache.commons.io.IOUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URISyntaxException;
import java.nio.charset.StandardCharsets;
import java.util.Objects;

import static java.util.Objects.requireNonNull;

public final class Utils {
    private Utils() {
    }

    public static InputStream getInputStreamForFileName(String fileName) {
        ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
        return classLoader.getResourceAsStream(fileName);
    }

    public static String readFromDocx(String fileName) throws IOException, InvalidFormatException {
        InputStream stream = getInputStreamForFileName(fileName);
        XWPFDocument document = new XWPFDocument(OPCPackage.open(stream));
        XWPFWordExtractor extractor = new XWPFWordExtractor(document);
        return extractor.getText().trim();
    }

    public static String readFirstCellFromXlsx(String fileName) throws IOException {
        InputStream stream = getInputStreamForFileName(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(stream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        XSSFCell cell = row.getCell(row.getFirstCellNum());
        return cell.getStringCellValue();
    }

    public static String readFirstLineFromPdf(String fileName) throws IOException {
        InputStream stream = getInputStreamForFileName(fileName);
        try (PDDocument document = PDDocument.load(stream)) {

            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            stripper.setSortByPosition(true);

            PDFTextStripper tStripper = new PDFTextStripper();

            String pdfFileInText = tStripper.getText(document);

            String[] lines = pdfFileInText.split("\\r?\\n");
            return lines[0].trim();
        }
    }

    public static String readFromZip(String fileName, String fileNameInZip, String password) throws URISyntaxException, IOException {
        ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
        File file = new File(requireNonNull(classLoader.getResource(fileName), "Не найден файл " + fileName).toURI());
        ZipFile zipFile = new ZipFile(file, password.toCharArray());

        FileHeader fileHeader = zipFile.getFileHeader(fileNameInZip);
        InputStream inputStream = zipFile.getInputStream(fileHeader);
        return IOUtils.toString(inputStream, StandardCharsets.UTF_8);
    }
}