package com.isuru.docxpoi.utils;

import com.isuru.docxpoi.dto.EmployeeDetails;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import lombok.SneakyThrows;
import org.apache.poi.common.usermodel.PictureType;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.springframework.stereotype.Component;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

@Component
public class DocumentHelper {

    public ByteArrayOutputStream createDocument(Integer id) throws Exception {

        //get employee by id from database. this is dummy data.
        EmployeeDetails employeeDetails = EmployeeDetails.builder()
                .firstName("Ranil sdasdasd  asdasdasdasda sdasdasdasdasdasdasdasdasdasdas dasdasd asdasdasd asdasdasdas dasdasdasdasd asdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasdasda")
                .lastName("Perera")
                .address("100,Temple Street,Colombo")
                .dob(LocalDate.now().minusYears(26))
                .employeeId(id)
                .gender("Male")
                .position("Software Engineer")
                .build();

        String resourcePath = "template.docx";
        Path templatePath = Paths.get(DocumentHelper.class.getClassLoader().getResource(resourcePath).toURI());
        XWPFDocument doc = new XWPFDocument(Files.newInputStream(templatePath));

        Map<String, String> map = new HashMap<>();
        map.put(VariableTypes.FIRST_NAME.getName(), employeeDetails.getFirstName());
        map.put(VariableTypes.LAST_NAME.getName(), employeeDetails.getLastName());
        map.put(VariableTypes.POSITION.getName(), employeeDetails.getPosition());
        map.put(VariableTypes.GENDER.getName(), employeeDetails.getGender());
        map.put(VariableTypes.DATE_OF_BIRTH.getName(), employeeDetails.getDob().toString());
        map.put(VariableTypes.ADDRESS.getName(), employeeDetails.getAddress());
        map.put(VariableTypes.EMPLOYEE_ID.getName(), employeeDetails.getEmployeeId().toString());

        replaceTextFor(doc, map);
        fillPictureByAltText(doc);
        //get data from database. this is dummy data.
        List<SalaryRecord> salaryRecordList = Arrays.asList(
                SalaryRecord.builder().month("Jan 2020").amount(String.valueOf(1200.30)).build(),
                SalaryRecord.builder().month("Feb 2020").amount(String.valueOf(1200.30)).build(),
                SalaryRecord.builder().month("Mar 2020").amount(String.valueOf(1200.30)).build(),
                SalaryRecord.builder().month("Apr 2020").amount(String.valueOf(1200.30)).build(),
                SalaryRecord.builder().month("May 2020").amount(String.valueOf(1500.70)).build(),
                SalaryRecord.builder().month("Jun 2020").amount(String.valueOf(1500.70)).build()
        );

        replaceSalaryTable(doc, salaryRecordList);

//        savePdf("src/main/resources/employee-report.pdf", doc);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        PdfOptions options = PdfOptions.create();
        PdfConverter.getInstance().convert(doc, out, options);
        out.close();
        return out;
    }

    private void replaceTextFor(XWPFDocument doc, Map<String, String> map) {

        doc.getParagraphs().forEach(p -> {
            String paragraphText = p.getRuns().stream().map(XWPFRun::text).collect(Collectors.joining());
            System.out.println(doc.getBodyElements().indexOf(p) + "  / " + paragraphText);
        });

        doc.getParagraphs().forEach(p -> p.getRuns().forEach(run -> {
            String text = run.text();
            map.forEach((findText, replaceText) -> {
                if (text.contains(findText)) {
                    run.setText(text.replace(findText, replaceText), 0);
                }
            });
        }));
    }

    private void removeGroupIfNoData(XWPFDocument doc) {
        AtomicInteger startIndex = new AtomicInteger(-1);
        AtomicInteger endIndex = new AtomicInteger(-1);

        doc.getParagraphs().forEach(p -> {
            String paragraphText = p.getRuns().stream().map(i -> i.text()).collect(Collectors.joining());
            System.out.println(doc.getBodyElements().indexOf(p) + "  / " + paragraphText);
        });

        doc.getParagraphs().forEach(p -> p.getRuns().forEach(run -> {
            String text = run.text();
//            System.out.println(doc.getBodyElements().indexOf(p) + "  / " + text);
            if ("#111".equals(text)) {
                if (startIndex.get() > -1)
                    endIndex.set(doc.getBodyElements().indexOf(p));
                else
                    startIndex.set(doc.getBodyElements().indexOf(p));
            }
        }));

        for (int i = endIndex.get(); i >= startIndex.get(); i--) {
//            doc.getParagraphs().remove(i);
            doc.removeBodyElement(i);
        }
    }

    private void fillPictureByAltText(XWPFDocument doc) throws IOException {
        doc.getParagraphs().forEach(p -> p.getRuns().forEach(run -> {
            if (run.text().equals("##imageGoesHere##")) {
//                ParagraphReplacer r = new ParagraphReplacer("imageGoesHere", "");
//                r.replace(p);
                p.setAlignment(ParagraphAlignment.CENTER);
                run.setText("", 0);
                try (InputStream in = DocumentHelper.class.getResourceAsStream("/logo-search-grid-1x.png")) {
                    XWPFPicture img = run.addPicture(in, PictureType.JPEG, "img", (int) (Units.EMU_PER_CENTIMETER * 8), (int) (Units.EMU_PER_CENTIMETER * 8));
                } catch (IOException e) {
                    throw new RuntimeException(e);
                } catch (InvalidFormatException e) {
                    throw new RuntimeException(e);
                }
            }
        }));

        doc.getParagraphs().forEach(p -> p.getRuns().forEach(run -> {
            run.getEmbeddedPictures().forEach(ep -> {
                if (ep.getDescription().equals("altText")) {

                }
            });
            if (run.text().equals("##imageGoesHere##")) {
//                ParagraphReplacer r = new ParagraphReplacer("imageGoesHere", "");
//                r.replace(p);
                p.setAlignment(ParagraphAlignment.CENTER);
                run.setText("", 0);
                try (InputStream in = DocumentHelper.class.getResourceAsStream("/logo-search-grid-1x.png")) {
                    XWPFPicture img = run.addPicture(in, PictureType.JPEG, "img", (int) (Units.EMU_PER_CENTIMETER * 10), (int) (Units.EMU_PER_CENTIMETER * 10));
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            }
        }));
    }

    public static void deleteParagraph(XWPFParagraph p) {
        XWPFDocument doc = p.getDocument();
        int pPos = doc.getPosOfParagraph(p);
        //doc.getDocument().getBody().removeP(pPos);
        doc.removeBodyElement(pPos);
    }

    static void replacePictureData(XWPFPictureData source, byte[] data) {
        try (ByteArrayInputStream in = new ByteArrayInputStream(data);
             OutputStream out = source.getPackagePart().getOutputStream();
        ) {
            byte[] buffer = new byte[2048];
            int length;
            while ((length = in.read(buffer)) > 0) {
                out.write(buffer, 0, length);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private void savePdf(String filePath, XWPFDocument doc) throws IOException {
        PdfOptions options = PdfOptions.create();
        OutputStream out = new FileOutputStream(filePath);
        PdfConverter.getInstance().convert(doc, out, options);
    }

    public void replaceSalaryTable(XWPFDocument doc, List<SalaryRecord> salaryRecordList) {
        XWPFTable table = doc.getTableArray(0);
//        int posOfTable = doc.getPosOfTable(table);
//        doc.removeBodyElement(posOfTable);
        int templateRowId = 1;
        XWPFTableRow rowTemplate = table.getRow(templateRowId);

        salaryRecordList.forEach(salaryRecord -> {

            CTRow ctrow = getCtRow(rowTemplate);

            XWPFTableRow newRow = new XWPFTableRow(ctrow, table);

            newRow.getCell(0).getParagraphArray(0).getRuns().get(0).setText(salaryRecord.getMonth(), 0);
            newRow.getCell(1).getParagraphArray(0).getRuns().get(0).setText(salaryRecord.getAmount(), 0);

            table.addRow(newRow);
        });

        table.removeRow(templateRowId);
    }

    @SneakyThrows
    private CTRow getCtRow(XWPFTableRow rowTemplate) {
        return CTRow.Factory.parse(rowTemplate.getCtRow().newInputStream());
    }
}