package su.logix.xmltoexcel.main;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ListView;
import javafx.stage.FileChooser;
import org.apache.log4j.Logger;
import org.docx4j.docProps.core.dc.elements.SimpleLiteral;
import org.docx4j.docProps.extended.Properties;
import org.docx4j.docProps.variantTypes.Vector;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.DocPropsCorePart;
import org.docx4j.openpackaging.parts.DocPropsExtendedPart;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.jetbrains.annotations.NotNull;
import org.xlsx4j.jaxb.Context;
import org.xlsx4j.sml.*;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class MainController {

    private static final Logger log = Logger.getLogger(MainController.class);

    @FXML
    public ListView<String> lvFiles;

    private ObservableList<String> filePathList = FXCollections.observableArrayList();
    private List<StudentModel> studentList = new ArrayList<>();

    private void addCoreProps(@NotNull SpreadsheetMLPackage pkg) {
        pkg.addDocPropsCorePart();
        DocPropsCorePart corePart = pkg.getDocPropsCorePart();
        org.docx4j.docProps.core.CoreProperties coreProps = corePart.getJaxbElement();
        org.docx4j.docProps.core.dc.elements.ObjectFactory dcElFactory = new org.docx4j.docProps.core.dc.elements.ObjectFactory();
        SimpleLiteral creator = dcElFactory.createSimpleLiteral();
        coreProps.setCreator(creator);
        creator.getContent().add("User");
    }

    private void addExtendedProps(@NotNull SpreadsheetMLPackage pkg) {
        pkg.addDocPropsExtendedPart();
        DocPropsExtendedPart extendedPart = pkg.getDocPropsExtendedPart();
        org.docx4j.docProps.extended.Properties extendedProps = extendedPart.getJaxbElement();
        extendedProps.setApplication("xlsx4j");

        org.docx4j.docProps.extended.ObjectFactory factoryExtended = new org.docx4j.docProps.extended.ObjectFactory();
        org.docx4j.docProps.variantTypes.ObjectFactory factoryVariantTypes = new org.docx4j.docProps.variantTypes.ObjectFactory();

        Properties.TitlesOfParts titleOfParts = new Properties.TitlesOfParts();
        extendedProps.setTitlesOfParts(titleOfParts);
        Vector vector = factoryVariantTypes.createVector();
        titleOfParts.setVector(vector);
        vector.setSize(1);
        vector.setBaseType("lpstr");
        JAXBElement<String> lpstr = factoryVariantTypes.createLpstr("Sheet1");
        vector.getVariantOrI1OrI2().add(lpstr);

    }

    @FXML
    public void initialize() {
        lvFiles.setItems(filePathList);
    }

    @FXML
    public void selectFiles() {
        List<File> fileList = showSelectFilesDialog();
        if (fileList != null) {
            fileList.forEach(file -> {
                if (!filePathList.contains(file.getAbsolutePath())) {
                    filePathList.add(file.getAbsolutePath());
                } else {
                    showMessage(Alert.AlertType.WARNING, "Предупреждение", "Файл " + file.getAbsolutePath() + " был добавлен ранее");
                }
            });
        }
    }

    private void showMessage(Alert.AlertType alertType, String title, String message) {
        Alert alert = new Alert(alertType);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.getButtonTypes().setAll(ButtonType.OK);
        alert.showAndWait();
    }

    @FXML
    public void export() {
        if (filePathList != null && !filePathList.isEmpty()) {
            // creating students list
            for (String file : filePathList) {
                addFileToList(file);
            }
            File outputFile = showSaveFileDialog();
            if (outputFile != null) {
                createXlsxFile(outputFile);
                showMessage(Alert.AlertType.INFORMATION, "Успешно", "Таблица сохранена: " + outputFile.getAbsolutePath());
            }
        } else {
            showMessage(Alert.AlertType.WARNING, "Предупреждение", "Не выбрано ни одного XML файла");
        }
    }

    private void createXlsxFile(File outputFile) {
        try {
            SpreadsheetMLPackage pkg = SpreadsheetMLPackage.createPackage();
            addCoreProps(pkg);
            addExtendedProps(pkg);
            WorksheetPart sheet = pkg.createWorksheetPart(new PartName("/xl/worksheets/sheet1.xml"), "Sheet1", 1);
            addContent(sheet);
            pkg.save(outputFile);
        } catch (InvalidFormatException e) {
            showMessage(Alert.AlertType.ERROR, "Ошибка", e.getMessage());
            log.error("create package error", e);
            e.printStackTrace();
        } catch (JAXBException e) {
            showMessage(Alert.AlertType.ERROR, "Ошибка", e.getMessage());
            log.error("create worksheet part error", e);
            e.printStackTrace();
        } catch (Docx4JException e) {
            showMessage(Alert.AlertType.ERROR, "Ошибка", e.getMessage());
            log.error("package save error", e);
            e.printStackTrace();
        }
    }

    private void createHeadLine(@NotNull SheetData sheetData) {
        Row row = Context.getsmlObjectFactory().createRow();

        row.getC().add(createStringCell("ФИО", "A1"));
        row.getC().add(createStringCell("Факультет", "B1"));
        row.getC().add(createStringCell("Семестр", "C1"));
        row.getC().add(createStringCell("Номер аттестации", "D1"));
        row.getC().add(createStringCell("Кафедра", "E1"));
        row.getC().add(createStringCell("Группа", "F1"));
        row.getC().add(createStringCell("Предмет", "G1"));
        row.getC().add(createStringCell("Средний балл", "H1"));
        row.getC().add(createStringCell("Лекций пропущено", "I1"));
        row.getC().add(createStringCell("Лекций отработано", "J1"));
        row.getC().add(createStringCell("Практики пропущено", "K1"));
        row.getC().add(createStringCell("Практики отработано", "L1"));

        sheetData.getRow().add(row);
    }

    private void addContent(@NotNull WorksheetPart sheet) {
        SheetData sheetData = sheet.getJaxbElement().getSheetData();
        createHeadLine(sheetData);
        int counter = 2;
        for (StudentModel studentModel : studentList) {
            Row row = Context.getsmlObjectFactory().createRow();
            row.getC().add(createStringCell(studentModel.getName(), "A" + counter));
            row.getC().add(createStringCell(studentModel.getFaculty(), "B" + counter));
            row.getC().add(createIntCell(studentModel.getSemester(), "C" + counter));
            row.getC().add(createIntCell(studentModel.getCertNum(), "D" + counter));
            row.getC().add(createStringCell(studentModel.getDepartment(), "E" + counter));
            row.getC().add(createStringCell(studentModel.getGroup(), "F" + counter));
            row.getC().add(createStringCell(studentModel.getSubject(), "G" + counter));
            row.getC().add(createFloatCell(studentModel.getMark(), "H" + counter));
            row.getC().add(createIntCell(studentModel.getLecturesMissed(), "I" + counter));
            row.getC().add(createIntCell(studentModel.getLecturesCorrected(), "J" + counter));
            row.getC().add(createIntCell(studentModel.getPracticeMissed(), "K" + counter));
            row.getC().add(createIntCell(studentModel.getPracticeCorrected(), "L" + counter));
            sheetData.getRow().add(row);
            counter++;
        }
    }

    private Cell createIntCell(int content, String cellName) {
        Cell cell = Context.getsmlObjectFactory().createCell();
        cell.setV(Integer.toString(content));
        cell.setR(cellName);
        return cell;
    }

    private Cell createFloatCell(float content, String cellName) {
        Cell cell = Context.getsmlObjectFactory().createCell();
        cell.setV(Float.toString(content));
        cell.setR(cellName);
        return cell;
    }

    private Cell createStringCell(String content, String cellName) {
        Cell cell = Context.getsmlObjectFactory().createCell();
        CTXstringWhitespace ctx = Context.getsmlObjectFactory().createCTXstringWhitespace();
        ctx.setValue(content);
        CTRst ctrst = new CTRst();
        ctrst.setT(ctx);
        cell.setT(STCellType.INLINE_STR);
        cell.setIs(ctrst);
        cell.setR(cellName);
        return cell;
    }

    private File showSaveFileDialog() {
        final FileChooser fileChooser = new FileChooser();
        FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Файлы таблиц (*.xlsx)", "*.xlsx");
        fileChooser.setTitle("Сохранение файла");
        fileChooser.getExtensionFilters().add(extFilter);
        return fileChooser.showSaveDialog(lvFiles.getScene().getWindow());
    }

    private List<File> showSelectFilesDialog() {
        final FileChooser fileChooser = new FileChooser();
        FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("XML files (*.xml)", "*.xml");
        fileChooser.setTitle("Выберите XML файлы");
        fileChooser.getExtensionFilters().add(extFilter);
        return fileChooser.showOpenMultipleDialog(lvFiles.getScene().getWindow());
    }

    private void addFileToList(String file) {
        try {
            InputStream fileInputStream = new BufferedInputStream(new FileInputStream(file));
            XMLInputFactory xmlInputFactory = XMLInputFactory.newInstance();
            XMLStreamReader xmlStreamReader = xmlInputFactory.createXMLStreamReader(fileInputStream, "UTF-8");
            int event;
            ActionType action = ActionType.SKIP;
            String localName;
            int faculty = 0;
            int semester = 0;
            int certNum = 0;
            String department = "";
            String group = "";
            String subject = "";
            StudentModel studentModel = new StudentModel();
            while (xmlStreamReader.hasNext()) {
                event = xmlStreamReader.next();
                switch (event) {
                    case XMLStreamConstants.START_ELEMENT:
                        localName = xmlStreamReader.getLocalName();
                        switch (localName) {
                            case "ДанныеСтудента":
                                studentModel = new StudentModel();
                                action = ActionType.SKIP;
                                break;
                            case "ФИО":
                                action = ActionType.READ_NAME;
                                break;
                            case "СреднийБалл":
                                action = ActionType.READ_MARK;
                                break;
                            case "ПропускиЛекций":
                                action = ActionType.READ_LECTURES_MISSED;
                                break;
                            case "ПропускиПрактических":
                                action = ActionType.READ_PRACTICE_MISSED;
                                break;
                            case "ОтработкиЛекций":
                                action = ActionType.READ_LECTURES_CORRECTED;
                                break;
                            case "ОтработкиПрактических":
                                action = ActionType.READ_PRACTICE_CORRECTED;
                                break;
                            case "Факультет":
                                action = ActionType.READ_FACULTY;
                                break;
                            case "ПериодКонтроля":
                                action = ActionType.READ_SEMESTER;
                                break;
                            case "АттестацияНомер":
                                action = ActionType.READ_CERT_NUM;
                                break;
                            case "Кафедра":
                                action = ActionType.READ_DEPARTMENT;
                                break;
                            case "Группа":
                                action = ActionType.READ_GROUP;
                                break;
                            case "Дисциплина":
                                action = ActionType.READ_SUBJECT;
                                break;
                            default:
                                action = ActionType.SKIP;
                                break;
                        }
                        break;

                    case XMLStreamConstants.CHARACTERS:
                        switch (action) {
                            case READ_NAME:
                                studentModel.setName(xmlStreamReader.getText());
                                break;
                            case READ_MARK:
                                studentModel.setMark(Float.parseFloat(xmlStreamReader.getText()));
                                break;
                            case READ_LECTURES_MISSED:
                                studentModel.setLecturesMissed(Integer.parseInt(xmlStreamReader.getText()));
                                break;
                            case READ_PRACTICE_MISSED:
                                studentModel.setPracticeMissed(Integer.parseInt(xmlStreamReader.getText()));
                                break;
                            case READ_LECTURES_CORRECTED:
                                studentModel.setLecturesCorrected(Integer.parseInt(xmlStreamReader.getText()));
                                break;
                            case READ_PRACTICE_CORRECTED:
                                studentModel.setPracticeCorrected(Integer.parseInt(xmlStreamReader.getText()));
                                break;
                            case READ_FACULTY:
                                faculty = Integer.parseInt(xmlStreamReader.getText());
                                break;
                            case READ_SEMESTER:
                                semester = Integer.parseInt(xmlStreamReader.getText());
                                break;
                            case READ_CERT_NUM:
                                certNum = Integer.parseInt(xmlStreamReader.getText());
                                break;
                            case READ_DEPARTMENT:
                                department = xmlStreamReader.getText();
                                break;
                            case READ_GROUP:
                                group = xmlStreamReader.getText();
                                break;
                            case READ_SUBJECT:
                                subject = xmlStreamReader.getText();
                                break;
                        }
                        break;
                    case XMLStreamConstants.END_ELEMENT:
                        // save model to list
                        if (xmlStreamReader.getLocalName().equals("ДанныеСтудента")) {
                            studentModel.setFaculty(StudentModel.getFacultyByCode(faculty));
                            studentModel.setSemester(semester);
                            studentModel.setCertNum(certNum);
                            studentModel.setDepartment(department);
                            studentModel.setGroup(group);
                            studentModel.setSubject(subject);
                            studentList.add(studentModel);
                        }
                        action = ActionType.SKIP;
                        break;
                }
            }
            fileInputStream.close();
        } catch (FileNotFoundException e) {
            showMessage(Alert.AlertType.ERROR, "Ошибка", "Не удалось открыть файл: " + e.getMessage());
            log.error("open file error", e);
            e.printStackTrace();
        } catch (XMLStreamException e) {
            showMessage(Alert.AlertType.ERROR, "Ошибка", e.getMessage());
            log.error("xml stream exception", e);
            e.printStackTrace();
        } catch (IOException e) {
            showMessage(Alert.AlertType.ERROR, "Ошибка", "Не удалось закрыть поток чтения xml: " + e.getMessage());
            log.error("input stream close error", e);
            e.printStackTrace();
        }
    }

    @FXML
    public void clearFileList() {
        filePathList.clear();
        studentList.clear();
    }

    private enum ActionType {
        SKIP, READ_NAME, READ_MARK, READ_PRACTICE_MISSED, READ_PRACTICE_CORRECTED, READ_LECTURES_MISSED,
        READ_LECTURES_CORRECTED, READ_FACULTY, READ_SEMESTER, READ_CERT_NUM, READ_DEPARTMENT, READ_GROUP, READ_SUBJECT
    }
}
