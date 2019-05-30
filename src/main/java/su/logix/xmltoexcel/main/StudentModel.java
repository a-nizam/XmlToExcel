package su.logix.xmltoexcel.main;

import javafx.beans.property.SimpleFloatProperty;
import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleStringProperty;
import org.jetbrains.annotations.Contract;
import org.jetbrains.annotations.NotNull;

public class StudentModel {
    private final SimpleStringProperty name;
    private final SimpleFloatProperty mark;
    private final SimpleIntegerProperty practiceMissed;
    private final SimpleIntegerProperty practiceCorrected;
    private final SimpleIntegerProperty lecturesMissed;
    private final SimpleIntegerProperty lecturesCorrected;
    private final SimpleStringProperty faculty;
    private final SimpleIntegerProperty semester;
    private final SimpleIntegerProperty certNum;
    private final SimpleStringProperty department;
    private final SimpleStringProperty group;
    private final SimpleStringProperty subject;

    public StudentModel() {
        this.name = new SimpleStringProperty();
        this.mark = new SimpleFloatProperty();
        this.practiceMissed = new SimpleIntegerProperty();
        this.practiceCorrected = new SimpleIntegerProperty();
        this.lecturesMissed = new SimpleIntegerProperty();
        this.lecturesCorrected = new SimpleIntegerProperty();
        this.faculty = new SimpleStringProperty();
        this.semester = new SimpleIntegerProperty();
        this.certNum = new SimpleIntegerProperty();
        this.department = new SimpleStringProperty();
        this.group = new SimpleStringProperty();
        this.subject = new SimpleStringProperty();
    }

    public StudentModel(String name, float mark, int practiceMissed, int practiceCorrected, int lecturesMissed, int lecturesCorrected,
                        String faculty, int semester, int certNum, String department, String group, String subject) {
        this.name = new SimpleStringProperty(name);
        this.mark = new SimpleFloatProperty(mark);
        this.practiceMissed = new SimpleIntegerProperty(practiceMissed);
        this.practiceCorrected = new SimpleIntegerProperty(practiceCorrected);
        this.lecturesMissed = new SimpleIntegerProperty(lecturesMissed);
        this.lecturesCorrected = new SimpleIntegerProperty(lecturesCorrected);
        this.faculty = new SimpleStringProperty(faculty);
        this.semester = new SimpleIntegerProperty(semester);
        this.certNum = new SimpleIntegerProperty(certNum);
        this.department = new SimpleStringProperty(department);
        this.group = new SimpleStringProperty(group);
        this.subject = new SimpleStringProperty(subject);
    }

    @NotNull
    @Contract(pure = true)
    public static String getFacultyByCode(int code) {
        switch (code) {
            case 3:
                return "Лечебный";
            case 4:
                return "Педиатрический";
            case 5:
                return "Стоматологический";
            case 6:
                return "Фармацевтический";
            case 7:
                return "Медико-профилактический";
        }
        return "";
    }

    public String getName() {
        return name.get();
    }

    public void setName(String name) {
        this.name.set(name);
    }

    public float getMark() {
        return mark.get();
    }

    public void setMark(float mark) {
        this.mark.set(mark);
    }

    public int getPracticeMissed() {
        return practiceMissed.get();
    }

    public void setPracticeMissed(int practiceMissed) {
        this.practiceMissed.set(practiceMissed);
    }

    public int getPracticeCorrected() {
        return practiceCorrected.get();
    }

    public void setPracticeCorrected(int practiceCorrected) {
        this.practiceCorrected.set(practiceCorrected);
    }

    public int getLecturesMissed() {
        return lecturesMissed.get();
    }

    public void setLecturesMissed(int lecturesMissed) {
        this.lecturesMissed.set(lecturesMissed);
    }

    public int getLecturesCorrected() {
        return lecturesCorrected.get();
    }

    public void setLecturesCorrected(int lecturesCorrected) {
        this.lecturesCorrected.set(lecturesCorrected);
    }

    public String getFaculty() {
        return faculty.get();
    }

    public void setFaculty(String faculty) {
        this.faculty.set(faculty);
    }

    public int getSemester() {
        return semester.get();
    }

    public void setSemester(int semester) {
        this.semester.set(semester);
    }

    public int getCertNum() {
        return certNum.get();
    }

    public void setCertNum(int certNum) {
        this.certNum.set(certNum);
    }

    public String getDepartment() {
        return department.get();
    }

    public void setDepartment(String department) {
        this.department.set(department);
    }

    public String getGroup() {
        return group.get();
    }

    public void setGroup(String group) {
        this.group.set(group);
    }

    public String getSubject() {
        return subject.get();
    }

    public void setSubject(String subject) {
        this.subject.set(subject);
    }
}
