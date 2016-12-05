package main.java;

/**
 * Created by Administrator on 2016/12/2.
 */
public class ClassItem {
    private String className;
    private String unitName;
    private String catalogName;
    private String fileName;
    private String filePath;
    private String status;

    public String getClassName() {
        return className;
    }

    public void setClassName(String className) {
        this.className = className;
    }

    public String getUnitName() {
        return unitName;
    }

    public void setUnitName(String unitName) {
        this.unitName = unitName;
    }

    public String getCatalogName() {
        return catalogName;
    }

    public void setCatalogName(String catalogName) {
        this.catalogName = catalogName;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }

    @Override
    public String toString() {
        return "ClassItem{" +
                "className='" + className + '\'' +
                ", unitName='" + unitName + '\'' +
                ", catalogName='" + catalogName + '\'' +
                ", fileName='" + fileName + '\'' +
                ", filePath='" + filePath + '\'' +
                ", status='" + status + '\'' +
                '}';
    }
}
