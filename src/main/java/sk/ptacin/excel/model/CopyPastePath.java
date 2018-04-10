package sk.ptacin.excel.model;

/**
 * Created by Michal Ptacin (michal.ptacin@icz.sk) on 2. 6. 2016.
 */
public class CopyPastePath {

    private String sourceSheet;
    private String sourceId;
    private String sourceCell;
    private String targetSheet;
    private String tagetCell;

    public String getSourceSheet() {
        return sourceSheet;
    }

    public void setSourceSheet(String sourceSheet) {
        this.sourceSheet = sourceSheet;
    }

    public String getSourceId() {
        return sourceId;
    }

    public void setSourceId(String sourceId) {
        this.sourceId = sourceId;
    }

    public String getSourceCell() {
        return sourceCell;
    }

    public void setSourceCell(String sourceCell) {
        this.sourceCell = sourceCell;
    }

    public String getTargetSheet() {
        return targetSheet;
    }

    public void setTargetSheet(String targetSheet) {
        this.targetSheet = targetSheet;
    }

    public String getTagetCell() {
        return tagetCell;
    }

    public void setTagetCell(String tagetCell) {
        this.tagetCell = tagetCell;
    }

    @Override
    public String toString() {
        return "CopyPastePath{" +
                "sourceSheet='" + sourceSheet + '\'' +
                ", sourceId='" + sourceId + '\'' +
                ", sourceCell='" + sourceCell + '\'' +
                ", targetSheet='" + targetSheet + '\'' +
                ", tagetCell='" + tagetCell + '\'' +
                '}';
    }
}
