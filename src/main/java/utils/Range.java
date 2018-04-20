package utils;

/**
 * Created by ws5103 on 4/18/18.
 */
public class Range {
	private String sheetTitle;
	private String firstCellPointer;
	private String lastCellPointer;

	public Range(String sheetTitle) {
		this.sheetTitle = sheetTitle;
	}

	public Range(String sheetTitle, String firstCellPointer,
			String lastCellPointer) {
		this.sheetTitle = sheetTitle;
		this.firstCellPointer = firstCellPointer;
		this.lastCellPointer = lastCellPointer;
	}

	public String getSheetTitle() {
		return sheetTitle;
	}

	public void setSheetTitle(String sheetTitle) {
		this.sheetTitle = sheetTitle;
	}

	public String getFirstCellPointer() {
		return firstCellPointer;
	}

	public void setFirstCellPointer(String firstCellPointer) {
		this.firstCellPointer = firstCellPointer;
	}

	public String getLastCellPointer() {
		return lastCellPointer;
	}

	public void setLastCellPointer(String lastCellPointer) {
		this.lastCellPointer = lastCellPointer;
	}

	public String toString(){
		if(getFirstCellPointer().isEmpty() || getLastCellPointer().isEmpty()){
			if(getFirstCellPointer().isEmpty()){
				return sheetTitle + "!" + getLastCellPointer();
			}

			if(getLastCellPointer().isEmpty()){
				return sheetTitle + "!" + getFirstCellPointer();
			}
		}

		return sheetTitle + "!" + getFirstCellPointer() + ":" + getLastCellPointer();

	}
}
