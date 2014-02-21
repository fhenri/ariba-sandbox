package sandbox.excel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * Helper method to get some nice layout to an xl file
 *
 * @author fhenri
 */
public class ExcelHelper {

    private HSSFWorkbook workbook;
    
    private HSSFCellStyle hlinkStyle;
    private HSSFCellStyle headerTableStyle;
    private HSSFCellStyle nonVisibleFieldStyle;
    private HSSFCellStyle defaultFieldStyle;
    
    /**
     * Public constructor. Initializes the workbook
     * 
     * @param workbook
     */
    public ExcelHelper (HSSFWorkbook workbook) {
        this.workbook = workbook;
        initializeStyle();
    }
    
    /**
     * Return workbook
     * 
     * @return
     */
    public HSSFWorkbook getWorkbook () {
        return workbook;
    }
    
    /**
     * create a formula which represents an hyperlink to a given sheet with a specified name
     * 
     * @param sheetName
     * @param name name to use in the cell
     * @return
     */
    public String createFormula (String sheetName, String name) {
        return "HYPERLINK(\"#" + sheetName + "!B2\", \"" + name + "\")";
    }
    
    /**
     * Initialize all the styles that we can use in the xl file.
     * Central place to define styles, create style here so it can be used if any place.
     */
    private void initializeStyle () {
        hlinkStyle = workbook.createCellStyle();
        HSSFFont hlinkFont = workbook.createFont();
        hlinkFont.setFontName("Verdana");
        hlinkFont.setUnderline(HSSFFont.U_SINGLE);
        hlinkFont.setColor(HSSFColor.DARK_BLUE.index);
        hlinkStyle.setFont(hlinkFont);
        
        headerTableStyle = workbook.createCellStyle();
        HSSFFont headerTableFont = workbook.createFont();
        headerTableFont.setFontName("Verdana");
        headerTableFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        headerTableFont.setFontHeightInPoints((short)12);
        headerTableStyle.setFont(headerTableFont);

        defaultFieldStyle = workbook.createCellStyle();
        HSSFFont defaultFieldFont = workbook.createFont();
        defaultFieldFont.setFontName("Verdana");
        defaultFieldFont.setFontHeightInPoints((short)10);
        defaultFieldStyle.setFont(defaultFieldFont);

        nonVisibleFieldStyle = workbook.createCellStyle();
        HSSFFont nonVisibleFieldFont = workbook.createFont();
        nonVisibleFieldFont.setFontName("Verdana");
        nonVisibleFieldFont.setStrikeout(true);
        nonVisibleFieldFont.setFontHeightInPoints((short)8);
        nonVisibleFieldStyle.setFont(nonVisibleFieldFont);
    }
    
    // -------------------------------------------------
    // GETTER METHODS WHICH RETURNS THE DIFFERENT STYLES
    // -------------------------------------------------
    public HSSFCellStyle getHyperLinkStyle () {
        return hlinkStyle;
    }
    public HSSFCellStyle getHeaderTableStyle () {
        return headerTableStyle;
    }
    public HSSFCellStyle getNonVisibleFieldStyle () {
        return nonVisibleFieldStyle;
    }
    public HSSFCellStyle getDefaultFieldStyle () {
        return defaultFieldStyle;
    }
}
