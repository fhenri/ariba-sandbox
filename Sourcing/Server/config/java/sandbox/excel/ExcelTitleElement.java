/**
 * 
 */
package org.fhsolution.admin.excel;

/**
 *
 * @author fhenri
 *
 */
public class ExcelTitleElement {

    public static int DEFAULT_COL_SIZE = 30*256;
    
    private String name;
    private int columnSize;
    
    public ExcelTitleElement (String name) {
        this(name, DEFAULT_COL_SIZE);
    }
    
    public ExcelTitleElement (String name, int colSize) {
        this.name = name;
        this.columnSize = colSize;
    }

    /**
     * @return the name
     */
    public String getName() {
        return name;
    }

    /**
     * @param name the name to set
     */
    public void setName(String name) {
        this.name = name;
    }

    /**
     * @return the columnSize
     */
    public int getColumnSize() {
        return columnSize;
    }

    /**
     * @param columnSize the columnSize to set
     */
    public void setColumnSize(int columnSize) {
        this.columnSize = columnSize;
    }
    
    
}
