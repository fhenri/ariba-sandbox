/**
 * 
 */
package org.fhsolution.admin.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;


/**
 * 
 * @author fhenri
 *
 */
public interface ExcelExport {

    public void initializeSheet(HSSFWorkbook workbook);
    
    public void processExport();
}
