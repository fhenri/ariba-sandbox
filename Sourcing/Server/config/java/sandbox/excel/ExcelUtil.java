/**
 * 
 */
package org.fhsolution.admin.excel;

import ariba.base.core.Base;
import ariba.base.core.BaseId;
import ariba.base.core.Partition;
import ariba.base.core.aql.AQLOptions;
import ariba.base.core.aql.AQLQuery;
import ariba.base.core.aql.AQLResultCollection;
import ariba.util.core.Fmt;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import java.util.List;
import java.util.Locale;

/**
 * Collection of Util methods to work with excel file and report.
 *
 * @author fhenri
 *
 */
public class ExcelUtil {

    /**
     * Create the header row for the worksheet
     *
     * @param sheet
     * @param header
     */
    public static void initializeHeaderSheet (HSSFSheet sheet, List<ExcelTitleElement> header) {
        HSSFRow title = sheet.createRow(0);
        for (int s=0; s<header.size(); s++) {
            ExcelTitleElement element = (ExcelTitleElement) header.get(s);
            title.createCell(s).setCellValue(new HSSFRichTextString(element.getName()));
            sheet.setColumnWidth(s, element.getColumnSize());
        }
    }

    /**
     * Write information about a subquery results
     * 
     * @param queryText query to run
     * @param baseId criteria for the query - must be a baseId object (not verified)
     * @param sheet sheet to write the result
     * @param cptRow
     * @return the update counter for the row and points to last row from the sheet
     */
    public static int writeSubQueryForGroup (
            String queryText, BaseId baseId, HSSFSheet sheet, int cptRow) {
        
        AQLOptions options = new AQLOptions(Partition.None);
        options.setPartition(Partition.Any);
        options.setUserLocale(Locale.ENGLISH);
        options.setUserPartition(Partition.Any);
        options.setAssertOnMaxRowsReturned(false);
        
        AQLQuery query = AQLQuery.parseQuery(Fmt.S(queryText, baseId.toDBString()));
        AQLResultCollection results = Base.getService().executeQuery(query, options);
        
        while (results.next()) {
            
            HSSFRow row = sheet.createRow(cptRow++);
            row.createCell(0).setCellValue(new HSSFRichTextString(results.getString(0)));
            row.createCell(1).setCellValue(new HSSFRichTextString(results.getString(1)));
            row.createCell(2).setCellValue(new HSSFRichTextString(results.getString(2)));
        }
        
        return cptRow;
    }

}
