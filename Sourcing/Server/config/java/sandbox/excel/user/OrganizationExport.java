/**
 * 
 */
package org.fhsolution.admin.excel.user;

import ariba.base.core.Base;
import ariba.base.core.Partition;
import ariba.base.core.aql.AQLOptions;
import ariba.base.core.aql.AQLQuery;
import ariba.base.core.aql.AQLResultCollection;
import ariba.util.core.Fmt;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.fhsolution.admin.excel.ExcelExport;
import org.fhsolution.admin.excel.ExcelTitleElement;
import org.fhsolution.admin.excel.ExcelUtil;

import java.util.LinkedList;
import java.util.List;
import java.util.Locale;

/**
 *
 * @author fhenri
 *
 */
public class OrganizationExport implements ExcelExport {

    protected HSSFSheet organizationSheet;
    
    protected List<ExcelTitleElement> organizationSheetHeader;
    
    protected String organizationSheetName;
   
    /**
     * Public constructor
     */
    public OrganizationExport (String sheetName) {
        this.organizationSheetName = sheetName;
        
        organizationSheetHeader = new LinkedList<ExcelTitleElement>();
        organizationSheetHeader.add(new ExcelTitleElement("SystemID"));
        organizationSheetHeader.add(new ExcelTitleElement("Full Name"));
        organizationSheetHeader.add(new ExcelTitleElement("Email Address"));
        organizationSheetHeader.add(new ExcelTitleElement("Corporate URL"));
        organizationSheetHeader.add(new ExcelTitleElement("Address"));
        organizationSheetHeader.add(new ExcelTitleElement("StreetAddress"));
        organizationSheetHeader.add(new ExcelTitleElement("PostalCode"));
        organizationSheetHeader.add(new ExcelTitleElement("City"));
        organizationSheetHeader.add(new ExcelTitleElement("Country"));
        organizationSheetHeader.add(new ExcelTitleElement("Phone"));
        
    }
    
    public void initializeSheet (HSSFWorkbook workbook) {
        organizationSheet = workbook.createSheet(organizationSheetName);
        ExcelUtil.initializeHeaderSheet(organizationSheet, organizationSheetHeader);
    }
    
    protected String additionalFields = "";
    protected String whereCriteria    = "";

    // this is the main query to search on organization including main components
    // there's 2 placeholders to add additional field + joins and/or where criteria
    protected static String BASE_QUERY =
            "select o.SystemID, o.Name, o.CorporateEmailAddress, o.CorporateURL, " +
            "a.Name, pa.Lines, pa.PostalCode, pa.City, co.Name, o.CorporatePhone %s"  +
            " from ariba.user.core.Organization o " +
            "left outer join ariba.basic.core.Address a using o.CorporateAddress " +
            "left outer join ariba.basic.core.PostalAddress pa using a.PostalAddress " +
            "left outer join ariba.basic.core.Country co using pa.Country %s";

    protected String getQueryText () {
        return Fmt.S(BASE_QUERY, additionalFields, whereCriteria);
    }

    public void processExport () {
        processQuery(getQueryText());
    }
    
    protected void processQuery (String queryText) {
        
        AQLQuery query = AQLQuery.parseQuery(queryText);
        AQLOptions options = new AQLOptions(Partition.None);
        options.setPartition(Partition.Any);
        options.setUserLocale(Locale.ENGLISH);
        options.setUserPartition(Partition.Any);
        AQLResultCollection results = Base.getService().executeQuery(query, options);
        
        // loop on results
        int cptRow = 1;
        while (results.next()) {
            HSSFRow row = organizationSheet.createRow(cptRow++);
            writeResultInRow(row, results);
        }
    }

    protected void writeResultInRow (HSSFRow row, AQLResultCollection result) {
        row.createCell(0).setCellValue(new HSSFRichTextString(result.getString(0)));
        row.createCell(1).setCellValue(new HSSFRichTextString(result.getString(1)));
        row.createCell(2).setCellValue(new HSSFRichTextString(result.getString(2)));
        row.createCell(3).setCellValue(new HSSFRichTextString(result.getString(3)));
        row.createCell(4).setCellValue(new HSSFRichTextString(result.getString(4)));
        row.createCell(5).setCellValue(new HSSFRichTextString(result.getString(5)));
        row.createCell(6).setCellValue(new HSSFRichTextString(result.getString(6)));
        row.createCell(7).setCellValue(new HSSFRichTextString(result.getString(7)));
        row.createCell(8).setCellValue(new HSSFRichTextString(result.getString(8)));
        row.createCell(9).setCellValue(new HSSFRichTextString(result.getString(9)));
    }
    
}
