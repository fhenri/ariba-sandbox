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
import ariba.util.core.StringUtil;
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
 * Export User informatin in xl file
 * 
 * @author fhenri
 *
 */
public class UserExport implements ExcelExport {

    protected HSSFSheet userSheet;
    private String sheetName;
   
    //List<ExcelTitleElement> userSheetHeader;
    protected List<ExcelTitleElement> userSheetHeader;
    
    /**
     * Public constructor
     */
    public UserExport (String sheetName) {
        this.sheetName = sheetName;
        
        userSheetHeader = new LinkedList<ExcelTitleElement>();
        userSheetHeader.add(new ExcelTitleElement("UniqueName"));
        userSheetHeader.add(new ExcelTitleElement("PasswordAdapter"));
        userSheetHeader.add(new ExcelTitleElement("Full Name"));
        userSheetHeader.add(new ExcelTitleElement("Email Address"));
        userSheetHeader.add(new ExcelTitleElement("Supervisor"));
        userSheetHeader.add(new ExcelTitleElement("Currency"));
        userSheetHeader.add(new ExcelTitleElement("Locale"));
        userSheetHeader.add(new ExcelTitleElement("Organization.SystemID"));
        userSheetHeader.add(new ExcelTitleElement("Organization.Name"));
        userSheetHeader.add(new ExcelTitleElement("Department"));
        userSheetHeader.add(new ExcelTitleElement("Address"));
        userSheetHeader.add(new ExcelTitleElement("Phone"));
        
    }
    
    /**
     * Initialize xl sheet
     * 
     */
    public void initializeSheet (HSSFWorkbook workbook) {
        userSheet = workbook.createSheet(sheetName);
        ExcelUtil.initializeHeaderSheet(userSheet, userSheetHeader);
    }

    protected String additionalFields = "";
    protected String whereCriteria    = "";

    protected static String BASE_QUERY = 
        "Select u.UniqueName, u.PasswordAdapter, u.Name, u.EmailAddress, " +
        "s.UniqueName, c.UniqueName, l.UniqueName, " +
        "o.SystemID, o.Name, d.DepartmentID, d.Description, " +
        "pa.City, co.Name, u.Phone %s" +
        "from ariba.user.core.User u " +
        "left outer join ariba.user.core.User s using u.Supervisor " +
        "left outer join ariba.basic.core.Currency c using u.DefaultCurrency " +
        "left outer join ariba.basic.core.LocaleID l using u.LocaleID " +
        "left outer join ariba.user.core.Organization o using u.Organization " +
        "left outer join ariba.collaborate.basic.Department d using u.Department " +
        "left outer join ariba.basic.core.Address a using u.ShipTos " +
        "left outer join ariba.basic.core.PostalAddress pa using a.PostalAddress " +
        "left outer join ariba.basic.core.Country co using pa.Country %s";

    protected String getQueryText () {
        return Fmt.S(BASE_QUERY, additionalFields, whereCriteria);
    }

    public void processExport () {
        processQuery(getQueryText());
    }

    /**
     * Process the User object
     * 
     */
    protected void processQuery (String queryText) {
    
        // query user object
        //AQLQuery query = new AQLQuery(User.ClassName);
        AQLQuery query = AQLQuery.parseQuery(queryText);
        AQLOptions options = new AQLOptions(Partition.None);
        options.setPartition(Partition.Any);
        options.setUserLocale(Locale.ENGLISH);
        options.setUserPartition(Partition.Any);
        AQLResultCollection results = Base.getService().executeQuery(query, options);
        
        // loop on results
        int cptRow = 1;
        while (results.next()) {
            HSSFRow row = userSheet.createRow(cptRow++);
            row.createCell(0).setCellValue(new HSSFRichTextString(results.getString(0)));
            row.createCell(1).setCellValue(new HSSFRichTextString(results.getString(1)));
            row.createCell(2).setCellValue(new HSSFRichTextString(results.getString(2)));
            row.createCell(3).setCellValue(new HSSFRichTextString(results.getString(3)));
            row.createCell(4).setCellValue(new HSSFRichTextString(results.getString(4)));
            row.createCell(5).setCellValue(new HSSFRichTextString(results.getString(5)));
            row.createCell(6).setCellValue(new HSSFRichTextString(results.getString(6)));
            row.createCell(7).setCellValue(new HSSFRichTextString(results.getString(7)));
            row.createCell(8).setCellValue(new HSSFRichTextString(results.getString(8)));

            String dptId = results.getString(9);
            if (!StringUtil.nullOrEmptyOrBlankString(dptId)) {
                String userDpt = StringUtil.strcat(dptId, " - ", results.getString(10));
                row.createCell(9).setCellValue(new HSSFRichTextString(userDpt));
            }
            
            String address = results.getString(11);
            if (!StringUtil.nullOrEmptyOrBlankString(address)) {
                String userAddress = StringUtil.strcat(address, " - ", results.getString(12));
                row.createCell(10).setCellValue(new HSSFRichTextString(userAddress));
            }
            
            row.createCell(11).setCellValue(new HSSFRichTextString(results.getString(13)));

        }
    }
    
}
