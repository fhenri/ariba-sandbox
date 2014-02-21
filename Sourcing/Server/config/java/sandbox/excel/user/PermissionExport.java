/**
 * 
 */
package org.fhsolution.admin.excel.user;

import ariba.base.core.Base;
import ariba.base.core.BaseId;
import ariba.base.core.Partition;
import ariba.base.core.aql.AQLOptions;
import ariba.base.core.aql.AQLQuery;
import ariba.base.core.aql.AQLResultCollection;
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
public class PermissionExport implements ExcelExport {

    protected HSSFSheet permissionSheet;
    protected HSSFSheet permissionAllUsersSheet;
    
    protected String permissionsheetName;
   
    protected List<ExcelTitleElement> permissionSheetHeader;
    protected List<ExcelTitleElement> permissionAllUsersSheetHeader;
    
    /**
     * Public constructor
     */
    public PermissionExport (String sheetName) {
        this.permissionsheetName = sheetName;
    
        permissionSheetHeader = new LinkedList<ExcelTitleElement>();
        permissionSheetHeader.add(new ExcelTitleElement("UniqueName"));
        permissionSheetHeader.add(new ExcelTitleElement("Name"));
        permissionSheetHeader.add(new ExcelTitleElement("Description"));

        permissionAllUsersSheetHeader = new LinkedList<ExcelTitleElement>();
        permissionAllUsersSheetHeader.add(new ExcelTitleElement("Permission.UniqueName"));
        permissionAllUsersSheetHeader.add(new ExcelTitleElement("User.UniqueName"));

        aqlText = BASE_QUERY;
    }
    
    /**
     */
    public void initializeSheet (HSSFWorkbook workbook) {
        permissionSheet = workbook.createSheet(permissionsheetName);
        ExcelUtil.initializeHeaderSheet(permissionSheet, permissionSheetHeader);

        permissionAllUsersSheet = workbook.createSheet("Permission User Map");
        ExcelUtil.initializeHeaderSheet(permissionAllUsersSheet, permissionAllUsersSheetHeader);
    }

    protected String aqlText;
    protected static String BASE_QUERY = 
        "select p, p.UniqueName, p.Name, p.Description " +
        "from ariba.user.core.Permission p subclass NONE order by p.UniqueName";

    // -- subquery --
    protected static String QUERY_USER = 
        "select p.UniqueName, u.UniqueName, u.Name, u.PasswordAdapter " +
        "from ariba.user.core.User u " +
        "left outer join ariba.user.core.Permission p using u.Permissions where p = BaseId('%s') " +
        "and u is not null order by u.UniqueName";

    /**
     */
    public void processExport () {
        
        AQLOptions options = new AQLOptions(Partition.None);
        options.setPartition(Partition.Any);
        options.setUserLocale(Locale.ENGLISH);
        options.setUserPartition(Partition.Any);
        
        AQLQuery query = AQLQuery.parseQuery(aqlText);
        AQLResultCollection results = Base.getService().executeQuery(query, options);
        
        // loop on results
        int cptPermissionRow = 1;
        int cptPermUserRow = 1;
        while (results.next()) {
            BaseId permId = results.getBaseId(0);
            
            HSSFRow row = permissionSheet.createRow(cptPermissionRow++);
            row.createCell(0).setCellValue(new HSSFRichTextString(results.getString(1)));
            row.createCell(1).setCellValue(new HSSFRichTextString(results.getString(2)));
            row.createCell(2).setCellValue(new HSSFRichTextString(results.getString(3)));
    
            cptPermUserRow = 
                ExcelUtil.writeSubQueryForGroup(QUERY_USER, permId, permissionAllUsersSheet, cptPermUserRow);
        }
    }
    
    /*
    public void processExport () {
        
        // query Group object
        AQLQuery query = new AQLQuery(Permission.ClassName);
        AQLOptions options = new AQLOptions(Partition.None);
        AQLResultCollection results = Base.getService().executeQuery(query, options);
        
        // loop on results
        int cptPermissionRow = 1;
        //int cptPermissionAllUsersRow = 1;
        while (results.next()) {
            Permission perm = (Permission) results.getBaseId(0).get();
            
            // create a new row to write results
            HSSFRow permissionRow = permissionSheet.createRow(cptPermissionRow++);
            writePermissionDataInRow(perm, permissionRow);
            cptPermissionAllUsersRow =
                ExcelUtil.writeIteratorData(perm.getUniqueName(), perm.getAllUsers().iterator(), permissionAllUsersSheet, cptPermissionAllUsersRow);
        }
    }

    protected void writePermissionDataInRow (Permission perm, HSSFRow row) {
        row.createCell(0).setCellValue(new HSSFRichTextString(perm.getUniqueName()));
        row.createCell(1).setCellValue(new HSSFRichTextString(perm.getName().getPrimaryString()));
        if (perm.getDescription() != null) {
            row.createCell(2).setCellValue(
                    new HSSFRichTextString(perm.getDescription().getPrimaryString()));
        }
    }
    */
}