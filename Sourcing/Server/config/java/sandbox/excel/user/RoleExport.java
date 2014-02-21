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
public class RoleExport implements ExcelExport {

    private String roleSheetName;
   
    protected HSSFSheet roleSheet;
    protected HSSFSheet roleMapSheet;
    protected HSSFSheet rolePermissionMapSheet;
    protected HSSFSheet roleUserMapSheet;
    protected HSSFSheet roleAllPermissionsSheet;
    protected HSSFSheet roleAllUsersSheet;
    
    protected List<ExcelTitleElement> roleSheetHeader;
    protected List<ExcelTitleElement> roleMapSheetHeader;
    protected List<ExcelTitleElement> rolePermissionSheetHeader;
    protected List<ExcelTitleElement> roleUserSheetHeader;

    /**
     * Public constructor
     */
    public RoleExport (String sheetName) {
        this.roleSheetName = sheetName;
        
        roleSheetHeader = new LinkedList<ExcelTitleElement>();
        roleSheetHeader.add(new ExcelTitleElement("UniqueName"));
        roleSheetHeader.add(new ExcelTitleElement("Name"));
        roleSheetHeader.add(new ExcelTitleElement("Description"));
        
        roleMapSheetHeader = new LinkedList<ExcelTitleElement>();
        roleMapSheetHeader.add(new ExcelTitleElement("ParentRole.UniqueName"));
        roleMapSheetHeader.add(new ExcelTitleElement("Role.UniqueName"));
        
        rolePermissionSheetHeader = new LinkedList<ExcelTitleElement>();
        rolePermissionSheetHeader.add(new ExcelTitleElement("ParentRole.UniqueName"));
        rolePermissionSheetHeader.add(new ExcelTitleElement("Permission.UniqueName"));
        
        roleUserSheetHeader = new LinkedList<ExcelTitleElement>();
        roleUserSheetHeader.add(new ExcelTitleElement("ParentRole.UniqueName"));
        roleUserSheetHeader.add(new ExcelTitleElement("User.UniqueName"));

        aqlText = BASE_QUERY;
    }
    
    /**
     */
    public void initializeSheet (HSSFWorkbook workbook) {
        roleSheet = workbook.createSheet(roleSheetName);
        ExcelUtil.initializeHeaderSheet(roleSheet, roleSheetHeader);
        
        roleMapSheet = workbook.createSheet("Role ChildRole Map");
        ExcelUtil.initializeHeaderSheet(roleMapSheet, roleMapSheetHeader);
        
        rolePermissionMapSheet = workbook.createSheet("Role Permission Map");
        ExcelUtil.initializeHeaderSheet(rolePermissionMapSheet, rolePermissionSheetHeader);
        
        roleUserMapSheet = workbook.createSheet("Role User Map");
        ExcelUtil.initializeHeaderSheet(roleUserMapSheet, roleUserSheetHeader);

        roleAllPermissionsSheet = workbook.createSheet("Role All Permissions Map");
        ExcelUtil.initializeHeaderSheet(roleAllPermissionsSheet, rolePermissionSheetHeader);
        
        roleAllUsersSheet = workbook.createSheet("Role All Users Map");
        ExcelUtil.initializeHeaderSheet(roleAllUsersSheet, roleUserSheetHeader);
    }

    protected String aqlText;
    protected static String BASE_QUERY = 
        "select r, r.UniqueName, r.Name, r.Description " +
        "from ariba.user.core.Role r subclass NONE order by r.UniqueName";

    // -- subquery --
    protected static String QUERY_USER = 
        "select r.UniqueName, u.UniqueName, u.Name, u.PasswordAdapter " +
        "from ariba.user.core.User u " +
        "left outer join ariba.user.core.Role r using u.Roles where r = BaseId('%s') " +
        "and u is not null order by u.UniqueName";

    protected static String QUERY_ROLE =
        "select r.UniqueName, cr.UniqueName, cr.Name " +
        "from ariba.user.core.Role r subclass NONE " +
        "left outer join ariba.user.core.Role cr using r.ChildRoles where r = BaseId('%s') " +
        "and cr is not null order by cr.UniqueName";
    
    protected static String QUERY_PERMISSION =
        "select r.UniqueName, p.UniqueName, p.Name " +
        "from ariba.user.core.Role r subclass NONE " +
        "left outer join ariba.user.core.Permission p using Permissions where r = BaseId('%s') " +
        "and p is not null order by p.UniqueName";
    
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
        int cptRoleRow = 1;
        int cptRoleUserRow = 1;
        int cptRoleMapRow = 1;
        int cptRolePermRow = 1;
        while (results.next()) {
            BaseId roleId = results.getBaseId(0);
            
            HSSFRow row = roleSheet.createRow(cptRoleRow++);
            row.createCell(0).setCellValue(new HSSFRichTextString(results.getString(1)));
            row.createCell(1).setCellValue(new HSSFRichTextString(results.getString(2)));
            row.createCell(2).setCellValue(new HSSFRichTextString(results.getString(3)));

            cptRoleUserRow = 
                ExcelUtil.writeSubQueryForGroup(QUERY_USER, roleId, roleUserMapSheet, cptRoleUserRow);
            cptRoleMapRow = 
                ExcelUtil.writeSubQueryForGroup(QUERY_ROLE, roleId, roleMapSheet, cptRoleMapRow);
            cptRolePermRow = 
                ExcelUtil.writeSubQueryForGroup(QUERY_PERMISSION, roleId, rolePermissionMapSheet, cptRolePermRow);
        }
    }
    /*
    public void processExport () {
        
        // query Group object
        AQLQuery query = new AQLQuery(Role.ClassName, true);
        AQLOptions options = new AQLOptions(Partition.None);
        options.setAssertOnMaxRowsReturned(false);
        AQLResultCollection results = Base.getService().executeQuery(query, options);
        
        // loop on results
        int cptRoleRow = 1;
        int cptRoleMapRow = 1;
        int cptRolePermRow = 1;
        int cptRoleUserRow = 1;
        int cptRoleAllPermsRow = 1;
        int cptRoleAllUsersRow = 1;
        while (results.next()) {
            Role role = (Role) results.getBaseId(0).get();
            
            // create a new row to write results
            HSSFRow roleRow = roleSheet.createRow(cptRoleRow++);
            writeRoleDataInRow(role, roleRow);

            String uniqueName = role.getUniqueName();
            cptRoleMapRow =
                ExcelUtil.writeIteratorData(uniqueName, role.getChildRolesIterator(), roleMapSheet, cptRoleMapRow);
            cptRolePermRow =
                ExcelUtil.writeIteratorData(uniqueName, role.getPermissionsIterator(), rolePermissionSheet, cptRolePermRow);
            cptRoleUserRow =
                ExcelUtil.writeIteratorData(uniqueName, role.getUsers().iterator(), roleUserSheet, cptRoleUserRow);
            cptRoleAllPermsRow =
                ExcelUtil.writeIteratorData(uniqueName, role.getAllPermissions().iterator(), roleAllPermissionsSheet, cptRoleAllPermsRow);
            cptRoleAllUsersRow =
                ExcelUtil.writeIteratorData(uniqueName, role.getAllUsers().iterator(), roleAllUsersSheet, cptRoleAllUsersRow);
        }
    }

    protected void writeRoleDataInRow (Role role, HSSFRow row) {
        row.createCell(0).setCellValue(new HSSFRichTextString(role.getUniqueName()));
        row.createCell(1).setCellValue(new HSSFRichTextString(role.getName().getPrimaryString()));
        if (role.getDescription() != null) {
            row.createCell(2).setCellValue(
                    new HSSFRichTextString(role.getDescription().getPrimaryString()));
        }
    }
    */
}