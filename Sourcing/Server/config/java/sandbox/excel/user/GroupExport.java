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
import ariba.util.core.ResourceService;
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
 *
 * @author fhenri
 *
 */
public class GroupExport implements ExcelExport {

    private static final String GroupDescriptionTable = "ariba.asm.main.GroupDescriptions";

    protected HSSFSheet groupSheet;
    protected HSSFSheet groupUserMapSheet;
    protected HSSFSheet groupChildGroupSheet;
    protected HSSFSheet groupRoleMapSheet;
    protected HSSFSheet groupPermissionMapSheet;
    protected HSSFSheet groupAllUsersSheet;
    protected HSSFSheet groupAllGroupsSheet;
    protected HSSFSheet groupAllRolesSheet;
    protected HSSFSheet groupAllPermissionsSheet;

    private String sheetName;

    protected List<ExcelTitleElement> groupSheetHeader;
    protected List<ExcelTitleElement> groupParentSheetHeader;
    protected List<ExcelTitleElement> groupRoleSheetHeader;
    protected List<ExcelTitleElement> groupPermissionSheetHeader;
    protected List<ExcelTitleElement> groupUserSheetHeader;

    /**
     * Public constructor
     */
    public GroupExport (String sheetName) {
        this.sheetName = sheetName;

        groupSheetHeader = new LinkedList<ExcelTitleElement>();
        groupSheetHeader.add(new ExcelTitleElement("AdapterSource"));
        groupSheetHeader.add(new ExcelTitleElement("UniqueName"));
        groupSheetHeader.add(new ExcelTitleElement("Name"));
        groupSheetHeader.add(new ExcelTitleElement("Description"));

        groupParentSheetHeader = new LinkedList<ExcelTitleElement>();
        groupParentSheetHeader.add(new ExcelTitleElement("ParentGroup.UniqueName"));
        groupParentSheetHeader.add(new ExcelTitleElement("Group.UniqueName"));

        groupRoleSheetHeader = new LinkedList<ExcelTitleElement>();
        groupRoleSheetHeader.add(new ExcelTitleElement("ParentGroup.UniqueName"));
        groupRoleSheetHeader.add(new ExcelTitleElement("Role.UniqueName"));

        groupPermissionSheetHeader = new LinkedList<ExcelTitleElement>();
        groupPermissionSheetHeader.add(new ExcelTitleElement("ParentGroup.UniqueName"));
        groupPermissionSheetHeader.add(new ExcelTitleElement("Permission.UniqueName"));

        groupUserSheetHeader = new LinkedList<ExcelTitleElement>();
        groupUserSheetHeader.add(new ExcelTitleElement("ParentGroup.UniqueName"));
        groupUserSheetHeader.add(new ExcelTitleElement("User.UniqueName"));

        aqlText = BASE_QUERY;
    }

    /**
     */
    public void initializeSheet (HSSFWorkbook workbook) {
        groupSheet = workbook.createSheet(sheetName);
        ExcelUtil.initializeHeaderSheet(groupSheet, groupSheetHeader);

        groupUserMapSheet = workbook.createSheet("Group User Map");
        ExcelUtil.initializeHeaderSheet(groupUserMapSheet, groupUserSheetHeader);

        groupChildGroupSheet = workbook.createSheet("Group Child Group");
        ExcelUtil.initializeHeaderSheet(groupChildGroupSheet, groupParentSheetHeader);

        groupRoleMapSheet = workbook.createSheet("Group Role Map");
        ExcelUtil.initializeHeaderSheet(groupRoleMapSheet, groupRoleSheetHeader);

        groupPermissionMapSheet = workbook.createSheet("Group Permission Map");
        ExcelUtil.initializeHeaderSheet(groupPermissionMapSheet, groupPermissionSheetHeader);

        groupAllUsersSheet = workbook.createSheet("Group All Users Map");
        ExcelUtil.initializeHeaderSheet(groupAllUsersSheet, groupUserSheetHeader);

        groupAllGroupsSheet = workbook.createSheet("Group All Groups Map");
        ExcelUtil.initializeHeaderSheet(groupAllGroupsSheet, groupParentSheetHeader);

        groupAllRolesSheet = workbook.createSheet("Group All Roles Map");
        ExcelUtil.initializeHeaderSheet(groupAllRolesSheet, groupRoleSheetHeader);

        groupAllPermissionsSheet = workbook.createSheet("Group All Permissions Map");
        ExcelUtil.initializeHeaderSheet(groupAllPermissionsSheet, groupPermissionSheetHeader);
    }

    protected String aqlText;
    protected static String BASE_QUERY =
        "select g, g.AdapterSource, g.UniqueName, g.Name, g.Description " +
        "from ariba.user.core.Group g subclass NONE order by g.UniqueName";

    // -- subquery --
    protected static String QUERY_USER =
        "select g.UniqueName, u.UniqueName, u.Name, u.PasswordAdapter " +
        "from ariba.user.core.Group g subclass NONE " +
        "left outer join ariba.user.core.User u using g.Users where g = BaseId('%s') " +
        "and u is not null order by u.UniqueName";

    protected static String QUERY_ROLE =
        "select g.UniqueName, r.UniqueName, r.Name from ariba.user.core.Group g subclass NONE " +
        "left outer join ariba.user.core.Role r using g.Roles where g = BaseId('%s') " +
        "and r is not null order by r.UniqueName";

    protected static String QUERY_PERMISSION =
        "select g.UniqueName, p.UniqueName, p.Name from ariba.user.core.Group g subclass NONE " +
        "left outer join ariba.user.core.Permission p using g.Permissions where g = BaseId('%s') " +
        "and p is not null order by p.UniqueName";

    protected static String QUERY_CHILDGROUP =
        "select g.UniqueName, cg.UniqueName, cg.Name from ariba.user.core.Group g subclass NONE " +
        "left outer join ariba.user.core.Group cg using g.ChildGroups where g = BaseId('%s') " +
        "and cg is not null order by cg.UniqueName";
    /*
    select fg.UniqueName, fg.Name from ariba.user.core.Group g subclass NONE
    left outer join ariba.user.core.Group fg using g.FlattenedGroups where g = %s
    */

    protected static String QUERY_ALLROLE =
        "SELECT p.Roles FROM ariba.user.core.Group AS p WHERE p = baseid('%s') " +
        "UNION SELECT p.Roles.FlattenedRoles FROM ariba.user.core.Group AS p WHERE p = baseid('%s') " +
        "UNION SELECT g2.Roles.FlattenedRoles FROM ariba.user.core.Group AS g1 SUBCLASS NONE, ariba.user.core.Group AS g2 SUBCLASS NONE WHERE g1 = baseid('%s') AND g2.FlattenedGroups = g1 AND g2.IsGlobal = TRUE " +
        "UNION SELECT g2.Roles FROM ariba.user.core.Group AS g1 SUBCLASS NONE, ariba.user.core.Group AS g2 SUBCLASS NONE WHERE g1 = baseid('%s') AND g2.FlattenedGroups = g1 AND g2.IsGlobal = TRUE";

    protected static String QUERY_ALLPERMISSION =
        "SELECT g.Permissions FROM ariba.user.core.Group AS g WHERE g = baseid('%s') AND g.Permissions.Active = 'TRUE' " +
        "UNION SELECT g.Roles.FlattenedRoles.Permissions FROM ariba.user.core.Group AS g WHERE g = baseid('%s') AND g.Roles.FlattenedRoles.Permissions.Active = 'TRUE' " +
        "UNION SELECT g2.Permissions FROM ariba.user.core.Group AS g1 SUBCLASS NONE, ariba.user.core.Group AS g2 SUBCLASS NONE WHERE g1 = baseid('%s') AND g2.FlattenedGroups = g1 AND g2.IsGlobal = 'TRUE' AND g2.Permissions.Active = 'TRUE' " +
        "UNION SELECT g2.Roles.FlattenedRoles.Permissions FROM ariba.user.core.Group AS g1 SUBCLASS NONE, ariba.user.core.Group AS g2 SUBCLASS NONE WHERE g1 = baseid('%s') AND g2.FlattenedGroups = g1 AND g2.IsGlobal = 'TRUE' AND g2.Roles.FlattenedRoles.Permissions.Active = 'TRUE'";

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
        int cptGroupRow = 1;
        int cptGroupUserRow = 1;
        int cptGroupRoleRow = 1;
        int cptGroupPermRow = 1;
        int cptChildGroupRow = 1;
        while (results.next()) {
            BaseId groupId = results.getBaseId(0);

            HSSFRow row = groupSheet.createRow(cptGroupRow++);
            // write adapter source
            row.createCell(0).setCellValue(new HSSFRichTextString(results.getString(1)));
            // write unique name
            row.createCell(1).setCellValue(new HSSFRichTextString(results.getString(2)));
            // write name
            row.createCell(2).setCellValue(new HSSFRichTextString(results.getString(3)));
            // write description or pull from csv file
            //row.createCell(3).setCellValue(new HSSFRichTextString(results.getString(4)));
            String description = results.getString(4);
            if (StringUtil.nullOrEmptyOrBlankString(description)) {
                description = ResourceService.getString(GroupDescriptionTable, results.getString(1));
            }
            row.createCell(3).setCellValue(new HSSFRichTextString(description));

            cptGroupUserRow =
                ExcelUtil.writeSubQueryForGroup(QUERY_USER, groupId, groupUserMapSheet, cptGroupUserRow);
            cptGroupRoleRow =
                ExcelUtil.writeSubQueryForGroup(QUERY_ROLE, groupId, groupRoleMapSheet, cptGroupRoleRow);
            cptGroupPermRow =
                ExcelUtil.writeSubQueryForGroup(QUERY_PERMISSION, groupId, groupPermissionMapSheet, cptGroupPermRow);
            cptChildGroupRow =
                ExcelUtil.writeSubQueryForGroup(QUERY_CHILDGROUP, groupId, groupChildGroupSheet, cptChildGroupRow);

            // getAllPermission
            //Group.getAllRolesQuery(groupId.get())
        }
    }
}