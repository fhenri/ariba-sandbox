/**
 * 
 */
package org.fhsolution.admin.excel.user;


/**
 *
 * @author fhenri
 *
 */
public class UserInternalExport extends UserExport {
    
    /**
     * Public constructor
     */
    public UserInternalExport (String sheetName) {
        super(sheetName);

        whereCriteria = " where u.PasswordAdapter = 'PasswordAdapter1'";
    }

}
