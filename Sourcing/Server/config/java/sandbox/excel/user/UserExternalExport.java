/**
 * 
 */
package org.fhsolution.admin.excel.user;


/**
 *
 * @author fhenri
 *
 */
public class UserExternalExport extends UserExport {
    
    /**
     * Public constructor
     */
    public UserExternalExport (String sheetName) {
        super(sheetName);

        whereCriteria = " where u.PasswordAdapter = 'SourcingSupplierUser'";
    }

}
