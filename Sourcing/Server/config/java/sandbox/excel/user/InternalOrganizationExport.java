/**
 * 
 */
package org.fhsolution.admin.excel.user;


/**
 *
 * @author fhenri
 *
 */
public class InternalOrganizationExport extends OrganizationExport {

    public InternalOrganizationExport(String sheetName) {
        super(sheetName);

        whereCriteria =
                " where o.IsSupplier = false";

    }

}
