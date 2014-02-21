package config.sandbox.eform.supplier.action;

import ariba.approvable.core.Approvable;
import ariba.base.core.ClusterRoot;
import ariba.base.fields.Action;
import ariba.base.fields.ActionExecutionException;
import ariba.base.fields.ValueInfo;
import ariba.base.fields.ValueSource;
import ariba.common.core.Supplier;
import ariba.common.core.action.IntegrationPostLoadSupplier;
import ariba.user.core.Organization;
import ariba.util.core.PropertyTable;

/**
 * @author fhenri
 */
public class SupplierEFormApproved extends Action
{
    public static String className = SupplierEFormApproved.class.getName();

    /**
     * Entry point of the Action
     * 
     * @param object
     * @param params
     * @see ariba.base.fields.Action#fire(ariba.base.fields.ValueSource, ariba.util.core.PropertyTable)
     */
    public void fire (ValueSource obj, PropertyTable arg1)
    throws ActionExecutionException
    {
        ClusterRoot eform = (ClusterRoot) obj;
        createSupplier(eform);
    }

    private void createSupplier (ClusterRoot cr) {
        // create a new partitioned supplier
        Supplier supplier = createPartitionedSupplier(cr);

        // use post load integration event to link with common supplier
        IntegrationPostLoadSupplier updateSupplier = new IntegrationPostLoadSupplier();
        updateSupplier.fire(supplier, null);
    }

    /**
     * This method creates a new partitioned supplier with the minimum necessary values
     * based from values in eform and link to a common supplier.
     * @param eform the eform which contains the supplier values
     * @return a new partitioned supplier
     */
    private Supplier createPartitionedSupplier (ClusterRoot eform) {
        // ---------------------------------
        // create the new supplier
        // ---------------------------------
        Supplier supplier = new Supplier (eform.getPartition());

        // set header supplier values from eForm
        supplier.setUniqueName(eform.getUniqueName());
        supplier.setName((String)eform.getFieldValue("SupplierName"));
        
        supplier.setSupplierIDDomain(Organization.BuyerSystemID);
        supplier.setSupplierIDValue(eform.getUniqueName());

        // add the duns as a new organizationIdPart
        
        // ---------------------------------
        // save object
        // ---------------------------------
        supplier.save();

        return supplier;
    }

    /**
     * allowed types for the value
     */
    protected ValueInfo getValueInfo ()
    {
        return new ValueInfo(IsScalar, Approvable.ClassName);
    }
}