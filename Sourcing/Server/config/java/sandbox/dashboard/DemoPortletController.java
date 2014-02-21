package sandbox.dashboard;

import ariba.analytics.widgets.AnalysisPortletController;
import ariba.user.core.User;

/**
 *
 * @author fhenri
 */
public class DemoPortletController extends AnalysisPortletController {

    /**
     * @see ariba.analytics.widgets.AnalysisPortletController#isUserAuthorizedForContent(ariba.user.core.User, ariba.user.core.User)
     */
    protected boolean isUserAuthorizedForContent(User realUser, User effectiveUser) {
        return super.isUserAuthorizedForContent(realUser, effectiveUser);
    }

	/**
	 * @see ariba.portlet.component.PortletController#editComponentName()
	 */
	@Override
	protected String editComponentName() {
		return null;
	}

	/**
	 * @see ariba.portlet.component.PortletController#viewComponentName()
	 */
	@Override
	protected String viewComponentName() {
		return DemoPortletContent.class.getName();
	}

	
    /**
     * @see ariba.portlet.component.PortletController#portletTypeName()
     */
    protected String portletTypeName () {
        return "Ariba Portlet Demo";
    }

}
