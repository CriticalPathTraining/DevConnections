using System.Web.Mvc;
using AcsSPOWeb.Models;

namespace AcsSPOWeb.Filters {
  public class SharePointSessionManagerFilterAttribute : ActionFilterAttribute {

    public override void OnActionExecuting(ActionExecutingContext filterContext) {
      SharePointSessionManager.InitializeRequest(filterContext.Controller);
    }

  }
}