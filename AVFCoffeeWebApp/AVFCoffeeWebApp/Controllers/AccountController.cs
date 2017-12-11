using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication.Extensions;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace AVFCoffeeWebApp.Controllers
{
    public class AccountController : Controller
    {
        public AccountController(IOptions<AzureAdB2COptions> b2cOptions)
        {
            Options = b2cOptions.Value;
        }

        public AzureAdB2COptions Options { get; set; }

        //
        // GET: /Account/SignIn
        [HttpGet]
        public IActionResult SignIn()
        {
            return Challenge(
                new AuthenticationProperties { RedirectUri = "/home" }, OpenIdConnectDefaults.AuthenticationScheme);
        }

        [HttpGet]
        public IActionResult ResetPassword()
        {
            var properties = new AuthenticationProperties() { RedirectUri = "/" };
            properties.Items[AzureAdB2COptions.PolicyAuthenticationProperty] = Options.ResetPasswordPolicyId;
            return Challenge(properties, OpenIdConnectDefaults.AuthenticationScheme);
        }

        [HttpGet]
        public IActionResult EditProfile()
        {
            var properties = new AuthenticationProperties() { RedirectUri = "/" };
            properties.Items[AzureAdB2COptions.PolicyAuthenticationProperty] = Options.EditProfilePolicyId;
            return Challenge(properties, OpenIdConnectDefaults.AuthenticationScheme);
        }
        
        //
        // GET: /Account/SignOut
        [HttpGet]
        public IActionResult SignOut()
        {
            return SignOut(new AuthenticationProperties { RedirectUri = "/" },
                CookieAuthenticationDefaults.AuthenticationScheme, OpenIdConnectDefaults.AuthenticationScheme);
        }

        //
        // GET: /Account/SignedOut
        [HttpGet]
        public IActionResult SignedOut()
        {
            if (HttpContext.User.Identity.IsAuthenticated)
            {
                // Redirect to home page if the user is authenticated.
                return RedirectToAction(nameof(HomeController.Index), "Home");
            }

            return View();
        }

        #region Helpers

        private void AddErrors(IdentityResult result)
        {
            foreach (var error in result.Errors)
            {
                ModelState.AddModelError(string.Empty, error.Description);
            }
        }

        private IActionResult RedirectToLocal(string returnUrl)
        {
            if (Url.IsLocalUrl(returnUrl))
            {
                return Redirect(returnUrl);
            }
            else
            {
                return RedirectToAction(nameof(HomeController.Index), "Home");
            }
        }

        #endregion
    }
}
