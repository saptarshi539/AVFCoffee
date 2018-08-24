using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Coffee.APIControllers;
using CoffeeCore.DTO;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication.Extensions;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using CoffeeInfrastructure.Helpers;
using Microsoft.Extensions.Options;

namespace AVFCoffeeWebApp.Controllers
{
    public class AccountController : Controller
    {
        //CellSumController cellSumController;

        public AccountController( IOptions<AzureAdB2COptions> b2cOptions)
        {
            //cellSumController = cellSum;
            Options = b2cOptions.Value;
        }

        public AzureAdB2COptions Options { get; set; }

        //
        // GET: /Account/SignIn
        [HttpGet]
        public IActionResult SignIn(string lang)
        {
            if(lang == "ES")
            {
                var properties = new AuthenticationProperties() { RedirectUri = "/TechnicianHome" };
                properties.Items[AzureAdB2COptions.PolicyAuthenticationProperty] = Options.SpanishSignUpSignInPolicyId;
                //if (User.Identity.IsAuthenticated)
                //{
                //    //var cooperativeID = User.GetCooperativeID();
                //    var username = User.GetGivenName();
                //    var userID = User.GetId();
                //    var language = User.GetSiupinPolicyName();
                //    UserInfoDTO user = new UserInfoDTO();
                //    user.Language = language;
                //    user.UserID = userID;
                //    user.UserName = username;
                //    cellSumController.UpdateUser(user);
                //}
                return Challenge(properties, OpenIdConnectDefaults.AuthenticationScheme);
                
            }
            else //(lang == "EN")
            {
                return Challenge(
                    new AuthenticationProperties { RedirectUri = "/TechnicianHome" }, OpenIdConnectDefaults.AuthenticationScheme);
            }
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
