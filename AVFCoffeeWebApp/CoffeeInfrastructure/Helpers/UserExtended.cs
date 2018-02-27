using CoffeeInfrastructure.B2C;
using System;
using System.Collections.Generic;
using System.Security.Claims;
using System.Security.Principal;
using System.Text;

namespace CoffeeInfrastructure.Helpers
{
    public static class UserExtended
    {
        public static string GetGivenName(this IPrincipal user)
        {
            var claim = ((ClaimsIdentity)user.Identity).FindFirst(ClaimTypes.GivenName);
            return claim == null ? null : claim.Value;
        }
        public static string GetSurname(this IPrincipal user)
        {
            var claim = ((ClaimsIdentity)user.Identity).FindFirst(ClaimTypes.Surname);
            return claim == null ? null : claim.Value;
        }
        public static string GetCountry(this IPrincipal user)
        {
            var claim = ((ClaimsIdentity)user.Identity).FindFirst(B2cClaims.Country);
            return claim == null ? null : claim.Value;
        }
        public static string GetState(this IPrincipal user)
        {
            var claim = ((ClaimsIdentity)user.Identity).FindFirst(B2cClaims.State);
            return claim == null ? null : claim.Value;
        }
        public static string GetCity(this IPrincipal user)
        {
            var claim = ((ClaimsIdentity)user.Identity).FindFirst(B2cClaims.City);
            return claim == null ? null : claim.Value;
        }
        public static string GetId(this IPrincipal user)
        {
            var claim = ((ClaimsIdentity)user.Identity).FindFirst(B2cClaims.ObjectId);
            return claim == null ? null : claim.Value;
        }
        public static string GetJob(this IPrincipal user)
        {
            var claim = ((ClaimsIdentity)user.Identity).FindFirst(B2cClaims.Job);
            return claim == null ? null : claim.Value;
        }
        // TODO: Get all emails
        public static string GetEmails(this IPrincipal user)
        {
            var claim = ((ClaimsIdentity)user.Identity).FindFirst(B2cClaims.Emails);
            return claim == null ? null : claim.Value;
        }
        public static string GetSiupinPolicyName(this IPrincipal user)
        {
            var claim = ((ClaimsIdentity)user.Identity).FindFirst(B2cClaims.SiupinPolicy);
            return claim == null ? null : claim.Value;
        }
        public static string GetCooperativeID(this IPrincipal user)
        {
            var claim = ((ClaimsIdentity)user.Identity).FindFirst(B2cClaims.CooperativeID);
            return claim == null ? null : claim.Value;
        }
    }
}
