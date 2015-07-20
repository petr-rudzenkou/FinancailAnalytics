using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using Caliburn.Micro;
using FinancialAnalytics.AuthenticationClient;
using FinancialAnalytics.Views.Login.Interfaces;

namespace FinancialAnalytics.Views.Login
{
    public class LoginViewModel : Screen, ILoginViewModel
    {
        private readonly IAuthenticationClient _authenticationClient;
        
        public LoginViewModel(IAuthenticationClient authenticationClient)
        {
            _authenticationClient = authenticationClient;
            DisplayName = Resources.ViewsResources.Login_WindowTitle;
        }

        public void Login(string userId, string password)
        {
            bool authenticated = _authenticationClient.Login(new NetworkCredential()
            {
                UserName = userId,
                Password = password
            });

            if(authenticated)
                TryClose();
        }
    }
}
