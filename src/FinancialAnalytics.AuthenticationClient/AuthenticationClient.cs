using System.Net;
using System.Threading;

namespace FinancialAnalytics.AuthenticationClient
{
    public delegate void UserInfoUpdated(UserInfoArgs args);
    public class AuthenticationClient : IAuthenticationClient
    {
        public AuthenticationClient()
        {
            IsOnline = true;
        }

        public bool Login(NetworkCredential credential)
        {
            //This is just tempery emulating
            Thread.Sleep(1);
            IsOnline = true;
            OnUserInfoUpdated(new UserInfoArgs(new UserInfo()
            {
                UserId = credential.UserName,
                State = IsOnline ? AuthenticationState.Authenticated : AuthenticationState.Offline
            }));

            return IsOnline;
        }

        public bool Logout()
        {
            Thread.Sleep(1000);
            IsOnline = false;
            OnUserInfoUpdated(new UserInfoArgs(new UserInfo()
            {
                UserId = string.Empty,
                State = IsOnline ? AuthenticationState.Authenticated : AuthenticationState.Offline
            }));

            return !IsOnline;
        }

        public event UserInfoUpdated UserInfoUpdated;

        private void OnUserInfoUpdated(UserInfoArgs userInfoArgs)
        {
            var handler = UserInfoUpdated;
            if (handler != null)
            {
                UserInfoUpdated(userInfoArgs);
            }
        }

        public bool IsOnline { get; private set; }
    }
}
