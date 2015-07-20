using System;

namespace FinancialAnalytics.AuthenticationClient
{
    public class UserInfoArgs : EventArgs
    {
        public UserInfoArgs(UserInfo userInfo)
        {
            UserInfo = userInfo;
        }
        public UserInfo UserInfo { get; private set; }
    }
}
