using System.Net;

namespace FinancialAnalytics.AuthenticationClient
{
    public interface IAuthenticationClient
    {
        bool Login(NetworkCredential credential);
        bool Logout();
        bool IsOnline { get; }
        event UserInfoUpdated UserInfoUpdated;
    }
}
