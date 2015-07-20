using System.Text;

namespace FinancialAnalytics.Wrappers.Office
{
    public static class LocalPathToUncConverter
    {
        public static string Convert(string localPath)
        {
            if (string.IsNullOrEmpty(localPath))
            {
                return localPath;
            }
            string uncPath = localPath;
            if (localPath.Length >= 2 && localPath[1] == System.IO.Path.VolumeSeparatorChar)
            {
                char c = localPath[0];
                if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
                {
                    StringBuilder builder = new StringBuilder(1024);
                    int length = builder.Capacity;
                    int error = NativeMethods.WNetGetConnection(localPath.Substring(0, 2), builder, ref length);
                    if (error == 0)
                    {
                        string pathRoot = System.IO.Path.GetPathRoot(localPath);
                        if (string.IsNullOrEmpty(pathRoot))
                        {
                            return localPath;
                        }
                        string relativePath = localPath.Substring(pathRoot.Length);
                        uncPath = System.IO.Path.Combine(builder.ToString().TrimEnd(), relativePath);
                    }
                }
            }
            return uncPath;
        }
    }
}
