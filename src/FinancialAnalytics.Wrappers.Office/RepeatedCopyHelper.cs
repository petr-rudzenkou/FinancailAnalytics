using System;
using System.Windows.Forms;

namespace FinancialAnalytics.Wrappers.Office
{
    public static class RepeatedCopyHelper
    {
        private const int Delay = 10;
        private const int RepeatCount = 20;

        public static void ExecuteCopyRepeated(Action copyAction)
        {
            for (int i = 0; i < RepeatCount; i++)
            {
                //we can get exception in Office 2010 even during copying of an object
                try
                {
                    copyAction();
                }
                catch (Exception)
                {
                    Clipboard.Clear();
                }

                if (Clipboard.ContainsData(DataFormats.Text) || Clipboard.ContainsData(DataFormats.EnhancedMetafile)) // Check that any data present in Clipboard
                    break;

                System.Threading.Thread.CurrentThread.Join(Delay);
            }
        }

		public static void ExecutePasteRepeated(Action pasteAction)
		{
			for (int i = 0; i < RepeatCount; i++)
			{
				//we can get exception because of RDPclip even during pasting of an object
				//assumption is that after Paste fails we still have valid data in clipboard
				try
				{
					pasteAction();
					break;
				}
				catch (Exception)
				{
					//if we still have problems after trying - throw this exception
					if (i == RepeatCount - 1)
					{
						throw;
					}
				}

				System.Threading.Thread.CurrentThread.Join(Delay);
			}
		}
    }
}
