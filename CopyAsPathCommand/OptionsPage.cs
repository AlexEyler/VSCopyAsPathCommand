using System.ComponentModel;
using Microsoft.VisualStudio.Shell;

namespace CopyAsPathCommand
{
    internal class OptionsPage : DialogPage
    {
        internal const bool DefaultQuotesOption = true;
        private bool? quotesOption;

        [Category("Command Options")]
        [DisplayName("Include quotes in path")]
        [Description("If true, quotes will surround the copied path.")]
        public bool QuotesOption
        {
            get { return this.quotesOption ?? OptionsPage.DefaultQuotesOption; }
            set { this.quotesOption = value; }
        }
    }
}
