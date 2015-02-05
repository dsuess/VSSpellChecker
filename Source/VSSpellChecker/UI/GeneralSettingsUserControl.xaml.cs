//===============================================================================================================
// System  : Visual Studio Spell Checker Package
// File    : GeneralSettingsUserControl.xaml.cs
// Author  : Eric Woodruff  (Eric@EWoodruff.us)
// Updated : 02/02/2015
// Note    : Copyright 2014-2015, Eric Woodruff, All rights reserved
// Compiler: Microsoft Visual C#
//
// This file contains a user control used to edit the general spell checker configuration settings
//
// This code is published under the Microsoft Public License (Ms-PL).  A copy of the license should be
// distributed with the code and can be found at the project website: https://github.com/EWSoftware/VSSpellChecker
// This notice, the author's name, and all copyright notices must remain intact in all applications,
// documentation, and source files.
//
//    Date     Who  Comments
// ==============================================================================================================
// 06/09/2014  EFW  Moved the general settings to a user control
//===============================================================================================================

using System.Windows.Controls;

using VisualStudio.SpellChecker.Configuration;

namespace VisualStudio.SpellChecker.UI
{
    /// <summary>
    /// This user control is used to edit the general spell checker configuration settings
    /// </summary>
    public partial class GeneralSettingsUserControl : UserControl, ISpellCheckerConfiguration
    {
        #region Constructor
        //=====================================================================

        /// <summary>
        /// Constructor
        /// </summary>
        public GeneralSettingsUserControl()
        {
            InitializeComponent();
        }
        #endregion

        #region ISpellCheckerConfiguration Members
        //=====================================================================

        /// <inheritdoc />
        public UserControl Control
        {
            get { return this; }
        }

        /// <inheritdoc />
        public string Title
        {
            get { return "General-Settings"; }
        }

        /// <inheritdoc />
        public string HelpUrl
        {
            get { return this.Title; }
        }

        /// <inheritdoc />
        public bool IsValid
        {
            get { return true; }
        }

        /// <inheritdoc />
        public void LoadConfiguration(SpellingConfigurationFile configuration)
        {
            chkSpellCheckAsYouType.IsChecked = configuration.ToBoolean(PropertyNames.SpellCheckAsYouType);
            chkIgnoreWordsWithDigits.IsChecked = configuration.ToBoolean(PropertyNames.IgnoreWordsWithDigits);
            chkIgnoreAllUppercase.IsChecked = configuration.ToBoolean(PropertyNames.IgnoreWordsInAllUppercase);
            chkIgnoreFormatSpecifiers.IsChecked = configuration.ToBoolean(PropertyNames.IgnoreFormatSpecifiers);
            chkIgnoreFilenamesAndEMail.IsChecked = configuration.ToBoolean(
                PropertyNames.IgnoreFilenamesAndEMailAddresses);
            chkIgnoreXmlInText.IsChecked = configuration.ToBoolean(PropertyNames.IgnoreXmlElementsInText);
            chkTreatUnderscoresAsSeparators.IsChecked = configuration.ToBoolean(
                PropertyNames.TreatUnderscoreAsSeparator);

            txtExcludeByExtension.Text = configuration.ToString(PropertyNames.ExcludeByFilenameExtension);
        }

        /// <inheritdoc />
        public bool SaveConfiguration(SpellingConfigurationFile configuration)
        {
            configuration.StoreProperty(PropertyNames.SpellCheckAsYouType,
                chkSpellCheckAsYouType.IsChecked.Value);
            configuration.StoreProperty(PropertyNames.IgnoreWordsWithDigits,
                chkIgnoreWordsWithDigits.IsChecked.Value);
            configuration.StoreProperty(PropertyNames.IgnoreWordsInAllUppercase,
                chkIgnoreAllUppercase.IsChecked.Value);
            configuration.StoreProperty(PropertyNames.IgnoreFormatSpecifiers,
                chkIgnoreFormatSpecifiers.IsChecked.Value);
            configuration.StoreProperty(PropertyNames.IgnoreFilenamesAndEMailAddresses,
                chkIgnoreFilenamesAndEMail.IsChecked.Value);
            configuration.StoreProperty(PropertyNames.IgnoreXmlElementsInText,
                chkIgnoreXmlInText.IsChecked.Value);
            configuration.StoreProperty(PropertyNames.TreatUnderscoreAsSeparator,
                chkTreatUnderscoresAsSeparators.IsChecked.Value);
            configuration.StoreProperty(PropertyNames.ExcludeByFilenameExtension,
                txtExcludeByExtension.Text);

            return true;
        }
        #endregion
    }
}
