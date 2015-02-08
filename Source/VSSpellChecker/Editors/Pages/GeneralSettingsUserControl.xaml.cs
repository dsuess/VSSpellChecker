//===============================================================================================================
// System  : Visual Studio Spell Checker Package
// File    : GeneralSettingsUserControl.xaml.cs
// Author  : Eric Woodruff  (Eric@EWoodruff.us)
// Updated : 02/08/2015
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

using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

using VisualStudio.SpellChecker.Configuration;

namespace VisualStudio.SpellChecker.Editors.Pages
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
            get { return "General Settings"; }
        }

        /// <inheritdoc />
        public string HelpUrl
        {
            get { return this.Title; }
        }

        /// <inheritdoc />
        public void LoadConfiguration(SpellingConfigurationFile configuration)
        {
            var dataSource = new List<PropertyState>();

            if(configuration.ConfigurationType != ConfigurationType.Global)
                dataSource.AddRange(new[] { PropertyState.Inherited, PropertyState.Yes, PropertyState.No });
            else
                dataSource.AddRange(new[] { PropertyState.Yes, PropertyState.No });

            cboSpellCheckAsYouType.ItemsSource = cboIgnoreWordsWithDigits.ItemsSource =
                cboIgnoreAllUppercase.ItemsSource = cboIgnoreFormatSpecifiers.ItemsSource =
                cboIgnoreFilenamesAndEMail.ItemsSource = cboIgnoreXmlInText.ItemsSource =
                cboTreatUnderscoresAsSeparators.ItemsSource = dataSource;

            cboSpellCheckAsYouType.SelectedValue = configuration.ToPropertyState(
                PropertyNames.SpellCheckAsYouType);
            cboIgnoreWordsWithDigits.SelectedValue = configuration.ToPropertyState(
                PropertyNames.IgnoreWordsWithDigits);
            cboIgnoreAllUppercase.SelectedValue = configuration.ToPropertyState(
                PropertyNames.IgnoreWordsInAllUppercase);
            cboIgnoreFormatSpecifiers.SelectedValue = configuration.ToPropertyState(
                PropertyNames.IgnoreFormatSpecifiers);
            cboIgnoreFilenamesAndEMail.SelectedValue = configuration.ToPropertyState(
                PropertyNames.IgnoreFilenamesAndEMailAddresses);
            cboIgnoreXmlInText.SelectedValue = configuration.ToPropertyState(
                PropertyNames.IgnoreXmlElementsInText);
            cboTreatUnderscoresAsSeparators.SelectedValue = configuration.ToPropertyState(
                PropertyNames.TreatUnderscoreAsSeparator);

            chkInheritExcludeByExt.Visibility = (configuration.ConfigurationType != ConfigurationType.Global) ?
                Visibility.Visible : Visibility.Collapsed;
            chkInheritExcludeByExt.IsChecked = !configuration.HasProperty(PropertyNames.ExcludeByFilenameExtension);

            if(!chkInheritExcludeByExt.IsChecked.Value)
                txtExcludeByExtension.Text = configuration.ToString(PropertyNames.ExcludeByFilenameExtension);
        }

        /// <inheritdoc />
        public void SaveConfiguration(SpellingConfigurationFile configuration)
        {
            configuration.StoreProperty(PropertyNames.SpellCheckAsYouType,
                ((PropertyState)cboSpellCheckAsYouType.SelectedValue).ToPropertyValue());
            configuration.StoreProperty(PropertyNames.IgnoreWordsWithDigits,
                ((PropertyState)cboIgnoreWordsWithDigits.SelectedValue).ToPropertyValue());
            configuration.StoreProperty(PropertyNames.IgnoreWordsInAllUppercase,
                ((PropertyState)cboIgnoreAllUppercase.SelectedValue).ToPropertyValue());
            configuration.StoreProperty(PropertyNames.IgnoreFormatSpecifiers,
                ((PropertyState)cboIgnoreFormatSpecifiers.SelectedValue).ToPropertyValue());
            configuration.StoreProperty(PropertyNames.IgnoreFilenamesAndEMailAddresses,
                ((PropertyState)cboIgnoreFilenamesAndEMail.SelectedValue).ToPropertyValue());
            configuration.StoreProperty(PropertyNames.IgnoreXmlElementsInText,
                ((PropertyState)cboIgnoreXmlInText.SelectedValue).ToPropertyValue());
            configuration.StoreProperty(PropertyNames.TreatUnderscoreAsSeparator,
                ((PropertyState)cboTreatUnderscoresAsSeparators.SelectedValue).ToPropertyValue());

            // If not inherited, write it out even if blank.  We may want to turn the filter off lower down in
            // the chain (i.e. exclude globally but not in a specific solution or project).
            configuration.StoreProperty(PropertyNames.ExcludeByFilenameExtension,
                chkInheritExcludeByExt.IsChecked.Value ? null : txtExcludeByExtension.Text.Trim());
        }

        /// <inheritdoc />
        public event EventHandler ConfigurationChanged;

        #endregion

        #region Event handlers
        //=====================================================================

        /// <summary>
        /// Enable or disable the excluded extensions checkbox base on the Inherited checkbox state
        /// </summary>
        /// <param name="sender">The sender of the event</param>
        /// <param name="e">The event arguments</param>
        private void chkInheritExcludeByExt_CheckedChanged(object sender, System.Windows.RoutedEventArgs e)
        {
            txtExcludeByExtension.IsEnabled = !chkInheritExcludeByExt.IsChecked.Value;
            Property_Changed(sender, e);
        }

        /// <summary>
        /// Notify the parent of property changes that affect the file's dirty state
        /// </summary>
        /// <param name="sender">The sender of the event</param>
        /// <param name="e">The event arguments</param>
        private void Property_Changed(object sender, System.Windows.RoutedEventArgs e)
        {
            var handler = ConfigurationChanged;

            if(handler != null)
                handler(this, EventArgs.Empty);
        }
        #endregion
    }
}
