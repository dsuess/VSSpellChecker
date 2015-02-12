//===============================================================================================================
// System  : Visual Studio Spell Checker Package
// File    : UserDictionaryUserControl.xaml.cs
// Author  : Eric Woodruff  (Eric@EWoodruff.us)
// Updated : 02/11/2015
// Note    : Copyright 2014-2015, Eric Woodruff, All rights reserved
// Compiler: Microsoft Visual C#
//
// This file contains a user control used to edit the language and user dictionary spell checker configuration
// settings.
//
// This code is published under the Microsoft Public License (Ms-PL).  A copy of the license should be
// distributed with the code and can be found at the project website: https://github.com/EWSoftware/VSSpellChecker
// This notice, the author's name, and all copyright notices must remain intact in all applications,
// documentation, and source files.
//
//    Date     Who  Comments
// ==============================================================================================================
// 06/10/2014  EFW  Moved the language and user dictionary settings to a user control
//===============================================================================================================

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

using PackageResources = VisualStudio.SpellChecker.Properties.Resources;

using VisualStudio.SpellChecker.Configuration;

namespace VisualStudio.SpellChecker.Editors.Pages
{
    /// <summary>
    /// This user control is used to edit the default language and user dictionary spell checker configuration
    /// settings.
    /// </summary>
    public partial class UserDictionaryUserControl : UserControl, ISpellCheckerConfiguration
    {
        #region Constructor
        //=====================================================================

        public UserDictionaryUserControl()
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
            get { return "Default Language/User Dictionary"; }
        }

        /// <inheritdoc />
        public string HelpUrl
        {
            get { return "Default-Language-and-User-Dictionary"; }
        }

        /// <inheritdoc />
        public void LoadConfiguration(SpellingConfigurationFile configuration)
        {
            List<SpellCheckerDictionary> availableDictionaries = new List<SpellCheckerDictionary>();

            cboDefaultLanguage.Items.Clear();

            if(configuration.ConfigurationType != ConfigurationType.Global)
                availableDictionaries.Add(new SpellCheckerDictionary(CultureInfo.InvariantCulture, null, null, null));

            foreach(var lang in SpellCheckerDictionary.AvailableDictionaries(null).Values.OrderBy(d => d.ToString()))
                availableDictionaries.Add(lang);

            cboDefaultLanguage.ItemsSource = availableDictionaries;

            if(configuration.HasProperty(PropertyNames.DefaultLanguage) ||
              configuration.ConfigurationType == ConfigurationType.Global)
            {
                var defaultLang = configuration.ToCultureInfo(PropertyNames.DefaultLanguage);
                var match = cboDefaultLanguage.Items.OfType<SpellCheckerDictionary>().FirstOrDefault(
                    d => d.Culture.Name == defaultLang.Name);

                if(match != null)
                    cboDefaultLanguage.SelectedItem = match;
                else
                    cboDefaultLanguage.SelectedIndex = 0;
            }
            else
                cboDefaultLanguage.SelectedIndex = 0;
        }

        /// <inheritdoc />
        public void SaveConfiguration(SpellingConfigurationFile configuration)
        {
            if(cboDefaultLanguage.SelectedIndex == 0 && configuration.ConfigurationType != ConfigurationType.Global)
                configuration.StoreProperty(PropertyNames.DefaultLanguage, null);
            else
                configuration.StoreProperty(PropertyNames.DefaultLanguage,
                    ((SpellCheckerDictionary)cboDefaultLanguage.SelectedItem).Culture.Name);
        }

        /// <inheritdoc />
        public event EventHandler ConfigurationChanged;

        #endregion

        #region Event handlers
        //=====================================================================

        /// <summary>
        /// Load the user dictionary file when the selected language changes
        /// </summary>
        /// <param name="sender">The sender of the event</param>
        /// <param name="e">The event arguments</param>
        private void cboDefaultLanguage_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            lbUserDictionary.Items.Clear();
            grpUserDictionary.IsEnabled = false;

            if(cboDefaultLanguage.Items.Count != 0 && cboDefaultLanguage.SelectedItem.ToString() != "Inherited")
            {
                string filename = Path.Combine(SpellingConfigurationFile.GlobalConfigurationFilePath,
                    ((SpellCheckerDictionary)cboDefaultLanguage.SelectedItem).Culture.Name + "_User.dic");

                grpUserDictionary.IsEnabled = true;

                if(File.Exists(filename))
                    try
                    {
                        foreach(string word in File.ReadAllLines(filename))
                            lbUserDictionary.Items.Add(word);
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show("Unable to load user dictionary.  Reason: " + ex.Message,
                            PackageResources.PackageTitle, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                    finally
                    {
                        var sd = new SortDescription { Direction = ListSortDirection.Ascending };
                        lbUserDictionary.Items.SortDescriptions.Add(sd);
                    }

                Property_Changed(sender, e);
            }
        }

        /// <summary>
        /// Remove the selected word from the user dictionary
        /// </summary>
        /// <param name="sender">The sender of the event</param>
        /// <param name="e">The event arguments</param>
        private void btnRemoveDictionaryWord_Click(object sender, RoutedEventArgs e)
        {
            int idx = lbUserDictionary.SelectedIndex;
            string word = null;

            if(idx != -1)
            {
                word = (string)lbUserDictionary.Items[idx];
                lbUserDictionary.Items.RemoveAt(idx);
            }

            if(lbUserDictionary.Items.Count != 0)
            {
                if(idx < 0)
                    idx = 0;
                else
                    if(idx >= lbUserDictionary.Items.Count)
                        idx = lbUserDictionary.Items.Count - 1;

                lbUserDictionary.SelectedIndex = idx;
            }

            try
            {
                var selectedDictionary = (SpellCheckerDictionary)cboDefaultLanguage.SelectedItem;

                File.WriteAllLines(selectedDictionary.UserDictionaryFilePath,
                    lbUserDictionary.Items.OfType<string>());

                if(!String.IsNullOrWhiteSpace(word))
                    GlobalDictionary.RemoveWord(selectedDictionary.Culture, word);

                GlobalDictionary.LoadUserDictionaryFile(selectedDictionary.Culture);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Unable to save user dictionary.  Reason: " + ex.Message,
                    PackageResources.PackageTitle, MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        /// <summary>
        /// Import words from a text file
        /// </summary>
        /// <param name="sender">The sender of the event</param>
        /// <param name="e">The event arguments</param>
        private void btnImportDictionary_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dlg.Filter = "Dictionary Files (*.dic)|*.dic|Text documents (.txt)|*.txt|All Files (*.*)|*.*";
            dlg.CheckPathExists = dlg.CheckFileExists = true;

            if((dlg.ShowDialog() ?? false))
            {
                try
                {
                    // Parse words based on the common word break characters and add unique instances to the
                    // user dictionary if not already there excluding those containing digits and those less than
                    // three characters in length.
                    var uniqueWords = File.ReadAllText(dlg.FileName).Split(new[] { ',', '/', '<', '>', '?', ';',
                        ':', '\"', '[', ']', '\\', '{', '}', '|', '-', '=', '+', '~', '!', '#', '$', '%', '^',
                        '&', '*', '(', ')', ' ', '_', '.', '\'', '@', '\t', '\r', '\n' },
                        StringSplitOptions.RemoveEmptyEntries)
                            .Except(lbUserDictionary.Items.OfType<string>())
                            .Distinct()
                            .Where(w => w.Length > 2 && w.IndexOfAny(
                                new[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' }) == -1).ToList();

                    try
                    {
                        var selectedDictionary = (SpellCheckerDictionary)cboDefaultLanguage.SelectedItem;

                        File.WriteAllLines(selectedDictionary.UserDictionaryFilePath, uniqueWords);

                        GlobalDictionary.LoadUserDictionaryFile(selectedDictionary.Culture);

                        cboDefaultLanguage_SelectionChanged(sender, new SelectionChangedEventArgs(
                            e.RoutedEvent, new object[] { }, new object[] { }));
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show("Unable to save user dictionary.  Reason: " + ex.Message,
                            PackageResources.PackageTitle, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show(String.Format(CultureInfo.CurrentCulture, "Unable to load user dictionary " +
                        "from '{0}'.  Reason: {1}", dlg.FileName, ex.Message), PackageResources.PackageTitle,
                        MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
        }

        /// <summary>
        /// Export words to a text file
        /// </summary>
        /// <param name="sender">The sender of the event</param>
        /// <param name="e">The event arguments</param>
        private void btnExportDictionary_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();

            dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dlg.FileName = "UserDictionary.dic";
            dlg.DefaultExt = ".dic";
            dlg.Filter = "Dictionary Files (*.dic)|*.dic|Text documents (.txt)|*.txt|All Files (*.*)|*.*";

            if((dlg.ShowDialog() ?? false))
            {
                try
                {
                    File.WriteAllLines(dlg.FileName, lbUserDictionary.Items.OfType<string>());
                }
                catch(Exception ex)
                {
                    MessageBox.Show(String.Format(CultureInfo.CurrentCulture, "Unable to save user dictionary " +
                        "to '{0}'.  Reason: {1}", dlg.FileName, ex.Message), PackageResources.PackageTitle,
                        MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
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
