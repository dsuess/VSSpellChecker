//===============================================================================================================
// System  : Visual Studio Spell Checker Package
// File    : CSharpOptionsUserControl.xaml.cs
// Author  : Eric Woodruff  (Eric@EWoodruff.us)
// Updated : 02/02/2015
// Note    : Copyright 2014-2015, Eric Woodruff, All rights reserved
// Compiler: Microsoft Visual C#
//
// This file contains a user control used to edit the C# spell checker configuration settings
//
// This code is published under the Microsoft Public License (Ms-PL).  A copy of the license should be
// distributed with the code and can be found at the project website: https://github.com/EWSoftware/VSSpellChecker
// This notice, the author's name, and all copyright notices must remain intact in all applications,
// documentation, and source files.
//
//    Date     Who  Comments
// ==============================================================================================================
// 06/12/2014  EFW  Created the code
//===============================================================================================================

using System.Windows.Controls;

using VisualStudio.SpellChecker.Configuration;

namespace VisualStudio.SpellChecker.UI
{
    /// <summary>
    /// This user control is used to edit the C# spell checker configuration settings
    /// </summary>
    public partial class CSharpOptionsUserControl : UserControl, ISpellCheckerConfiguration
    {
        #region Constructor
        //=====================================================================

        /// <summary>
        /// Constructor
        /// </summary>
        public CSharpOptionsUserControl()
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
            get { return "C# Options"; }
        }

        /// <inheritdoc />
        public string HelpUrl
        {
            get { return "C#-Options"; }
        }

        /// <inheritdoc />
        public bool IsValid
        {
            get { return true; }
        }

        /// <inheritdoc />
        public void LoadConfiguration(SpellingConfigurationFile configuration)
        {
            chkIgnoreXmlDocComments.IsChecked = configuration.ToBoolean(
                PropertyNames.CSharpOptionsIgnoreXmlDocComments);
            chkIgnoreDelimitedComments.IsChecked = configuration.ToBoolean(
                PropertyNames.CSharpOptionsIgnoreDelimitedComments);
            chkIgnoreStandardSingleLineComments.IsChecked = configuration.ToBoolean(
                PropertyNames.CSharpOptionsIgnoreStandardSingleLineComments);
            chkIgnoreQuadrupleSlashComments.IsChecked = configuration.ToBoolean(
                PropertyNames.CSharpOptionsIgnoreQuadrupleSlashComments);
            chkIgnoreNormalStrings.IsChecked = configuration.ToBoolean(
                PropertyNames.CSharpOptionsIgnoreNormalStrings);
            chkIgnoreVerbatimStrings.IsChecked = configuration.ToBoolean(
                PropertyNames.CSharpOptionsIgnoreVerbatimStrings);
        }

        /// <inheritdoc />
        public bool SaveConfiguration(SpellingConfigurationFile configuration)
        {
            configuration.StoreProperty(PropertyNames.CSharpOptionsIgnoreXmlDocComments,
                chkIgnoreXmlDocComments.IsChecked.Value);
            configuration.StoreProperty(PropertyNames.CSharpOptionsIgnoreDelimitedComments,
                chkIgnoreDelimitedComments.IsChecked.Value);
            configuration.StoreProperty(PropertyNames.CSharpOptionsIgnoreStandardSingleLineComments,
                chkIgnoreStandardSingleLineComments.IsChecked.Value);
            configuration.StoreProperty(PropertyNames.CSharpOptionsIgnoreQuadrupleSlashComments,
                chkIgnoreQuadrupleSlashComments.IsChecked.Value);
            configuration.StoreProperty(PropertyNames.CSharpOptionsIgnoreNormalStrings,
                chkIgnoreNormalStrings.IsChecked.Value);
            configuration.StoreProperty(PropertyNames.CSharpOptionsIgnoreVerbatimStrings,
                chkIgnoreVerbatimStrings.IsChecked.Value);

            return true;
        }
        #endregion
    }
}
