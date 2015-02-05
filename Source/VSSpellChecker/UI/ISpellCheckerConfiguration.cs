//===============================================================================================================
// System  : Visual Studio Spell Checker Package
// File    : ISpellCheckerConfiguration.cs
// Author  : Eric Woodruff  (Eric@EWoodruff.us)
// Updated : 02/02/2015
// Note    : Copyright 2014-2015, Eric Woodruff, All rights reserved
// Compiler: Microsoft Visual C#
//
// This file contains an interface used by the configuration user controls
//
// This code is published under the Microsoft Public License (Ms-PL).  A copy of the license should be
// distributed with the code and can be found at the project website: https://github.com/EWSoftware/VSSpellChecker
// This notice, the author's name, and all copyright notices must remain intact in all applications,
// documentation, and source files.
//
//    Date     Who  Comments
// ==============================================================================================================
// 06/09/2013  EFW  Created the code
//===============================================================================================================

using System.Windows.Controls;

using VisualStudio.SpellChecker.Configuration;

namespace VisualStudio.SpellChecker.UI
{
    /// <summary>
    /// This interface is implemented by the spell checker configuration user controls
    /// </summary>
    public interface ISpellCheckerConfiguration
    {
        /// <summary>
        /// This read-only property returns a reference to the user control containing the settings
        /// </summary>
        UserControl Control { get; }

        /// <summary>
        /// This read-only property returns the category title
        /// </summary>
        string Title { get; }

        /// <summary>
        /// This read-only property returns the help URL
        /// </summary>
        string HelpUrl { get; }

        /// <summary>
        /// This read-only property is used to validate the configuration settings in the user control
        /// </summary>
        /// <returns>True if valid, false if not</returns>
        bool IsValid { get; }

        /// <summary>
        /// Load the configuration settings for the control
        /// </summary>
        /// <param name="configuration">The configuration file from which to load settings</param>
        void LoadConfiguration(SpellingConfigurationFile configuration);

        /// <summary>
        /// Save the configuration settings for the control to the given configuration file
        /// </summary>
        /// <returns>True if saved successfully, false if not</returns>
        bool SaveConfiguration(SpellingConfigurationFile configuration);
    }
}
