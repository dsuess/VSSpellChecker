//===============================================================================================================
// System  : Visual Studio Spell Checker Package
// File    : SpellingServiceFactory.cs
// Authors : Noah Richards, Roman Golovin, Michael Lehenbauer, Eric Woodruff
// Updated : 02/04/2015
// Note    : Copyright 2010-2015, Microsoft Corporation, All rights reserved
//           Portions Copyright 2013-2015, Eric Woodruff, All rights reserved
// Compiler: Microsoft Visual C#
//
// This file contains a class that implements the spelling dictionary service factory
//
// This code is published under the Microsoft Public License (Ms-PL).  A copy of the license should be
// distributed with the code and can be found at the project website: https://github.com/EWSoftware/VSSpellChecker
// This notice, the author's name, and all copyright notices must remain intact in all applications,
// documentation, and source files.
//
//    Date     Who  Comments
// ==============================================================================================================
// 04/14/2013  EFW  Imported the code into the project
// 04/30/2013  EFW  Moved the global dictionary creation into the GlobalDictionary class
//===============================================================================================================

using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.TextManager.Interop;

using VisualStudio.SpellChecker.Configuration;

namespace VisualStudio.SpellChecker
{
    /// <summary>
    /// This class implements the spelling dictionary service factory
    /// </summary>
    [Export]
    internal sealed class SpellingServiceFactory
    {
        #region Private data members
        //=====================================================================

        private const string SpellCheckerDisabledKey = "@@VisualStudio.SpellChecker.Disabled";

        #endregion

        #region Factory methods
        //=====================================================================

        /// <summary>
        /// Get the configuration settings for the specified buffer
        /// </summary>
        /// <param name="buffer">The buffer for which to get the configuration settings</param>
        /// <returns>The spell checker configuration settings for the buffer or null if one is not provided or
        /// is disabled for the given buffer.</returns>
        public SpellCheckerConfiguration GetConfiguration(ITextBuffer buffer)
        {
            SpellCheckerConfiguration config = null;
            bool isDisabled = false;

            // If not given a buffer or already checked for and found to be disabled, don't go any further
            if(buffer != null && !buffer.Properties.TryGetProperty(SpellCheckerDisabledKey, out isDisabled) &&
              !buffer.Properties.TryGetProperty(typeof(SpellCheckerConfiguration), out config))
            {
                // TODO: Generate the configuration settings for the file.  For now, all we have are the global
                // settings.
                config = SpellCheckerConfiguration.GlobalConfiguration;

                if(!config.SpellCheckAsYouType || config.IsExcludedByExtension(buffer.GetFilenameExtension()))
                {
                    // Mark it as disabled so that we don't have to check again
                    buffer.Properties[SpellCheckerDisabledKey] = true;
                    config = null;
                }
                else
                    buffer.Properties[typeof(SpellCheckerConfiguration)] = config;
            }

            return config;
        }

        /// <summary>
        /// Get the dictionary for the specified buffer
        /// </summary>
        /// <param name="buffer">The buffer for which to get a dictionary</param>
        /// <returns>The spelling dictionary for the buffer or null if one is not provided</returns>
        public SpellingDictionary GetDictionary(ITextBuffer buffer)
        {
            SpellingDictionary service = null;

            if(buffer != null)
            {
                // Get the configuration and create the dictionary based on the configuration
                var config = this.GetConfiguration(buffer);

                if(config != null && !buffer.Properties.TryGetProperty(typeof(SpellingDictionary), out service))
                {
                    // Create or get the existing global dictionary for the default language
                    var globalDictionary = GlobalDictionary.CreateGlobalDictionary(config.DefaultLanguage);

                    if(globalDictionary != null)
                    {
                        service = new SpellingDictionary(globalDictionary);
                        buffer.Properties[typeof(SpellingDictionary)] = service;
                    }
                }
            }

            return service;
        }
        #endregion

        #region Helper methods
        //=====================================================================

        /// <summary>
        /// Get the filename for the given text buffer
        /// </summary>
        /// <param name="buffer">The text buffer for which to obtain the filename</param>
        /// <returns>The filename if obtained, null if not</returns>
        private static string GetFileNameFromTextBuffer(ITextBuffer buffer)
        {
            IVsTextBuffer adapter;
            string filename = null;
            uint formatIndex;

            if(buffer.Properties.TryGetProperty(typeof(IVsTextBuffer), out adapter))
            {
                IPersistFileFormat pff = adapter as IPersistFileFormat;

                if(pff != null)
                    try
                    {
                        if(pff.GetCurFile(out filename, out formatIndex) != VSConstants.S_OK)
                            filename = null;
                    }
                    catch
                    {
                        // Ignore exceptions, we just won't return a filename
                    }
            }

            return filename;
        }
        #endregion
    }
}
