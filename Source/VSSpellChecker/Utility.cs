//===============================================================================================================
// System  : Sandcastle Help File Builder Visual Studio Package
// File    : Utility.cs
// Author  : Eric Woodruff  (Eric@EWoodruff.us)
// Updated : 02/09/2015
// Note    : Copyright 2013-2015, Eric Woodruff, All rights reserved
// Compiler: Microsoft Visual C#
//
// This file contains a utility class with extension and utility methods.
//
// This code is published under the Microsoft Public License (Ms-PL).  A copy of the license should be
// distributed with the code and can be found at the project website: https://github.com/EWSoftware/VSSpellChecker
// This notice, the author's name, and all copyright notices must remain intact in all applications,
// documentation, and source files.
//
//    Date     Who  Comments
// ==============================================================================================================
// 05/25/2013  EFW  Created the code
//===============================================================================================================

using System;
using System.Globalization;
using System.IO;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;

using VisualStudio.SpellChecker.Configuration;
using VisualStudio.SpellChecker.Editors;
using VisualStudio.SpellChecker.Properties;

namespace VisualStudio.SpellChecker
{
    /// <summary>
    /// This class contains utility and extension methods
    /// </summary>
    public static class Utility
    {
        #region Constants
        //=====================================================================

        /// <summary>
        /// This is an extra <c>PredefinedTextViewRoles</c> value that only exists in VS 2013.  It defines the
        /// Peek Definition window view role which doesn't exist in earlier versions of Visual Studio.  As such,
        /// we define it here.
        /// </summary>
        public const string EmbeddedPeekTextView = "EMBEDDED_PEEK_TEXT_VIEW";
        #endregion

        #region General utility methods
        //=====================================================================

        /// <summary>
        /// Get a service from the Sandcastle Help File Builder package
        /// </summary>
        /// <param name="throwOnError">True to throw an exception if the service cannot be obtained,
        /// false to return null.</param>
        /// <typeparam name="TInterface">The interface to obtain</typeparam>
        /// <typeparam name="TService">The service used to get the interface</typeparam>
        /// <returns>The service or null if it could not be obtained</returns>
        public static TInterface GetServiceFromPackage<TInterface, TService>(bool throwOnError)
            where TInterface : class
            where TService : class
        {
            IServiceProvider provider = VSSpellCheckerPackage.Instance;

            TInterface service = (provider == null) ? null : provider.GetService(typeof(TService)) as TInterface;

            if(service == null && throwOnError)
                throw new InvalidOperationException("Unable to obtain service of type " + typeof(TService).Name);

            return service;
        }

        /// <summary>
        /// This displays a formatted message using the <see cref="IVsUIShell"/> service
        /// </summary>
        /// <param name="icon">The icon to show in the message box</param>
        /// <param name="message">The message format string</param>
        /// <param name="parameters">An optional list of parameters for the message format string</param>
        public static void ShowMessageBox(OLEMSGICON icon, string message, params object[] parameters)
        {
            Guid clsid = Guid.Empty;
            int result;

            if(message == null)
                throw new ArgumentNullException("message");

            IVsUIShell uiShell = GetServiceFromPackage<IVsUIShell, SVsUIShell>(true);

            ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(0, ref clsid,
                Resources.PackageTitle, String.Format(CultureInfo.CurrentCulture, message, parameters),
                String.Empty, 0, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST, icon, 0,
                out result));
        }

        /// <summary>
        /// Get the filename from the given text buffer
        /// </summary>
        /// <param name="buffer">The text buffer from which to get the filename</param>
        /// <returns>The filename or null if it could not be obtained</returns>
        public static string GetFilename(this ITextBuffer buffer)
        {
            ITextDocument textDoc;

            if(buffer != null && buffer.Properties.TryGetProperty(typeof(ITextDocument), out textDoc))
                if(textDoc != null && !String.IsNullOrEmpty(textDoc.FilePath))
                    return textDoc.FilePath;

            return null;
        }

        /// <summary>
        /// Get the filename extension from the given text buffer
        /// </summary>
        /// <param name="buffer">The text buffer from which to get the filename extension</param>
        /// <returns>The filename extension or null if it could not be obtained</returns>
        public static string GetFilenameExtension(this ITextBuffer buffer)
        {
            ITextDocument textDoc;
            string extension = null;

            if(buffer != null && buffer.Properties.TryGetProperty(typeof(ITextDocument), out textDoc))
                if(textDoc != null && !String.IsNullOrEmpty(textDoc.FilePath))
                    extension = Path.GetExtension(textDoc.FilePath);

            return extension;
        }

        /// <summary>
        /// This returns the given absolute file path relative to the given base path
        /// </summary>
        /// <param name="absolutePath">The file path to convert to a relative path</param>
        /// <param name="basePath">The base path to which the absolute path is made relative</param>
        /// <returns>The file path relative to the given base path</returns>
        public static string ToRelativePath(this string absolutePath, string basePath)
        {
            bool hasBackslash = false;
            string relPath;
            int minLength, idx;

            // If not specified, use the current folder as the base path
            if(basePath == null || basePath.Trim().Length == 0)
                basePath = Directory.GetCurrentDirectory();
            else
                basePath = Path.GetFullPath(basePath);

            if(absolutePath == null)
                absolutePath = String.Empty;

            // Just in case, make sure the path is absolute
            if(!Path.IsPathRooted(absolutePath))
                if(!absolutePath.Contains("*") && !absolutePath.Contains("?"))
                    absolutePath = Path.GetFullPath(absolutePath);
                else
                    absolutePath = Path.Combine(Path.GetFullPath(Path.GetDirectoryName(absolutePath)),
                        Path.GetFileName(absolutePath));

            if(absolutePath.Length > 1 && absolutePath[absolutePath.Length - 1] == '\\')
            {
                absolutePath = absolutePath.Substring(0, absolutePath.Length - 1);
                hasBackslash = true;
            }

            // Split the paths into their component parts
            char[] separators = { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar,
                Path.VolumeSeparatorChar };
            string[] baseParts = basePath.Split(separators);
            string[] absParts = absolutePath.Split(separators);

            // Find the common base path
            minLength = Math.Min(baseParts.Length, absParts.Length);

            for(idx = 0; idx < minLength; idx++)
                if(String.Compare(baseParts[idx], absParts[idx], StringComparison.OrdinalIgnoreCase) != 0)
                    break;

            // Use the absolute path if there's nothing in common (i.e. they are on different drives or network
            // shares.
            if(idx == 0)
                relPath = absolutePath;
            else
            {
                // If equal to the base path, it doesn't have to go anywhere.  Otherwise, work up from the base
                // path to the common root.
                if(idx == baseParts.Length)
                    relPath = String.Empty;
                else
                    relPath = new String(' ', baseParts.Length - idx).Replace(" ", ".." +
                        Path.DirectorySeparatorChar);

                // And finally, add the path from the common root to the absolute path
                relPath += String.Join(Path.DirectorySeparatorChar.ToString(), absParts, idx,
                    absParts.Length - idx);
            }

            return (hasBackslash) ? relPath + "\\" : relPath;
        }
        #endregion

        #region Property state conversion methods
        //=====================================================================

        /// <summary>
        /// Convert the named property value to the appropriate selection state
        /// </summary>
        /// <param name="configuration">The configuration file from which to obtain the property value</param>
        /// <param name="propertyName">The name of the property to get</param>
        /// <returns>The selection state based on the specified property's value</returns>
        public static PropertyState ToPropertyState(this SpellingConfigurationFile configuration,
          string propertyName)
        {
            return !configuration.HasProperty(propertyName) &&
                configuration.ConfigurationType != ConfigurationType.Global ? PropertyState.Inherited :
                configuration.ToBoolean(propertyName) ? PropertyState.Yes : PropertyState.No;
        }

        /// <summary>
        /// Convert the selection state value to a property value
        /// </summary>
        /// <param name="state">The selection state to convert</param>
        /// <returns>The appropriate property value to store</returns>
        public static bool? ToPropertyValue(this PropertyState state)
        {
            return (state == PropertyState.Inherited) ? (bool?)null : state == PropertyState.Yes ? true : false;
        }

        #endregion
    }
}
