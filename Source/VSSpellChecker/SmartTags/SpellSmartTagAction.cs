//===============================================================================================================
// System  : Visual Studio Spell Checker Package
// File    : SpellSmartTagAction.cs
// Authors : Noah Richards, Roman Golovin, Michael Lehenbauer, Eric Woodruff
// Updated : 10/28/2015
// Note    : Copyright 2010-2015, Microsoft Corporation, All rights reserved
//           Portions Copyright 2013-2015, Eric Woodruff, All rights reserved
// Compiler: Microsoft Visual C#
//
// This file contains a class used to provide a smart tag action for inserting spelling suggestions
//
// This code is published under the Microsoft Public License (Ms-PL).  A copy of the license should be
// distributed with the code and can be found at the project website: https://github.com/EWSoftware/VSSpellChecker
// This notice, the author's name, and all copyright notices must remain intact in all applications,
// documentation, and source files.
//
//    Date     Who  Comments
// ==============================================================================================================
// 04/14/2013  EFW  Imported the code into the project
// 05/02/2013  EFW  Added support for Replace All
// 07/28/2015  EFW  Added support for culture information in the spelling suggestions
//===============================================================================================================

using System.Collections.ObjectModel;
using System.Windows.Input;

using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;

using VisualStudio.SpellChecker.Definitions;

namespace VisualStudio.SpellChecker.SmartTags
{
    /// <summary>
    /// Smart tag action for inserting spelling suggestions
    /// </summary>
    internal class SpellSmartTagAction : ISmartTagAction
    {
        #region Private data members
        //=====================================================================

        private ITrackingSpan span;
        private ISpellingSuggestion replaceWith;
        private SpellingDictionary dictionary;

        #endregion

        #region Constructor
        //=====================================================================

        /// <summary>
        /// Constructor for spelling suggestions smart tag actions
        /// </summary>
        /// <param name="span">The word span to replace</param>
        /// <param name="replaceWith">The suggestion to replace misspelled word with</param>
        /// <param name="dictionary">The dictionary used to perform the Replace All action</param>
        public SpellSmartTagAction(ITrackingSpan span, ISpellingSuggestion replaceWith, SpellingDictionary dictionary)
        {
            this.span = span;
            this.replaceWith = replaceWith;
            this.dictionary = dictionary;
        }
        #endregion

        #region ISmartTagAction members
        //=====================================================================

        /// <summary>
        /// Display text
        /// </summary>
        public virtual string DisplayText
        {
            get { return replaceWith.Suggestion.Replace("_", "__"); }
        }

        /// <summary>
        /// Icon to place next to the display text
        /// </summary>
        public System.Windows.Media.ImageSource Icon
        {
            get { return null; }
        }

        /// <summary>
        /// This method is executed when action is selected in the context menu
        /// </summary>
        public void Invoke()
        {
            if(dictionary != null && Keyboard.Modifiers == ModifierKeys.Control)
                dictionary.ReplaceAllOccurrences(span.GetText(span.TextBuffer.CurrentSnapshot), replaceWith);
            else
                span.TextBuffer.Replace(span.GetSpan(span.TextBuffer.CurrentSnapshot), replaceWith.Suggestion);
        }

        /// <summary>
        /// Always enabled unless a span is not specified
        /// </summary>
        public bool IsEnabled
        {
            get { return (span != null); }
        }

        /// <summary>
        /// This smart tag has no action sets
        /// </summary>
        public ReadOnlyCollection<SmartTagActionSet> ActionSets
        {
            get { return null; }
        }
        #endregion
    }
}
