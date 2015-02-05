//===============================================================================================================
// System  : Visual Studio Spell Checker Package
// File    : AssemblyInfo.cs
// Author  : Eric Woodruff  (Eric@EWoodruff.us)
// Updated : 02/23/2014
// Note    : Copyright 2013-2014, Eric Woodruff, All rights reserved
// Compiler: Microsoft Visual C#
//
// Visual Studio spell checker definition attributes.
//
// This code is published under the Microsoft Public License (Ms-PL).  A copy of the license should be
// distributed with the code and can be found at the project website: https://github.com/EWSoftware/VSSpellChecker
// This notice, the author's name, and all copyright notices must remain intact in all applications,
// documentation, and source files.
//
//    Date     Who  Comments
// ==============================================================================================================
// 05/20/2013  EFW  Created the code
//===============================================================================================================

using System;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;

// General assembly information
[assembly: AssemblyProduct("Visual Studio Spell Checker")]
[assembly: AssemblyCompany("Eric Woodruff")]
[assembly: AssemblyCopyright(AssemblyInfo.Copyright)]
[assembly: AssemblyCulture("")]
#if DEBUG
[assembly: AssemblyConfiguration("Debug")]
#else
[assembly: AssemblyConfiguration("Release")]
#endif

[assembly: AssemblyTitle("Visual Studio Spell Checker Package")]
[assembly: AssemblyDescription("This assembly contains a package that implement the Visual Studio spell checker")]

// This assembly is not CLS compliant
[assembly: CLSCompliant(false)]

// Not visible to COM
[assembly: ComVisible(false)]

// Resources contained within the assembly are English
[assembly: NeutralResourcesLanguageAttribute("en")]

// Version numbers.  See comments below.
[assembly: AssemblyVersion(AssemblyInfo.StrongNameVersion)]
[assembly: AssemblyFileVersion(AssemblyInfo.FileVersion)]
[assembly: AssemblyInformationalVersion(AssemblyInfo.ProductVersion)]

// This defines constants that can be used throughout the package.
//
// All version numbers for an assembly consists of the following four values:
//
//      Year of release
//      Month of release
//      Day of release
//      Revision (typically zero unless multiple releases are made on the same day)
//
// This versioning scheme allows build component and plug-in developers to use the same major, minor, and build
// numbers as the Sandcastle tools to indicate with which version their components are compatible.
//
internal static partial class AssemblyInfo
{
    // Configuration file schema version - DO NOT CHANGE UNLESS NECESSARY.
    //
    // This is used to set the version in the spell checker configuration files.  This should remain unchanged to
    // maintain compatibility with prior releases.  It should only be changed if a breaking change is made that
    // requires the configuration file to be upgraded to a newer format.
    public const string ConfigSchemaVersion = "2015.2.1.0";

    // Common assembly strong name version
    //
    // This is used to set the assembly version in the strong name.  Typically, this should remain unchanged to
    // maintain binary compatibility with prior releases.  However, since nothing has a dependency on this
    // package it can be kept in synch with the values below.
    public const string StrongNameVersion = "2015.2.1.0";

    // Common assembly file version
    //
    // This is used to set the assembly file version.  This will change with each new release.  MSIs only
    // support a Major value between 0 and 255 so we drop the century from the year on this one.
    public const string FileVersion = "15.2.1.0";

    // Common product version
    //
    // This may contain additional text to indicate Alpha or Beta states.  The version number will always match
    // the file version above but includes the century on the year.
    public const string ProductVersion = "2015.2.1.0";

    // Assembly copyright information
    public const string Copyright = "Copyright \xA9 2013-2015, Eric Woodruff, All Rights Reserved.\r\n" +
        "Portions Copyright \xA9 2010-2015, Microsoft Corporation, All Rights Reserved.";
}
