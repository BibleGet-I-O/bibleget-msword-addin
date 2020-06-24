﻿Imports System.Resources
Imports System.Reflection
Imports System.Runtime.InteropServices
<Assembly: NeutralResourcesLanguageAttribute("")>
<Assembly: CLSCompliant(False)>

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitle("BibleGet IO for Microsoft Word")>
<Assembly: AssemblyDescription("A tool for inserting Bible quotes into your documents.")>
<Assembly: AssemblyCompany("Cappellania Università degli Studi Roma Tre")>
<Assembly: AssemblyProduct("BibleGet IO for Microsoft Word")>
<Assembly: AssemblyCopyright("Copyright © 2015 BibleGet IO | E-Mail admin@bibleget.io")>
<Assembly: AssemblyTrademark("BibleGet IO")>

' Setting ComVisible to false makes the types in this assembly not visible 
' to COM components.  If you need to access a type in this assembly from 
' COM, set the ComVisible attribute to true on that type.
<Assembly: ComVisible(False)>

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("5288dc29-ba64-4411-88a6-02aeb7379d8e")>

' Version information for an assembly consists of the following four values:
'
'      Major Version
'      Minor Version 
'      Build Number
'      Revision
'
' You can specify all the values or you can default the Build and Revision Numbers 
' by using the '*' as shown below:
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("2.2.6.2")>
<Assembly: AssemblyFileVersion("2.2.6.2")>

Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module
