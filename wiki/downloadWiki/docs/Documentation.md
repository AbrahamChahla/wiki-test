# Overview
The Office Integration Pack for Microsoft LightSwitch is a LightSwitch Extension that developers can use to add Microsoft Office integration features to their LightSwitch projects. Using the Office Integration Pack developers can import/export from Microsoft Excel, import into Microsoft Word, associate a Word template with a LightSwitch application and export data to content controls in the document, send email, and create appointments. This document describes the methods included in the Office Integration Pack.

# The Office Integration Pack Sample Application
A sample application is available at [http://officeintegration.codeplex.com/releases/87595/download/378876](http://officeintegration.codeplex.com/releases/87595/download/378876) that demonstrates how to use the Office Integration Pack. The comments in code are extensive and should give you a good feel for how most of the functionality works. There are many screens in the sample that demonstrate a variety of common uses of Excel, Word, and Outlook.

The prerequisities for this sample are:

* Visual Studio LightSwitch 2011 
* Office Integration Pack Extension 
* Microsoft Office 2010 (Excel, Word, and Outlook) 

To use the sample application, extract the contents of the ZIP and copy all the Word and Excel documents to your “My Documents” folder. Then double-click on the BookStoreApp.SLN file to open the sample in Visual Studio LightSwitch. Press F5 to compile and run.

# Excel Integration Features
The methods in the Office Integration Pack are grouped by the Microsoft Office application to which they apply. The features related to Microsoft Excel are found in the OfficeIntegration.Excel namespace. 

## Export
Used to export a collection of items to a range in an Excel workbook. There are 12 overloads:

_**Export(Collection as IVisualCollection) as Object**_
This overload simply exports the Collection of items to Sheet1!A1 in a new workbook.
* Collection is the IVisualCollection of items that you want to export to the workbook.
* Example:
{{ Export(Me.Applicants) 
}}
_**Export(Collection as IVisualCollection, ColumnNames As List(Of String)) as Object**_
This overload simply exports the Collection of items to Sheet1!A1 in a new workbook with the specified columns.
* Collection is the IVisualCollection of items that you want to export to the workbook.
* ColumnNames contains a list of strings specifying the names of the columns to export.
* Example:
{{ 
Dim names as New List(Of String)
names.Add("Name")
names.Add("Address")
names.Add("City")
Export(Me.Applicants, names) 
}}
_**Export(Collection as IVisualCollection, ColumnNames As List(Of ColumnMapping)) as Object**_
This overload simply exports the Collection of items to Sheet1!A1 in a new workbook with the specified column mapping.
* Collection is the IVisualCollection of items that you want to export to the workbook.
* ColumnNames contains a list of ColumnMapping objects specifying the names of the columns to export.
* Example:
{{ 
Dim mappings as New List(Of OfficeIntegration.ColumnMapping)
mappings.Add(New OfficeIntegration.ColumnMapping("","Name"))
mappings.Add(New OfficeIntegration.ColumnMapping("","Address"))
mappings.Add(New OfficeIntegration.ColumnMapping("","City"))
Export(Me.Applicants, mappings) 
}}
**_ Export(Collection as IVisualCollection, Workbook as String, Worksheet as String, Range as String) as Object_**
* Collection is the IVisualCollection of items that you want to export to the workbook.
* Workbook represents the full path of the workbook that you want to export to.
* Worksheet represents the name of the worksheet that you want to export to.
* Range represents either the range name or address of the cell that will be the top-left cell of the range. If the range name or address contains more than 1 cell, the first cell in the range is used.
* Examples: 
{{Export(Me.Applicants, _
     "C:\Users\SHansen\Applicants.xlsx", _
     "Sheet2", _
     "C5")
Export(Me.Applicants, _
     "C:\Users\SHansen\Applicants.xlsx", _ 
     "Sheet2", _
     "DestinationRange")
}}
**_ Export(Collection as IVisualCollection, Workbook as String, Worksheet as String, Range as String, ColumnNames as List(Of String)) as Object_**
* Collection is the IVisualCollection of items that you want to export to the workbook.
* Workbook represents the full path of the workbook that you want to export to.
* Worksheet represents the name of the worksheet that you want to export to.
* Range represents either the range name or address of the cell that will be the top-left cell of the range. If the range name or address contains more than 1 cell, the first cell in the range is used.
* ColumnNames is a list containing the names of the fields that should be exported.
* Example:
{{Dim names as New List(Of String)
names.Add("Name")
names.Add("Address")
names.Add("City")
Export(Me.Applicants, _
     "C:\Users\SHansen\Applicants.xlsx", _
     "Sheet2", _
     "C5", _
     names)
}}
**_ Export(Collection as IVisualCollection, Workbook as String, Range as String, ColumnNames as List(Of OfficeIntegration.ColumnMapping)) as Object_**
* Collection is the IVisualCollection of items that you want to export to the workbook.
* Workbook represents the full path of the workbook that you want to export to.
* Worksheet represents the name of the worksheet that you want to export to.
* Range represents either the range name or address of the cell that will be the top-left cell of the range. If the range name or address contains more than 1 cell, the first cell in the range is used.
* ColumnNames is a list of ColumnMapping objects containing the names of the fields that should be exported.
* Example:
{{Dim mappings as New List(Of OfficeIntegration.ColumnMapping)
mappings.Add(New OfficeIntegration.ColumnMapping("","Name"))
mappings.Add(New OfficeIntegration.ColumnMapping("","Address"))
mappings.Add(New OfficeIntegration.ColumnMapping("","City"))
Export(Me.Applicants, _
     "C:\Users\SHansen\Applicants.xlsx", _
     "Sheet2", _
     "C5", _
     mappings)
}}
_**Export(Collection as IEnumerable) as Object**_
This overload simply exports the Collection of items to Sheet1!A1 in a new workbook.
* Collection is a generic collection of items that you want to export to the workbook.
* Example:
{{ 
Dim results = From a In Me.Applicants
     where a.Name.ToLower.Contains("s")
     Select a.Name, a.Address, a.City
Export(results) 
}}
_**Export(Collection as IEnumerable, ColumnNames As List(Of String)) as Object**_
This overload simply exports the Collection of items to Sheet1!A1 in a new workbook with the specified columns.
* Collection is a generic collection of items that you want to export to the workbook.
* ColumnNames contains a list of strings specifying the names of the columns to export.
* Example:
{{ 
Dim results = From a In Me.Applicants
     where a.Name.ToLower.Contains("s")
Dim names as New List(Of String)
names.Add("Name")
names.Add("Address")
names.Add("City")
Export(results, names) 
}}
_**Export(Collection as IEnumerable, ColumnNames As List(Of ColumnMapping)) as Object**_
This overload simply exports the Collection of items to Sheet1!A1 in a new workbook with the specified column mappings.
* Collection is a generic collection of items that you want to export to the workbook.
* ColumnNames contains a list of ColumnMapping objects specifying the names of the columns to export.
* Example:
{{ 
Dim results = From a In Me.Applicants
     where a.Name.ToLower.Contains("s")
Dim mappings as New List(Of OfficeIntegration.ColumnMapping)
mappings.Add(New OfficeIntegration.ColumnMapping("","Name"))
mappings.Add(New OfficeIntegration.ColumnMapping("","Address"))
mappings.Add(New OfficeIntegration.ColumnMapping("","City"))
Export(results, mappings) 
}}
**_ Export(Collection as IEnumerable, Workbook as String, Worksheet as String, Range as String) as Object_**
* Collection is a generic collection of items that you want to export to the workbook.
* Workbook represents the full path of the workbook that you want to export to.
* Worksheet represents the name of the worksheet that you want to export to.
* Range represents either the range name or address of the cell that will be the top-left cell of the range. If the range name or address contains more than 1 cell, the first cell in the range is used.
* Examples: 
{{
Dim results = From a In Me.Applicants
     where a.Name.ToLower.Contains("s")
Export(results, _
     "C:\Users\SHansen\Applicants.xlsx", _
     "Sheet2", _
     "C5")
Export(results, _
     "C:\Users\SHansen\Applicants.xlsx", _ 
     "Sheet2", _
     "DestinationRange")
}}
**_ Export(Collection as IEnumerable, Workbook as String, Worksheet as String, Range as String, ColumnNames as List(Of String)) as Object_**
* Collection is a generic collection of items that you want to export to the workbook.
* Workbook represents the full path of the workbook that you want to export to.
* Worksheet represents the name of the worksheet that you want to export to.
* Range represents either the range name or address of the cell that will be the top-left cell of the range. If the range name or address contains more than 1 cell, the first cell in the range is used.
* ColumnNames is a list containing the names of the fields that should be exported.
* Example:
{{
Dim results = From a In Me.Applicants
     where a.Name.Contains("s")
Dim names as New List(Of String)
names.Add("Name")
names.Add("Address")
names.Add("City")
Export(results, _
     "C:\Users\SHansen\Applicants.xlsx", _
     "Sheet2", _
     "C5", _
     names)
}}
**_ Export(Collection as IEnumerable, Workbook as String, Range as String, ColumnNames as List(Of OfficeIntegration.ColumnMapping)) as Object_**
* Collection is a generic collection of items that you want to export to the workbook.
* Workbook represents the full path of the workbook that you want to export to.
* Worksheet represents the name of the worksheet that you want to export to.
* Range represents either the range name or address of the cell that will be the top-left cell of the range. If the range name or address contains more than 1 cell, the first cell in the range is used.
* ColumnNames is a list of ColumnMapping objects containing the names of the fields that should be exported.
* Example:
{{
Dim results = From a In Me.Applicants
     where a.Name.Contains("s")
Dim mappings as New List(Of OfficeIntegration.ColumnMapping)
mappings.Add(New OfficeIntegration.ColumnMapping("","Name"))
mappings.Add(New OfficeIntegration.ColumnMapping("","Address"))
mappings.Add(New OfficeIntegration.ColumnMapping("","City"))
Export(results, _
     "C:\Users\SHansen\Applicants.xlsx", _
     "Sheet2", _
     "C5", _
     mappings)
}} 
## ExportEntityCollection
_ These methods have been deprecated and may be removed upon future releases. Please use Word.Export() overloads. _

ExportEntityCollection performs the same way as the Export method except that ExportEntityCollection accepts an IEntityCollection object rather than an IVisualCollection. As with the Export method, there are four overloads:
**_ ExportEntityCollection(Collection as IEntityCollection) as Boolean _**
This overload simply exports the Collection of items to Sheet1!A1 in a new workbook.
* Collection is the IEntityCollection of items that you want to export to the workbook.
* Example:
{{ ExportEntityCollection(Me.Authors.SelectedItem.Books)
}}
**_ ExportEntityCollection(Collection as IEntityCollection, Workbook as String, Worksheet as String, Range as String) as Boolean _**
* Collection is the IEntityCollection of items that you want to export to the workbook.
* Workbook represents the full path of the workbook that you want to export to.
* Worksheet represents the name of the worksheet that you want to export to.
* Range represents either the range name or address of the cell that will be the top-left cell of the range. If the range name or address contains more than 1 cell, the first cell in the range is used.
* Examples: 
{{ExportEntityCollection(Me.Applicants, _
     "C:\Users\SHansen\Applicants.xlsx", _
     "Sheet2", _
     "C5")
ExportEntityCollection(Me.Applicants, _
     "C:\Users\SHansen\Applicants.xlsx", _ 
     "Sheet2", _
     "DestinationRange")
}}
**_ ExportEntityCollection(Collection as IEntityCollection, Workbook as String, Worksheet as String, Range as String, ColumnNames as List(Of String)) as Boolean _**
* Collection is the IEntityCollection of items that you want to export to the workbook.
* Workbook represents the full path of the workbook that you want to export to.
* Worksheet represents the name of the worksheet that you want to export to.
* Range represents either the range name or address of the cell that will be the top-left cell of the range. If the range name or address contains more than 1 cell, the first cell in the range is used.
* ColumnNames is a list containing the names of the fields that should be exported.
* Example:
{{Dim names as New List(Of String)
names.Add("Name")
names.Add("Address")
names.Add("City")
ExportEntityCollection(Me.Applicants, _
     "C:\Users\SHansen\Applicants.xlsx", _
     "Sheet2", _
     "C5", _
     names)
}}
**_ ExportEntityCollection(Collection as IEntityCollection, Workbook as String, Range as String, ColumnNames as List(Of OfficeIntegration.ColumnMapping)) as Boolean _**
* Collection is the IEntityCollection of items that you want to export to the workbook.
* Workbook represents the full path of the workbook that you want to export to.
* Worksheet represents the name of the worksheet that you want to export to.
* Range represents either the range name or address of the cell that will be the top-left cell of the range. If the range name or address contains more than 1 cell, the first cell in the range is used.
* ColumnNames is a list of ColumnMapping objects containing the names of the fields that should be exported.
* Example:
{{Dim mappings as New List(Of OfficeIntegration.ColumnMapping)
mappings.Add(New OfficeIntegration.ColumnMapping("","Name"))
mappings.Add(New OfficeIntegration.ColumnMapping("","Address"))
mappings.Add(New OfficeIntegration.ColumnMapping("","City"))
ExportEntityCollection(Me.Applicants, _
     "C:\Users\SHansen\Applicants.xlsx", _
     "Sheet2", _
     "C5", _
     mappings)
}}
## GetExcel 
GetExcel returns an instance of Microsoft Excel.
**_ GetExcel() as Object _**
* The returned Object is of type Excel.Application
* See example under the GetWorkbook method. 
## GetWorkbook
GetWorkbook returns an Excel workbook.
**_ GetWorkbook(Excel as Object, WorkbookPath as String) as Object _**
* The returned Object is of type Excel.Workbook
* Example:
{{Dim xl as Object
xl = OfficeIntegration.Excel.GetExcel()
If Not xl is Nothing Then
	Dim wb as Object
	wb = GetWorkbook(xl, "c:\users\shansen\example.doc")
	If Not wb is Nothing Then
		‘ do something with the workbook here
	End If
	xl.Application.Visible = True
End If
}}
## Import
Imports a range from an Excel workbook into LightSwitch. There are three overloads:
**_ Import(Collection as IVisualCollection) _**
* Imports a range of data into the Collection
* Collection is the IVisualCollection that the data should be imported into.
* UI is displayed to allow the end-user to choose which workbook to import
* UI is displayed to allow the end-user to perform column mapping
* Assumes the data is located on the first worksheet in the workbook starting at cell A1.
* Example:
{{Import(Me.Applicants)
}}
**_ Import(Collection as IVisualCollection, ColumnMappings As List(Of ColumnMapping)) _**
* Imports a range of data into the Collection
* Collection is the IVisualCollection that the data should be imported into.
* ColumnMappings specifies the columns to import and the respective columns they should be mapped to in LightSwitch.
* UI is displayed to allow the end-user to choose which workbook to import
* Assumes the data is located on the first worksheet in the workbook starting at cell A1.
* Example:
{{
Dim mappings As New List(Of OfficeIntegration.ColumnMapping)
mappings.Add(New OfficeIntegration.ColumnMapping("Name", "Name"))
mappings.Add(New OfficeIntegration.ColumnMapping("Address", "Address"))
mappings.Add(New OfficeIntegration.ColumnMapping("City", "City"))
Import(Me.Applicants,mappings)
}}
**_ Import(Collection as IVisualCollection, Workbook as String, Worksheet as String, Range as String) _**
* Imports a range of data into the Collection
* Collection is the IVisualCollection that the data should be imported into.
* Workbook is the full path to the workbook containing the data to import
* Worksheet is the name of the worksheet containing the data
* Range is the address or named range representing the top-left cell in the data range. If the range consists of multiple cells, the first cell is assumed to be the top-left cell in the range.
* UI is displayed to allow the end-user to perform column mapping
* Example:
{{Import(Me.Applicants, _
     "c:\Users\SHansen\applicants.xlsx", _
     "Sheet1", _
     "B2") 
Import(Me.Applicants, _
     "c:\Users\SHansen\applicants.xlsx", _
     "Sheet3", _
     "DataRange")
}}
**_ Import(Collection as IVisualCollection, Workbook as String, Worksheet as String, Range as String, ColumnMapping as List(Of OfficeIntegration.ColumnMapping)) _**
* Imports a range of data into the Collection
* Collection is the IVisualCollection that the data should be imported into.
* Workbook is the full path to the workbook containing the data to import
* Worksheet is the name of the worksheet containing the data
* Range is the address or named range representing the top-left cell in the data range. If the range consists of multiple cells, the first cell is assumed to be the top-left cell in the range.
* ColumnMapping is a list of ColumnMapping objects containing the names of the fields that should be imported.
* Example:
{{Dim mappings as New List(Of OfficeIntegration.ColumnMapping)
mappings.Add( _
     New OfficeIntegration.ColumnMapping( _
     "Applicant Name", _
     "Name"))
mappings.Add( _
     New OfficeIntegration.ColumnMapping( _
     "Address", _
     "Address"))
mappings.Add( _
     New OfficeIntegration.ColumnMapping( _
     "City", _
     "City"))
Import(Me.Applicants, _
     "c:\Users\SHansen\applicants.xlsx", _
     "Sheet1", _
     "B2", _
     mappings)
}}
# Outlook Integration Features
The methods in the Office Integration Pack are grouped by the Microsoft Office application to which they apply. The features related to Microsoft Outlook are found in the OfficeIntegration.Outlook namespace. 
## CreateAppointment
Create a new Outlook appointment item.
**_ CreateAppointment(Address as String, Subject as String, Body as String, Location as String, StartDateTime as Date, EndDateTime as Date) as Object _**
## CreateEmail
Create a new Outlook email item. There are two overloads:
* CreateEmail(Address as String, Subject as String, Body as String) as Object
* CreateEmail(Address as String, Subject as String, Items as IVisualCollection) as Object
## HtmlExport
Creates an HTML table containing items from an IVisualCollection:
* HtmlExport(Items as IVisualCollection) as String
* HtmlExport(Items as IVisualCollection, ColumnNames as List(Of String)) as String
## HtmlExportEntityCollection
Creates an HTML table containing items from an IEntityCollection:
* HtmlExportEntityCollection(Items as IEntityCollection) as String
* HtmlExportEntityCollection(Items as IEntityCollection, ColumnNames as List(Of String)) as String
# Word Integration Features
The methods in the Office Integration Pack are grouped by the Microsoft Office application to which they apply. The features related to Microsoft Word are found in the OfficeIntegration.Word namespace. 
## Export
The Export method exports a collection of items to a table in a Word document. Each Export overload now supports exporting of images. There are 15 overloads:
**_ Export(Collection as IVisualCollection, UseActiveDocument as Boolean) as Object _**
* Use this overload to export data to a table in a new Word document.
* Returns an Object of type Word.Document
* Collection is the collection of items to export.
* If UseActiveDocument=True, then a table is added at the selection (cursor) of whatever document is currently active. If UseActiveDocument=False, then the table is added to a new document.
* Example:
{{Export(Me.Applicants, True)
}}
**_ Export(Collection as IVisualCollection, BuildColumnHeadings As Boolean, ColumnNames As List(Of ColumnMapping), UseActiveDocument as Boolean) as Object _**
* Use this overload to export data to a table in a new Word document.
* Returns an Object of type Word.Document
* Collection is the collection of items to export.
* If BuildColumnHeadings=True, then the header row of the table will be populated with data based on the name of the columns in ColumnNames
* ColumnNames is a list of ColumnMapping objects that specify the columns to export
* If UseActiveDocument=True, then a table is added at the selection (cursor) of whatever document is currently active. If UseActiveDocument=False, then the table is added to a new document.
* Example:
{{
Dim mappings As New List(Of OfficeIntegration.ColumnMapping)
mappings.Add(New OfficeIntegration.ColumnMapping("", "Name"))
mappings.Add(New OfficeIntegration.ColumnMapping("", "Address"))
mappings.Add(New OfficeIntegration.ColumnMapping("", "City"))
Export(Me.Applicants, True, mappings, False)
}}
**_ Export(Document as Object, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IVisualCollection) as Object _**
* Use this overload to export data to an existing table in a Word document.
* Returns an Object of type Word.Document
* Document is an Object of type Word.Document
* BookmarkName is the name of a bookmark in the document that indicates which table to export to.
* StartRow is the row number that the first item will be exported to. This is useful if you do not want this method to build column headings because the table already contains pre-labeled column headings. You can use a StartRow of 2  with BuildColumnHeadings = false to preserve the existing column headings in the table and start exporting data only to row 2.
* BuildColumnHeadings is a flag indicating whether or not the method should output column heading to the table.
* Collection is the collection of items to export.
* Example:
{{Dim word as Object
word = OfficeIntegration.Word.GetWord()
If Not word is Nothing Then
	Dim doc as Object
	doc = GetDocument(word, "c:\users\shansen\example.doc")
	If Not doc is Nothing Then
		Export(doc, "ApplicantTable", 1, True, Me.Applicants)
	End If
	Word.Application.Visible = True
End If
}}
**_ Export(Document as Object, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IVisualCollection, ColumnNames as List(Of OfficeIntegration.ColumnMapping)) as Object _**
* Same as above except you can limit which fields are exported using the ColumnNames parameter
**_ Export(Document as Object, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IVisualCollection, ColumnNames as List(Of String)) as Object _**
* Same as above except you can limit which fields are exported using the ColumnNames parameter
**_ Export(DocumentPath as String, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IVisualCollection) as Object _**
* Use this overload to export data to an existing table in a Word document.
* Returns an Object of type Word.Document
* DocumentPath is the full path to the desired document.
* BookmarkName is the name of a bookmark in the document that indicates which table to export to.
* StartRow is the row number that the first item will be exported to. This is useful if you do not want this method to build column headings because the table already contains pre-labeled column headings. You can use a StartRow of 2  with BuildColumnHeadings = false to preserve the existing column headings in the table and start exporting data only to row 2.
* BuildColumnHeadings is a flag indicating whether or not the method should output column heading to the table.
* Collection is the collection of items to export.
* Example:
{{Export("c:\users\shansen\example.doc", _
     "ApplicantTable", _
     1, _
     True, _
     Me.Applicants)
}}
**_ Export(DocumentPath as String, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IVisualCollection, ColumnNames as List(Of OfficeIntegration.ColumnMapping)) as Object _**
* Same as above except you can limit which fields are exported using the ColumnNames parameter
**_ Export(DocumentPath as String, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IVisualCollection, ColumnNames as List(Of String)) as Object _**
* Same as above except you can limit which fields are exported using the ColumnNames parameter
**_ Export(Collection as IEnumerable, UseActiveDocument as Boolean) as Object _**
* Use this overload to export a generic collection of data to a table in a new Word document.
* Returns an Object of type Word.Document
* Collection is a generic collection of items to export.
* If UseActiveDocument=True, then a table is added at the selection (cursor) of whatever document is currently active. If UseActiveDocument=False, then the table is added to a new document.
* Example:
{{
Dim results From a In Me.Applicants
     where a.Name.Contains("s")
Export(results, True)
}}
**_ Export(Collection as IEnumerable, BuildColumnHeadings as Boolean, ColumnNames As List(Of ColumnMapping), UseActiveDocument as Boolean) as Object _**
* Same as above except you can specify the columns to export.
* Collection is a generic collection of items to export.
* If BuildColumnHeadings=True, then the header row of the table will be populated with data based on the name of the columns in ColumnNames
* ColumnNames is a list of ColumnMapping objects that specify the columns to export
* If UseActiveDocument=True, then a table is added at the selection (cursor) of whatever document is currently active. If UseActiveDocument=False, then the table is added to a new document.
* Example:
{{
Dim results From a In Me.Applicants
     where a.Name.Contains("s")
Dim mappings As New List(Of OfficeIntegration.ColumnMapping)
mappings.Add(New OfficeIntegration.ColumnMapping("", "Name"))
mappings.Add(New OfficeIntegration.ColumnMapping("", "Address"))
mappings.Add(New OfficeIntegration.ColumnMapping("", "City"))
Export(results, True, mappings, False)
}}
**_ Export(Collection as IEnumerable, BuildColumnHeadings as Boolean, ColumnNames As List(Of String), UseActiveDocument as Boolean) as Object _**
* Same as above except you can specify the columns to export.
* Collection is a generic collection of items to export.
* If BuildColumnHeadings=True, then the header row of the table will be populated with data based on the name of the columns in ColumnNames
* ColumnNames contains a list of strings specifying the columns to export.
* If UseActiveDocument=True, then a table is added at the selection (cursor) of whatever document is currently active. If UseActiveDocument=False, then the table is added to a new document.
* Example:
{{
Dim results From a In Me.Applicants
     where a.Name.Contains("s")
Dim names as New List(Of String)
names.Add("Name")
names.Add("Address")
names.Add("City")
Export(results, True, names, True)
}}
**_ Export(Document as Object, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IEnumerable, ColumnNames as List(Of String)) as Object _**
* Use this overload to export a generic collection of data to an existing table in a Word document.
* Returns an Object of type Word.Document
* Document is an Object of type Word.Document
* BookmarkName is the name of a bookmark in the document that indicates which table to export to.
* StartRow is the row number that the first item will be exported to. This is useful if you do not want this method to build column headings because the table already contains pre-labeled column headings. You can use a StartRow of 2  with BuildColumnHeadings = false to preserve the existing column headings in the table and start exporting data only to row 2.
* BuildColumnHeadings is a flag indicating whether or not the method should output column heading to the table.
* Collection is a generic collection of items to export.
* ColumnNames is a list of strings specifying the columns to export
* Example:
{{Dim word as Object
word = OfficeIntegration.Word.GetWord()
If Not word is Nothing Then
	Dim doc as Object
	doc = GetDocument(word, "c:\users\shansen\example.doc")
	If Not doc is Nothing Then
                Dim results = From a In Me.Applicants
                     where a.Name.Contains("s")
                Dim names as New List(Of String)
                names.Add("Name")
                names.Add("Address")
                names.Add("City")
		Export(doc, "ApplicantTable", 1, True, results, names)
	End If
	Word.Application.Visible = True
End If
}}
**_ Export(DocumentPath as String, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IEnumerable, ColumnNames as List(Of String)) as Object _**
* Same as above except you can specify the document path.
**_ Export(Document as Object, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IEnumerable, ColumnNames as List(Of ColumnMapping)) as Object _**
* Use this overload to export a generic collection of data to an existing table in a Word document.
* Returns an Object of type Word.Document
* Document is an Object of type Word.Document.
* BookmarkName is the name of a bookmark in the document that indicates which table to export to.
* StartRow is the row number that the first item will be exported to. This is useful if you do not want this method to build column headings because the table already contains pre-labeled column headings. You can use a StartRow of 2  with BuildColumnHeadings = false to preserve the existing column headings in the table and start exporting data only to row 2.
* BuildColumnHeadings is a flag indicating whether or not the method should output column heading to the table.
* Collection is a generic collection of items to export.
* ColumnNames is a list of ColumnMapping objects used to specify the columns to export
* Example:
{{Dim word as Object
word = OfficeIntegration.Word.GetWord()
If Not word is Nothing Then
	Dim doc as Object
        doc = GetDocument(word, "c:\users\shansen\example.doc")
        If Not doc is Nothing Then
	        Dim results = From a In Me.Applicants
                   where a.Name.Contains("s")
                Dim mappings as New List(Of OfficeIntegration.ColumnMapping)
                mappings.Add(New OfficeIntegration.ColumnMapping("","Name"))
                mappings.Add(New OfficeIntegration.ColumnMapping("","Address"))
                mappings.Add(New OfficeIntegration.ColumnMapping("","City"))
	        Export(doc, "ApplicantTable", 1, True, results, mappings)
        End If
	Word.Application.Visible = True
End If
}}
**_ Export(DocumentPath as String, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IEnumerable, ColumnNames as List(Of ColumnMapping)) as Object _**
* Same as above except you can specify the document path.
## ExportEntityCollection 
_ These methods have been deprecated and may be removed upon future releases. Please use Word.Export() overloads. _

Use ExportEntityCollection to export a collection of items to a table in a Word document. There are seven overloads. The overloads are identical to the overloads for the Export method with one difference. ExportEntityCollection accepts an IEntityCollection rather than the IVisualCollection supported by the Export method. 
**_ ExportEntityCollection(Collection as IEntityCollection, UseActiveDocument as Boolean) as Object _**
* Use this overload to export data to a table in a new Word document.
* Returns an Object of type Word.Document
* Collection is the collection of items to export.
* If UseActiveDocument=True, then a table is added at the selection (cursor) of whatever document is currently active. If UseActiveDocument=False, then the table is added to a new document.
* Example:
{{ExportEntityCollection(Me.Authors.SelectedItem.Books, True)
}}
**_ ExportEntityCollection(Document as Object, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IEntityCollection) as Object _**
* Use this overload to export data to an existing table in a Word document.
* Returns an Object of type Word.Document
* Document is an Object of type Word.Document
* BookmarkName is the name of a bookmark in the document that indicates which table to export to.
* StartRow is the row number that the first item will be exported to. This is useful if you do not want this method to build column headings because the table already contains pre-labeled column headings. You can use a StartRow of 2  with BuildColumnHeadings = false to preserve the existing column headings in the table and start exporting data only to row 2.
* BuildColumnHeadings is a flag indicating whether or not the method should output column heading to the table.
* Collection is the collection of items to export.
* Example:
{{Dim word as Object
word = OfficeIntegration.Word.GetWord()
If Not word is Nothing Then
     Dim doc as Object
     doc = GetDocument(word, "c:\users\shansen\example.doc")
     If Not doc is Nothing Then
          ExportEntityCollection(doc, _
               "ApplicantTable", _
               1, _
               True, _
               Me.Applicants)
     End If
     Word.Application.Visible = True
End If
}}
**_ ExportEntityCollection(Document as Object, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IEntityCollection, ColumnNames as List(Of OfficeIntegration.ColumnMapping)) as Object _**
* Same as above except you can limit which fields are exported using the ColumnNames parameter
**_ ExportEntityCollection(Document as Object, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IEntityCollection, ColumnNames as List(Of String)) as Object _**
* Same as above except you can limit which fields are exported using the ColumnNames parameter
**_ ExportEntityCollection(DocumentPath as String, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IEntityCollection) as Object _**
* Use this overload to export data to an existing table in a Word document.
* Returns an Object of type Word.Document
* DocumentPath is the full path to the desired document.
* BookmarkName is the name of a bookmark in the document that indicates which table to export to.
* StartRow is the row number that the first item will be exported to. This is useful if you do not want this method to build column headings because the table already contains pre-labeled column headings. You can use a StartRow of 2  with BuildColumnHeadings = false to preserve the existing column headings in the table and start exporting data only to row 2.
* BuildColumnHeadings is a flag indicating whether or not the method should output column heading to the table.
* Collection is the collection of items to export.
* Example:
{{ExportEntityCollection("c:\users\shansen\example.doc", _
     "ApplicantTable", _
     1, _
     True, _
     Me.Applicants)
}}
**_ ExportEntityCollection(DocumentPath as String, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IEntityCollection, ColumnNames as List(Of OfficeIntegration.ColumnMapping)) as Object _**
* Same as above except you can limit which fields are exported using the ColumnNames parameter
**_ ExportEntityCollection(DocumentPath as String, BookmarkName as String, StartRow as Integer, BuildColumnHeadings as Boolean, Collection as IVisualCollection, ColumnNames as List(Of String)) as Object _**
* Same as above except you can limit which fields are exported using the ColumnNames parameter
## GenerateDocument
GenerateDocument allows you to associate a Word template containing content controls with a LightSwitch project and then map entity fields to the content controls.
**_ GenerateDocument(Template as String, Item as IEntityObject, ColumnMappings as List(Of OfficeIntegration.ColumnMapping)) as Object _**
* Returns an Object of type Word.Document
* Template is a string representing the full path to the document which will be used as a template
* Item is the IEntityObject whose data will be used to populate content controls in the template
* ColumnMappings is a list of column mappings that indicate how to map fields in the IEntityObject to content controls in the template. Content Controls in the document are located by title.
* Example:
{{Dim mappings as New List(Of OfficeIntegration.ColumnMapping)
mappings.Add( _
     New OfficeIntegration.ColumnMapping("Applicant","Name"))
mappings.Add( _
     New OfficeIntegration.ColumnMapping("Address","Address"))
mappings.Add( _
     New OfficeIntegration.ColumnMapping("City","City"))
Dim doc as Object
Dim sPath as String
sPath = "c:\users\shansen\documents\letter1.docx"
doc = GenerateDocument(sPath, _
       Me.Applicants.SelectedItem, _
       mappings)
}}
## GetDocument  
You can use GetDocument to open a Word document and return a Word Document object reference to the associated document.
**_ GetDocument(Word as Object, DocumentPath as String) as Object _**
* Returns an Object of type Word.Document
## GetWord 
GetWord is a method to conveniently obtain a reference to Microsoft Word
**_ GetWord() as Object _**
* Returns an Object of type Word.Application
## SaveAsPdf
SaveAsPdf saves a Word document as PDF, optionally showing the PDF document after it is created
**_ SaveAsPDF(Document as Object, FullName as String, ShowPDF as Boolean) _**
* Document is an Object of type Word.Document
* FullName represents the file path where the PDF document will be saved
* ShowPDF is a flag indicating whether to display the PDF document after it is created or not.
* Example (building on the GenerateDocument example from above)
{{Dim mappings as New List(Of OfficeIntegration.ColumnMapping)
mappings.Add( _
     New OfficeIntegration.ColumnMapping("Applicant","Name"))
mappings.Add( _
     New OfficeIntegration.ColumnMapping("Address","Address"))
mappings.Add( _
     New OfficeIntegration.ColumnMapping("City","City"))
Dim doc as Object
Dim sPath as String
sPath = "c:\users\shansen\documents\letter1.docx"
doc = GenerateDocument(sPath, _
  Me.Applicants.SelectedItem, _
  mappings)
Dim sPDFPath as String
sPDFPath = "c:\users\shansen\documents\OfferLetter.pdf"
SaveAsPDF(doc, sPDFPath, True)
}}
# SMTP Integration Features
This class is only applicable to code running on the server tier. To use the methods in the OfficeIntegration.Smtp namespace, call them from data oriented event handlers within the ApplicationDataService class.
## CreateAppointment 
Create and send an appointment using the iCal format.
**_ CreateAppointment(SendFrom as String, SendTo as String, Subject as String, Body as String, Location as String, StartTime as Date, EndTime as Date, MsgID as String, Sequence as Integer, IsCancelled as Boolean, Server as SmtpServer) as Boolean _**
## CreateEmail 
Create and send an email. There are two overloads:
**_ CreateEmail(SendFrom as String, SendTo as String, Subject as String, Body as String, Server as SmtpServer) as Boolean _**
**_ CreateEmail(SendFrom as String, SendTo as String, Subject as String, Body as IEnumerable, Server as SmtpServer) as Boolean _**
# OfficeIntegration.SmtpServer
SmtpServer is a class representing the connection details associated with an SMTP server. This class is required for the methods exposed in OfficeIntegration.Smtp. This class contains the following properties:
**_ SmtpServer – The name of the SMTP server. _**
**_ SmtpPort – The port used by the SMTP server. _**
**_ SmtpUserId – The name of the account to use when logging on to the SMTP server. _**
**_ SmtpPassword – The password to use when authenticating to the SMTP server. _**
