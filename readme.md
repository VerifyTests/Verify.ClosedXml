# <img src="/src/icon.png" height="30px"> Verify.ClosedXml

[![Discussions](https://img.shields.io/badge/Verify-Discussions-yellow?svg=true&label=)](https://github.com/orgs/VerifyTests/discussions)
[![Build status](https://github.com/VerifyTests/Verify.ClosedXml/actions/workflows/build.yml/badge.svg)](https://github.com/VerifyTests/Verify.ClosedXml/actions/workflows/build.yml)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.ClosedXml.svg)](https://www.nuget.org/packages/Verify.ClosedXml/)

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of Excel documents via [ClosedXML](https://github.com/ClosedXML/ClosedXML).<!-- singleLineInclude: intro. path: /docs/intro.include.md -->

Converts Excel documents (xlsx) to csv for verification.

**See [Milestones](../../milestones?state=closed) for release notes.**


## Sponsors


### Entity Framework Extensions<!-- include: sponsors. path: /docs/sponsors.include.md -->

[Entity Framework Extensions](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.ClosedXml) is a major sponsor and is proud to contribute to the development this project.

[![Entity Framework Extensions](https://raw.githubusercontent.com/VerifyTests/Verify.ClosedXml/refs/heads/main/docs/zzz.png)](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.ClosedXml)

### Developed using JetBrains IDEs

[![JetBrains logo.](https://raw.githubusercontent.com/VerifyTests/Verify.ClosedXml/main/docs/jetbrains.png)](https://jb.gg/OpenSourceSupport)<!-- endInclude -->


## NuGet

 * https://nuget.org/packages/Verify.ClosedXml


## Usage


### Enable Verify.ClosedXml

<!-- snippet: enable -->
<a id='snippet-enable'></a>
```cs
[ModuleInitializer]
public static void Initialize() =>
    VerifyClosedXml.Initialize();
```
<sup><a href='/src/Tests/ModuleInitializer.cs#L3-L9' title='Snippet source file'>snippet source</a> | <a href='#snippet-enable' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Choosing what is verified

A workbook is verified as the xlsx, an info file, and a csv of each sheet. What is left out is controlled by Verify's own settings, which every Verify plugin that splits a document shares: see [Leaving out what was derived](https://github.com/VerifyTests/Verify/blob/main/docs/paged-documents.md#leaving-out-what-was-derived) and [Excluding targets](https://github.com/VerifyTests/Verify/blob/main/docs/converter.md#excluding-targets). Anything left out is not produced at all (the csv is not built, the workbook is not saved), so these also save work.

`ExcludeDerivedTargets("csv")` leaves out the csv of each sheet, keeping the workbook and its info file:

<!-- snippet: ExcludeCsv -->
<a id='snippet-ExcludeCsv'></a>
```cs
[Test]
public Task ExcludeCsv() =>
    VerifyFile("sample.xlsx")
        .ExcludeDerivedTargets("csv");
```
<sup><a href='/src/Tests/Samples.cs#L87-L94' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcludeCsv' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`ExcludeTargets("xlsx")` leaves out the workbook, keeping the info file and the csv of each sheet:

<!-- snippet: ExcludeXlsx -->
<a id='snippet-ExcludeXlsx'></a>
```cs
[Test]
public Task ExcludeXlsx() =>
    VerifyFile("sample.xlsx")
        .ExcludeTargets("xlsx");
```
<sup><a href='/src/Tests/Samples.cs#L96-L103' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcludeXlsx' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Each can also be set for every test, on `VerifierSettings`:

<!-- snippet: InitializeOutputs -->
<a id='snippet-InitializeOutputs'></a>
```cs
[ModuleInitializer]
public static void Initialize()
{
    VerifyClosedXml.Initialize();

    // For every test: no csv, so only the workbook and its info are verified
    VerifierSettings.ExcludeDerivedTargets("csv");
}
```
<sup><a href='/src/StaticSettingsTests/ModuleInitializer.cs#L3-L14' title='Snippet source file'>snippet source</a> | <a href='#snippet-InitializeOutputs' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

A workbook has no pages, so the settings for the pages of a document, `PageText` and `PagesToInclude`, have no effect on it.


### Input 

For a given input Excel file.

<img src="/src/Tests/sample_Sheet1.png">


### Verify a file

<!-- snippet: VerifyExcel -->
<a id='snippet-VerifyExcel'></a>
```cs
[Test]
public Task VerifyExcel() =>
    VerifyFile("sample.xlsx");
```
<sup><a href='/src/Tests/Samples.cs#L29-L35' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Snapshot Result

For a given Verify, the result is 3 (or more files)


#### Metadata

<!-- snippet: Samples.VerifyExcel.DotNet9_0.verified.txt -->
<a id='snippet-Samples.VerifyExcel.DotNet9_0.verified.txt'></a>
```txt
{
  SheetNames: [
    Sheet1
  ],
  Properties: {
    Title: The Title
  },
  WorksheetCount: 1,
  DefaultFont: Arial,
  CalculateMode: Default,
  Style: {
    Font: {
      Name: Arial
    }
  }
}
```
<sup><a href='/src/Tests/Samples.VerifyExcel.DotNet9_0.verified.txt#L1-L16' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyExcel.DotNet9_0.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### CSV

One per sheet, named for the sheet: `#Sheet1.verified.csv`. A workbook with one sheet is no exception, so that a second sheet adds a file rather than renaming the first. A hidden sheet has a csv as any other, and is named under `HiddenSheets` in the info file. A sheet is a page, numbered in tab order with hidden sheets counted, so `PagesToInclude` leaves out the csv of a sheet. The info file still names every sheet, and the xlsx is still the whole workbook.

<!-- snippet: Samples.VerifyExcel.DotNet9_0#Sheet1.verified.csv -->
<a id='snippet-Samples.VerifyExcel.DotNet9_0#Sheet1.verified.csv'></a>
```csv
0,First Name,Last Name,Gender,Country,Date,Age,Id,Formula
1,Dulce,Abril,Female,United States,DateTime_1,32,1562,1594 (G2+H2)
2,Mara,Hashimoto,Female,Great Britain,DateTime_2,25,1582,1607 (G3+H3)
3,Philip,Gent,Male,France,DateTime_3,36,2587,2623 (G4+H4)
4,Kathleen,Hanner,Female,United States,DateTime_1,25,3549,3574 (G5+H5)
5,Nereida,Magwood,Female,United States,DateTime_2,58,2468,2526 (G6+H6)
6,Gaston,Brumm,Male,United States,DateTime_3,24,2554,2578 (G7+H7)
```
<sup><a href='/src/Tests/Samples.VerifyExcel.DotNet9_0%23Sheet1.verified.csv#L1-L7' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyExcel.DotNet9_0#Sheet1.verified.csv' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Characters of a sheet name that a file name cannot hold are replaced with `-`: a sheet named `Q2 <draft>` is `#Q2 -draft-.verified.csv`. Sheets that are then named the same are told apart by an index: `Q1 <draft>` and `Q1 |draft|` are `#Q1 -draft-.00.verified.csv` and `#Q1 -draft-.01.verified.csv`.


#### Excel file

<img src="/src/Tests/Samples.VerifyExcel.DotNet9_0_Sheet1.png">


### Verify a Stream

<!-- snippet: VerifyExcelStream -->
<a id='snippet-VerifyExcelStream'></a>
```cs
[Test]
public Task VerifyExcelStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
    return Verify(stream, "xlsx");
}
```
<sup><a href='/src/Tests/Samples.cs#L76-L85' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcelStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Verify a ClosedXML SpreadsheetDocument

<!-- snippet: XLWorkbook -->
<a id='snippet-XLWorkbook'></a>
```cs
[Test]
public Task XLWorkbook()
{
    using var book = new XLWorkbook();

    var sheet = book.Worksheets.Add("Basic Data");

    sheet.Cell("A1").Value = "ID";
    sheet.Cell("B1").Value = "Name";

    sheet.Cell("A2").Value = 1;
    sheet.Cell("B2").Value = "John Doe";

    sheet.Cell("A3").Value = 2;
    sheet.Cell("B3").Value = "Jane Smith";

    return Verify(book);
}
```
<sup><a href='/src/Tests/Samples.cs#L45-L66' title='Snippet source file'>snippet source</a> | <a href='#snippet-XLWorkbook' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Verify a named target

An xlsx can be a named target of a verification, the way an attachment of a mail message is:

<!-- snippet: NamedTarget -->
<a id='snippet-NamedTarget'></a>
```cs
[Test]
public Task NamedTarget()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
    return Verify(new Target("xlsx", stream, "Attachment1"));
}
```
<sup><a href='/src/Tests/Samples.cs#L105-L114' title='Snippet source file'>snippet source</a> | <a href='#snippet-NamedTarget' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Verify names the files relative to the target that was converted, so they take its name: `#Attachment1.verified.xlsx`, `#Attachment1.verified.txt` and `#Attachment1.Sheet1.verified.csv`.


### Binary output across .NET frameworks

When verifying binary package output (xlsx, docx, nupkg, etc.) across multiple target frameworks (e.g. net48 and net10.0), the binary output may differ due to Deflate compression implementation differences. The XML content within entries is identical — only the compressed bytes differ. Use `UniqueForRuntime` to generate framework-specific verified files:

```cs
await Verify(stream, extension: "xlsx")
    .UniqueForRuntime();
```

See [Verify Naming docs](https://github.com/VerifyTests/Verify/blob/main/docs/naming.md) for more details.


## Reviewing changes

A change to a workbook is a change to several files: the xlsx, its info file, and the csv of every sheet. Verify tells the diff tool that the csv files and the info file were derived from the xlsx, and [DiffEngineViewer](https://github.com/VerifyTests/DiffEngine/blob/main/docs/viewer.md#files-derived-from-a-document), which reads and draws an xlsx itself, shows them as one row and accepts them together. Other diff tools are given each file, as before.


## Migrating from 1.x

Version 2 moves to the [source and derived targets](https://github.com/VerifyTests/Verify/blob/main/docs/converter.md#source-and-derived-targets) of Verify 33.3. The `outputs` parameter of `Initialize` is gone, and `ClosedXmlOutputs` is obsolete as an error, so that code naming it is pointed here. What they chose is chosen with Verify's settings, which can be set for one verification as well as for every test:

| 1.x | 2.x |
| --- | --- |
| `Initialize(ClosedXmlOutputs.None)` | `Initialize()` and `VerifierSettings.ExcludeDerivedTargets("csv")` |
| `Initialize(ClosedXmlOutputs.Csv)` | `Initialize()` |
| `Initialize(ClosedXmlOutputs.All)` | `Initialize()` |

Verify's `ExcludeTargets("xlsx")` already left out the workbook. With it, the workbook is now not saved at all.

The snapshots of a workbook that is verified directly, as a file, a stream or an `XLWorkbook`, keep their names and their content. Nothing has to be accepted again.

An xlsx that is a [named target](#verify-a-named-target) is the exception. The plugin built the name of each csv from the name of the target, and left the workbook and its info file with no name. All three are now named by Verify, relative to the target. For a test `Tests.Mail` whose target is an xlsx named `Attachment1`:

| 1.x | 2.x |
| --- | --- |
| `Tests.Mail.verified.xlsx` | `Tests.Mail#Attachment1.verified.xlsx` |
| `Tests.Mail.verified.txt` | `Tests.Mail#Attachment1.verified.txt` |
| `Tests.Mail#Attachment1-Sheet1.verified.csv` | `Tests.Mail#Attachment1.Sheet1.verified.csv` |

Renamed snapshots show as a new file and a pending delete. Accepting both, or running once with [AutoVerify](https://github.com/VerifyTests/Verify/blob/main/docs/autoverify.md), moves a test over, and since the content of the files is unchanged, source control shows each as a rename.
