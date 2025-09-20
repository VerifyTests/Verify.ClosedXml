# <img src="/src/icon.png" height="30px"> Verify.ClosedXml

[![Discussions](https://img.shields.io/badge/Verify-Discussions-yellow?svg=true&label=)](https://github.com/orgs/VerifyTests/discussions)
[![Build status](https://ci.appveyor.com/api/projects/status/xyn3eaf6i5tc9l5e?svg=true)](https://ci.appveyor.com/project/SimonCropp/verify-closedxml)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.ClosedXml.svg)](https://www.nuget.org/packages/Verify.ClosedXml/)

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of Excel documents via [ClosedXML](https://github.com/ClosedXML/ClosedXML).<!-- singleLineInclude: intro. path: /docs/intro.include.md -->

Converts Excel documents (xlsx) to csv for verification.

**See [Milestones](../../milestones?state=closed) for release notes.**


## Sponsors


### Entity Framework Extensions<!-- include: zzz. path: /docs/zzz.include.md -->

[Entity Framework Extensions](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.ClosedXml) is a major sponsor and is proud to contribute to the development this project.

[![Entity Framework Extensions](https://raw.githubusercontent.com/VerifyTests/Verify.ClosedXml/refs/heads/main/docs/zzz.png)](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.ClosedXml)<!-- endInclude -->


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


### Verify a file

<!-- snippet: VerifyExcel -->
<a id='snippet-VerifyExcel'></a>
```cs
[Test]
public Task VerifyExcel() =>
    VerifyFile("sample.xlsx");
```
<sup><a href='/src/Tests/Samples.cs#L30-L36' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


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
<sup><a href='/src/Tests/Samples.cs#L73-L82' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcelStream' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L42-L63' title='Snippet source file'>snippet source</a> | <a href='#snippet-XLWorkbook' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Example snapshot

For a given Verify, the result is 3 (or more files)


#### Metadata

snippet: Samples.VerifyExcel.DotNet9_0.verified.txt


#### CSV

One per sheet

<!-- snippet: Samples.VerifyExcel.DotNet9_0.verified.csv -->
<a id='snippet-Samples.VerifyExcelStream.DotNet9_0.verified.csv'></a>
```csv
0,First Name,Last Name,Gender,Country,Date,Age,Id,Formula
1,Dulce,Abril,Female,United States,DateTime_1,32,1562,1594 (G2+H2)
2,Mara,Hashimoto,Female,Great Britain,DateTime_2,25,1582,1607 (G3+H3)
3,Philip,Gent,Male,France,DateTime_3,36,2587,2623 (G4+H4)
4,Kathleen,Hanner,Female,United States,DateTime_1,25,3549,3574 (G5+H5)
5,Nereida,Magwood,Female,United States,DateTime_2,58,2468,2526 (G6+H6)
6,Gaston,Brumm,Male,United States,DateTime_3,24,2554,2578 (G7+H7)
```
<sup><a href='/src/Tests/Samples.VerifyExcelStream.DotNet9_0.verified.csv#L1-L7' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyExcelStream.DotNet9_0.verified.csv' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Excel file

<img src="/src/Tests/Samples.VerifyExcel.DotNet9_0_Sheet1.png">
