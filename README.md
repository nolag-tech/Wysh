# Wysh : An alternative to Windows Scripting Host

### Why does this exist?

The goals are simple: Provide an environment that has the flexibility of modern JavaScript via V8 while retaining the functionality of a native Windows scripting environment similar enough to Windows Script Host that code can be easily ported.

With Windows 11 24H2, it has become clear that Windows Script Host (WSH) is deprecated even though Microsoft doesn't formally say it. There were changes made to the JScript engine that were meant to address issues with Internet Explorer, but spilled over into WSH. Also, Microsoft has already announced their plans to remove VBScript entirely by 2027. WSH is seen as feature-complete and has not received any updates in quite a while. Also, PowerShell is the scripting environment preferred by Microsoft.

PowerShell is a great environment for many scripting tasks. However, if you have a business with dependency on WSH for Office automation and other tasks, you may find PowerShell to be awkward, daunting, or limiting. Moreover, existing JScript code cannot be moved easily to PowerShell.


### What's broken in Wysh?

Wysh is very early - really, just proof-of-concept stage. As such, there are many things broken and many more things that are subject to change.

First, let's look at the JavaScript functionality that is missing or problematic in Wysh:
* `setTimeout`, `setInterval`, `clearTimeout`, `clearInterval`
* `Promise`
* `console.log` and related functions don't do anything. However, the .NET `Console` object is exposed, so one can use `Console.WriteLine` and `Console.ReadLine`.
* The `WScript` API is mostly missing. This isn't WSH. Some parts of the original API are available. See below.
* This is not a Node environment, so none of the modules from Node are available. There is no plan to fix this. If you want Node, use Node.
* You can import scripts that are in the same directory structure. There are no global modules *yet*.
* Support for XML script files is in the works.
* Elegant stack traces from scripts. Better exception handling is needed.


### What's working?

* `WScript.Shell` is exposed as `Shell`
* `WScript.Network` is exposed as `Network`
* `FileSystemObject`
* `Dictionary`
* `ADODB.Stream` is typically used for buffering and is available as `StreamBuffer`
* `CDO.Message` is exposed as `CDOMessage` (May change. See below.)
* `CDO.Configuration` is exposed as `CDOConfiguration` (May change. See below.)
* `Excel.Application` is exposed as `ExcelApp` (May change. See below.)
* `Msxml2.DOMDocument.6.0` is exposed as `DOMDocument` (May change. See below.)

### What about the original Wysh functionality?

Wysh has been several things along the way. Early on, it was called Foundation, then GreasePencil, now Wysh. It was not a runtime environment originally, but a collection of libraries that provided a rich scripting environment for reporting tools and task automation.

Wysh will continue in this tradition by making available most of the APIs that were once part of these earlier iterations. Some of the older APIs will be retired, but the idea is to keep most of the environment intact while adding .NET and V8 functionality.

Not all of this is available yet, but will be added in short order.

* Excel three ways - Excel will be available through the traditional COM API as well as the augmented Wysh version. An interface to the .NET API will be added later.
* SQL, ODBC, etc - All the things that made querying easy in Wysh will be ported. This includes embedded queries.
* XML - XML-formatted script files will be supported along with the Resources API. Essentially, this allows non-flat script files that carry embedded SQL queries, JSON blobs, etc.
* The `CDO` APIs are exposed now (mentioned above), but may disappear in favor of their Wysh counterparts.

### What about SiteServASP and Trapese?

These environments have been retired. Both relied on Classic ASP and have not been used or maintained in over a decade. There is no server-side equivalent of Wysh planned.

### Improvements?

.NET offers many improvements over WSH / COM environments.

* SQL/ODBC functionality might be ported to .NET native.
* The original Wysh API for Excel may end up ported to .NET.
* Functionality for CrystalReports might be added. (This was a separate binary before.)


### How do I use this?

```cmd
C:\Users\User> wysh myScript.js
```

myScript.js
```javascript

Console.WriteLine("Hello World!");

// Write output to a buffer
let myOutput = new Stream();
myOutput.WriteLine("This is a test.");
myOutput.WriteLine("We are adding lines ...");
myOutput.WriteLine("... to a buffer.");
myOutput.Position = 0;

// Write a buffer to a file
let fso = new FileSystemObject();
let tempFolder = fso.GetSpecialFolder(2);
let tempFile = fso.CreateTextFile(tempFolder + "\tempFile.txt", true);
tempFile.Write(myOutput.ReadText(-1));
tempFile.Close();


// Simple Excel automation
oExcel = new ExcelApp();
oExcel.Visible = true;
oExcel.Workbooks.Add();
oExcel.ActiveWorkbook.ActiveSheet.Name = "HelloSheet";
```