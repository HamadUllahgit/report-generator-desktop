# Automated Report Generator

A .NET 8 Windows Forms desktop application for generating professional assessment reports with Word/PDF output, database storage, and email integration.

## The Problem

A consulting firm was manually creating assessment reports in Word. Each report required:

- Copying data from websites into forms
- Manually filling 25+ fields in Word templates
- Tracking job IDs in spreadsheets
- Saving reports with consistent naming
- Attaching PDFs to emails

The process took 15-20 minutes per report and was prone to errors.

## The Solution

Built a desktop application that automates the entire workflow:

1. **Paste website data** → Auto-parses address, lot, plan numbers
2. **Fill form** → Dropdowns with saved options, validation
3. **Add images** → Paste from clipboard or browse
4. **Click Generate** → Word doc + PDF + database record + email draft

**Time per report: 2-3 minutes**

## Features

### Smart Data Entry

| Feature | Description |
|---------|-------------|
| Clipboard Parser | Paste raw website text, extracts structured data automatically |
| Auto-complete Dropdowns | User-added options persist across sessions |
| Field Validation | Required fields highlighted, prevents incomplete submissions |
| Auto Job ID | Sequential IDs generated from database |

### Document Generation

| Feature | Description |
|---------|-------------|
| Word Templates | Content controls + `[[placeholders]]` replaced with form data |
| Image Insertion | Up to 5 images inserted into designated template locations |
| PDF Export | Automatic conversion via Word automation |
| Consistent Naming | `{JobID}_{timestamp}/` folder structure |

### Database Integration

| Feature | Description |
|---------|-------------|
| MS Access Backend | Local database for job tracking |
| Auto Column Mapping | Intelligent matching of form fields to database columns |
| Upsert Logic | Updates existing records or inserts new ones |

### Email Integration

| Feature | Description |
|---------|-------------|
| Outlook Draft | Creates email with PDF attached |
| Auto Subject | Populated from job data |
| One-Click Send | Review and send from Outlook |

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Windows Forms UI                          │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐         │
│  │ Form Fields │  │  Actions    │  │   Images    │         │
│  │ (25+ inputs)│  │  Panel      │  │  (5 slots)  │         │
│  └─────────────┘  └─────────────┘  └─────────────┘         │
└─────────────────────────────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────┐
│                    Generation Engine                         │
│                                                              │
│  ┌──────────┐    ┌──────────┐    ┌──────────┐              │
│  │  Parse   │───▶│ Validate │───▶│  Build   │              │
│  │  Input   │    │  Fields  │    │  Output  │              │
│  └──────────┘    └──────────┘    └──────────┘              │
│                                        │                    │
│                    ┌───────────────────┼───────────────┐   │
│                    ▼                   ▼               ▼   │
│              ┌──────────┐       ┌──────────┐    ┌────────┐│
│              │  Word    │──────▶│   PDF    │    │ Access ││
│              │  .docx   │       │  Export  │    │   DB   ││
│              └──────────┘       └──────────┘    └────────┘│
│                    │                                       │
│                    ▼                                       │
│              ┌──────────┐                                  │
│              │ Outlook  │                                  │
│              │  Draft   │                                  │
│              └──────────┘                                  │
└─────────────────────────────────────────────────────────────┘
```

## Technical Implementation

### Clipboard Parser

Extracts structured data from pasted text using regex patterns:

```csharp
// Parse address: "123 Main Street Sydney 2000"
var parts = lines[0].Split(' ');
int pcIndex = parts.FindLastIndex(tok => Regex.IsMatch(tok, @"^\d{4}$"));
// Extract: HouseNo, RoadName, Suburb, Postcode

// Parse lot/plan: "Lot 1/Section 2/DP123456"
var m = Regex.Match(lotLine, @"^(?<lot>[^/]+)/(?<sec>[^/]+)/(?<ptype>[A-Za-z]+)(?<pnum>\d+)$");
```

### Word Automation

Uses COM interop to fill templates:

```csharp
// Fill content controls by tag
foreach (var cc in doc.ContentControls)
{
    string tag = cc.Tag ?? cc.Title;
    if (fields.TryGetValue(tag, out var val))
    {
        cc.LockContents = false;
        cc.Range.Text = val;
    }
}

// Replace [[placeholders]]
doc.Content.Find.Execute(
    FindText: "[[" + key + "]]",
    ReplaceWith: value,
    Replace: 2  // wdReplaceAll
);

// Insert images at tagged locations
range.InlineShapes.AddPicture(imgPath, false, true);
```

### Database Auto-Mapping

Intelligently maps form fields to database columns:

```csharp
var alias = new Dictionary<string, string>
{
    ["JobID"] = dbCols.FirstOrDefault(c => 
        Norm(c) is "jobid" or "reportid" or "id") ?? "JobID",
    ["HouseNo"] = dbCols.FirstOrDefault(c => 
        Norm(c) is "houseno" or "house" or "streetno") ?? "HouseNo",
    // ... more mappings
};
```

### STA Thread for COM

Office automation requires STA thread:

```csharp
private static Task<T> RunSTA<T>(Func<T> func)
{
    var tcs = new TaskCompletionSource<T>();
    var th = new Thread(() => {
        try { tcs.SetResult(func()); }
        catch (Exception ex) { tcs.SetException(ex); }
    });
    th.SetApartmentState(ApartmentState.STA);
    th.Start();
    return tcs.Task;
}
```

## Tech Stack

- **.NET 8** Windows Forms
- **C#** with async/await
- **Microsoft Word** COM automation
- **Microsoft Outlook** COM automation
- **MS Access** via OleDb
- **System.Text.Json** for config/data
- **Regex** for text parsing

## Project Structure

```
ReportApp/
├── src/
│   ├── Program.cs           → Entry point
│   ├── MainForm.cs          → UI and generation logic
│   ├── AppConfig.cs         → Configuration model
│   ├── FieldMapping.cs      → Form field definitions
│   ├── JobIdSequence.cs     → Auto-increment ID generator
│   └── AccessIdGuard.cs     → Database ID management
├── Templates/
│   └── BAL_Template.docx    → Word template with placeholders
├── Data/
│   └── database.accdb       → MS Access database
└── Reports/
    └── {JobID}_{timestamp}/ → Generated output folders
        ├── {JobID}.docx
        ├── {JobID}.pdf
        ├── fields.json
        └── Images/
```

## Configuration

```json
{
  "TemplatesDir": "./Templates",
  "BaseOutputDir": "./Reports",
  "AccessDbPath": "./Data/database.accdb",
  "AccessTable": "Jobs",
  "AccessKeyCol": "JobID",
  "OpenOutlookDraft": true
}
```

## Results

| Metric | Before | After |
|--------|--------|-------|
| Time per report | 15-20 min | 2-3 min |
| Data entry errors | Common | Rare (validation) |
| Naming consistency | Inconsistent | Automated |
| Job tracking | Manual spreadsheet | Database |

## Key Learnings

1. **COM requires STA** - Office automation must run on Single-Threaded Apartment thread

2. **Content controls > bookmarks** - More reliable for repeated template use

3. **Auto-mapping saves setup** - Client doesn't need to configure database columns

4. **Clipboard parsing is powerful** - Regex can extract structured data from messy website pastes

---

## About Me

I'm Hamad, a Microsoft 365 & Power Platform consultant. This project showcases my ability to build desktop solutions that integrate with the Microsoft Office ecosystem.

- 🌐 [hamad365.com](https://hamad365.com)
- 💼 [Upwork](https://www.upwork.com/freelancers/~01f28bdcd32df9fe01)
- 📧 h@hamad365.com

---

*Note: Client-specific details and sensitive data have been removed. This case study documents the technical approach and architecture.*
