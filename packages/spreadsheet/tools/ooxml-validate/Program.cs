// Checks .xlsx and .xlsm files against the Office Open XML schemas and Excel's
// own rules, with Microsoft's Open XML SDK: the closest automated stand-in for
// "Excel opens it without offering to repair it". Prints one JSON line per
// file: { file, errors: [{ part, path, id, description }] }.
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var failed = false;
foreach (var file in args)
{
    var errors = new List<object>();
    try
    {
        using var doc = SpreadsheetDocument.Open(file, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        foreach (var e in validator.Validate(doc).Take(500))
            errors.Add(new { part = e.Part?.Uri.ToString(), path = e.Path?.XPath, id = e.Id, description = e.Description });
    }
    catch (Exception ex)
    {
        errors.Add(new { part = (string?)null, path = (string?)null, id = "open", description = ex.Message });
    }
    if (errors.Count > 0) failed = true;
    Console.WriteLine(JsonSerializer.Serialize(new { file, errors }));
}
return failed ? 1 : 0;
