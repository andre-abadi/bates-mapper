# Bates Mapper

PowerShell tool that finds Bates numbers in a Word document and replaces each with its new number from an Excel mapping file. The Word document is saved in place only if at least one replacement was made; the Excel file is never modified. Each change is logged to the console.

## Requirements

- Windows with PowerShell
- Microsoft Excel and Microsoft Word installed (used through COM)

## Usage

1. Place the `.xlsx` mapping file and the `.docx` document in the same folder as the script.
2. Run from that folder:

   ```powershell
   .\bates-mapper.ps1
   ```

3. Pick the Excel file, then the Word document, from the numbered lists (blank quits).
4. Answer the fuzzy mode prompt (`Y` or `y` = on, anything else = off).

## Mapping file

The first worksheet is used. The script scans each row (up to `$Global:traverse_rows`, default 10), across columns 1 to `$Global:traverse_columns` (default 5), for the first Bates number. That cell's row is where the data starts and its column holds the old number; the column immediately to its right holds the new number. Other columns are ignored. If the next column contains no Bates numbers, the found column is treated as a list of correct numbers only: numbers in the list are left alone, and with fuzzy mode on any other number falls back to the nearest lower number in the list (without fuzzy mode it is skipped). Rows whose old-number cell does not match the Bates pattern (such as headers) are skipped. Duplicate old numbers keep the first occurrence, and rows with a blank column B are reported as warnings and not mapped. Any number in the document that already appears in column B is left untouched, so the tool is safe to run repeatedly (including in fuzzy mode).

## Bates format

Defined by `$Global:BatesPattern` at the top of the script. Default: `ABC.001.001.0001` (`[A-Z]{3}\.\d{3}\.\d{3}\.\d{4}`).

All document stories are scanned: body, footnotes, endnotes, headers, footers and text boxes.

## Output

Each Bates number found is printed with its result:

| Colour | Meaning |
| --- | --- |
| Green | Mapped to a new number |
| Dark gray | Mapped, number unchanged |
| Cyan | Fuzzy match, noted `fuzzy` |
| Yellow | No mapping, skipped |

A summary of totals follows.

## Fuzzy mode

When on, a number that is not in the mapping file takes the new value of the closest lower mapped number with the same prefix (everything before the last four digits). For example, with `ABC.001.001.0001` and `ABC.001.001.0004` mapped, `ABC.001.001.0002` is treated as a page of `ABC.001.001.0001` and gets its new number. If no lower mapped number exists in that prefix, the number is skipped.

## Built-in test resources

| File | Purpose |
| --- | --- |
| `incorrect.docx` | Sample document containing Bates numbers, some of which are not in the mapping |
| `mapping.xlsx` | Sample mapping: column A = incorrect number, column B = correct number, column C = page (informational, not read) |
| `incorrect.zip` | Pristine copy of `incorrect.docx` |
| `reset.ps1` | Extracts every `.zip` in the folder over the current directory, overwriting existing files, to restore the originals |

Run `.\reset.ps1` to restore the sample document after it has been changed or deleted. It restores only what is in the zip; anything missing from the zip is not recreated.

### What the sample data exercises

- **Direct mappings**, for example `ABC.001.001.0002` -> `ABC.001.001.0001` and `XYZ.002.005.0102` -> `XYZ.002.005.0100`.
- **Unchanged mappings**, `ABC.001.001.0025` and `ABC.001.001.0030`, which map to themselves.
- **Unmapped numbers**, `ABC.001.001.0045` and `XYZ.002.005.0110`, which are skipped with fuzzy mode off.
- **Fuzzy matches** with fuzzy mode on:
  - `ABC.001.001.0045` -> `ABC.001.001.0030` (closest lower mapped number is `0030`)
  - `XYZ.002.005.0110` -> `XYZ.002.005.0104` (closest lower mapped number is `0106`, which maps to `0104`)
- **Rows with no new number**, `ABC.001.002.0456` and `ABC.002.003.0678`, which appear only in column A and are reported as warnings.

### Suggested test run

```powershell
.\reset.ps1
.\bates-mapper.ps1   # run once with fuzzy off, once with fuzzy on, and compare
```
