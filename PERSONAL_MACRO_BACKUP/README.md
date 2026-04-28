# PERSONAL.XLSB Macro Backup

This folder stores the Excel `PERSONAL.XLSB` macro workbook backup.

Use it when moving to a new laptop or restoring your Excel macro buttons.

## Restore Steps

1. Close Excel.
2. Copy `PERSONAL.XLSB`.
3. Paste it into:

```text
C:\Users\<YourUser>\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB
```

4. Open Excel.
5. Your personal macros should load automatically.

## Quick Access Toolbar Buttons

`PERSONAL.XLSB` restores the macros, but it may not restore the toolbar icons.

To restore buttons too, also back up and import the Excel customization file:

```text
Excel > File > Options > Quick Access Toolbar > Import/Export
```

Recommended backup pair:

- `PERSONAL.XLSB`
- `Excel Buttons.exportedUI`

