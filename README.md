# officeforge

Cross-platform .NET library and CLI for creating, editing and converting Word/Excel/PowerPoint documents on any OS - no Microsoft Office, no COM interop. Built on the OpenXML SDK.

```bash
dotnet add package OfficeForge
officeforge write-cell report.xlsx B2 42.5
officeforge convert report.xlsx --format markdown
```
