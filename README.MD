# ExcelImportExport

A lightweight C# library for importing and exporting lists of objects (`List<T>`) to and from Excel `.xlsx` files using OpenXML.

> **Purpose:** Provide a simple, dependency-free, and reusable solution for serializing data to Excel, suitable for .NET 6+ applications.

---

## ✨ Features

- 📤 Export generic objects (`List<T>`) to `.xlsx`
- 📥 Import `.xlsx` files into strongly typed objects
- 🧱 Based on pure `DocumentFormat.OpenXml` (no external dependencies)
- ♻️ Designed to be reused across multiple projects
- 🔒 Compatible with open-source and commercial environments

---

## 📦 Installation

This is a class library. You can:
- Copy the `.cs` file into your project
- Or reference the compiled `.dll` directly

Install `DocumentFormat.OpenXml` via NuGet:

```bash
dotnet add package DocumentFormat.OpenXml
