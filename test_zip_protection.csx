#!/usr/bin/env dotnet-script

// Test script to verify zip bomb protection works

using System;
using System.IO;
using System.IO.Compression;
using OfficeForge;
using OfficeForge.Models;

Console.WriteLine("Testing zip bomb protection...\n");

// Test 1: ReaderOptions validation
Console.WriteLine("Test 1: ReaderOptions validation");
try
{
    var badOptions = new ReaderOptions
    {
        MaxUncompressedSize = -1,
        MaxEntryCount = -1,
        MaxCompressionRatio = 0,
        MaxPartSize = -1
    };
    badOptions.Validate();
    Console.WriteLine("FAIL: Should have thrown for negative values");
    Environment.Exit(1);
}
catch (ArgumentException ex)
{
    Console.WriteLine($"PASS: Caught expected exception: {ex.Message}");
}

// Test 2: Default options have reasonable values
Console.WriteLine("\nTest 2: Default options values");
var defaultOptions = ReaderOptions.Default;
Console.WriteLine($"Default MaxUncompressedSize: {defaultOptions.MaxUncompressedSize}");
Console.WriteLine($"Default MaxEntryCount: {defaultOptions.MaxEntryCount}");
Console.WriteLine($"Default MaxCompressionRatio: {defaultOptions.MaxCompressionRatio}");
Console.WriteLine($"Default MaxPartSize: {defaultOptions.MaxPartSize}");

if (defaultOptions.MaxUncompressedSize <= 0 ||
    defaultOptions.MaxEntryCount <= 0 ||
    defaultOptions.MaxCompressionRatio <= 0 ||
    defaultOptions.MaxPartSize <= 0)
{
    Console.WriteLine("FAIL: Default options have invalid values");
    Environment.Exit(1);
}
Console.WriteLine("PASS: Default options are valid");

// Test 3: DocumentTooLargeException properties
Console.WriteLine("\nTest 3: DocumentTooLargeException properties");
try
{
    throw new DocumentTooLargeException("Test message")
    {
        PartName = "TestPart",
        MaxLimit = 1000,
        ActualValue = 2000,
        LimitType = DocumentTooLargeException.SizeLimitType.MaxPartSize
    };
}
catch (DocumentTooLargeException ex)
{
    Console.WriteLine($"PASS: Exception thrown with properties:");
    Console.WriteLine($"  Message: {ex.Message}");
    Console.WriteLine($"  PartName: {ex.PartName}");
    Console.WriteLine($"  MaxLimit: {ex.MaxLimit}");
    Console.WriteLine($"  ActualValue: {ex.ActualValue}");
    Console.WriteLine($"  LimitType: {ex.LimitType}");
}

// Test 4: Create a minimal valid docx and try to read it
Console.WriteLine("\nTest 4: Reading valid document");
try
{
    // Create a minimal valid docx in memory
    using var memStream = new MemoryStream();
    using (var archive = new ZipArchive(memStream, ZipArchiveMode.Create, true))
    {
        var entry = archive.CreateEntry("[Content_Types].xml");
        using var entryStream = entry.Open();
        using var writer = new StreamWriter(entryStream);
        writer.Write("<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"></Types>");
    }
    memStream.Position = 0;

    var reader = new DocxReader();
    var doc = reader.Read(memStream);
    Console.WriteLine("PASS: Successfully read minimal document");
}
catch (Exception ex)
{
    Console.WriteLine($"FAIL: Failed to read minimal document: {ex.Message}");
    Environment.Exit(1);
}

// Test 5: Try to read with restrictive options
Console.WriteLine("\nTest 5: Reading with restrictive options");
try
{
    var strictOptions = new ReaderOptions
    {
        MaxUncompressedSize = 100, // Very small limit
        MaxEntryCount = 1,
        MaxCompressionRatio = 1,
        MaxPartSize = 100
    };

    using var memStream = new MemoryStream();
    using (var archive = new ZipArchive(memStream, ZipArchiveMode.Create, true))
    {
        var entry = archive.CreateEntry("test.txt");
        using var entryStream = entry.Open();
        using var writer = new StreamWriter(entryStream);
        writer.Write("test");
    }
    memStream.Position = 0;

    var reader = new DocxReader();
    var doc = reader.Read(memStream, strictOptions);
    Console.WriteLine("FAIL: Should have thrown DocumentTooLargeException");
    Environment.Exit(1);
}
catch (DocumentTooLargeException)
{
    Console.WriteLine("PASS: Caught expected DocumentTooLargeException with restrictive options");
}
catch (Exception ex)
{
    Console.WriteLine($"FAIL: Caught wrong exception type: {ex.GetType().Name}");
    Environment.Exit(1);
}

Console.WriteLine("\n=== ALL TESTS PASSED ===");
