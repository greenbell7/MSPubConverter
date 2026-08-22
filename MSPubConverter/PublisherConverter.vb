Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Publisher
Imports System.Linq

Public Class PublisherConverter
    Private publisher As Application
    Private ReadOnly logger As ILogger

    Public Sub New(Optional logger As ILogger = Nothing)
        Me.logger = If(logger, New ConsoleLogger())
    End Sub

    ''' <summary>
    ''' Converts a single Publisher (.pub) file to PDF
    ''' </summary>
    Public Function ConvertPubToPdf(inputPub As String, outputPdf As String) As Boolean
        inputPub = Path.GetFullPath(inputPub)
        outputPdf = Path.GetFullPath(outputPdf)
        Dim doc As Document = Nothing

        Try
            logger.Log("Opening: " & inputPub)
            doc = publisher.Open(inputPub)

            Dim openedName As String = Nothing
            Dim openedPath As String = Nothing

            Try
                openedName = doc.FullName
                openedPath = doc.Path
                logger.Log($"Opened doc FullName={openedName} Path={openedPath}")
            Catch ex As Exception
                logger.LogException("Unable to read opened document properties", ex)
            End Try

            If doc Is Nothing OrElse Not String.Equals(openedName, inputPub, StringComparison.OrdinalIgnoreCase) Then
                logger.LogError($"Publisher did not open the requested file; opened={openedName} expected={inputPub}")
                Return False
            End If

            ' Ensure target directory exists
            Dim outDir = Path.GetDirectoryName(outputPdf)
            If Not String.IsNullOrEmpty(outDir) AndAlso Not Directory.Exists(outDir) Then
                Directory.CreateDirectory(outDir)
            End If

            logger.Log("Attempting ExportAsFixedFormat: " & outputPdf)

            Try
                Dim fmt_pdf = PbFixedFormatType.pbFixedFormatTypePDF
                Dim intent_print = PbFixedFormatIntent.pbFixedFormatIntentPrint
                doc.ExportAsFixedFormat(fmt_pdf, outputPdf, intent_print)
            Catch ex As Exception
                logger.LogException("ExportAsFixedFormat failed", ex)
                Return False
            End Try

            ' Verify file was created
            If Not File.Exists(outputPdf) Then
                logger.LogError("Export completed but output file not found: " & outputPdf)
                Return False
            End If

            logger.Log("Export succeeded: " & outputPdf)
            Return True

        Catch ex As Exception
            logger.LogException($"ERROR converting {inputPub}", ex)
            Return False
        Finally
            If doc IsNot Nothing Then
                Try
                    doc.Close()
                Catch ex As Exception
                    logger.LogException($"Error closing document for {inputPub}", ex)
                End Try
            End If
        End Try
    End Function

    ''' <summary>
    ''' Converts all .pub files in a folder hierarchy to PDF
    ''' </summary>
    Public Sub ConvertAllPubFiles(parentFolder As String, outputRoot As String, preservePaths As Boolean)
        ' Collect all .pub files first
        Dim pubFiles As New List(Of String)
        For Each filePath In Directory.EnumerateFiles(parentFolder, "*.pub", SearchOption.AllDirectories)
            pubFiles.Add(filePath)
        Next

        If pubFiles.Count = 0 Then
            Console.WriteLine("No .pub files found.")
            Return
        End If

        Try
            InitializePublisher()

            For Each inputPath In pubFiles
                ' Determine output folder based on user's choice to preserve structure
                Dim outputDir As String
                If preservePaths Then
                    Dim relativePath = GetRelativePath(parentFolder, Path.GetDirectoryName(inputPath))
                    If relativePath = "." Then
                        outputDir = outputRoot
                    Else
                        outputDir = Path.Combine(outputRoot, relativePath)
                    End If
                Else
                    ' Save all files directly into the parent output folder
                    outputDir = outputRoot
                End If

                outputDir = Path.GetFullPath(outputDir)
                Directory.CreateDirectory(outputDir)

                Dim baseName = Path.GetFileNameWithoutExtension(inputPath)
                Dim camelCase = ToCamelCase(baseName)
                If String.IsNullOrEmpty(camelCase) Then
                    camelCase = "file"
                End If

                ' Generate a unique, no-spaces camelCase filename
                Dim pdfName = UniqueFilenameNoSpaces(outputDir, camelCase, ".pdf")
                Dim outputPath = Path.Combine(outputDir, pdfName)
                outputPath = Path.GetFullPath(outputPath)

                Dim success = ConvertPubToPdf(inputPath, outputPath)
                If Not success Then
                    logger.LogWarning("Failed to convert: " & inputPath)
                End If
            Next

        Finally
            ReleasePublisher()
        End Try
    End Sub

    ''' <summary>
    ''' Converts a string to camelCase
    ''' </summary>
    Public Shared Function ToCamelCase(s As String) As String
        If String.IsNullOrEmpty(s) Then
            Return String.Empty
        End If

        ' Split on any non-alphanumeric characters, drop empties
        Dim parts = Regex.Split(s, "[^A-Za-z0-9]+").Where(Function(p) Not String.IsNullOrEmpty(p)).ToList()

        If parts.Count = 0 Then
            Return String.Empty
        End If

        Dim first = parts(0).ToLower()
        Dim rest = String.Concat(parts.Skip(1).[Select](Function(p) Char.ToUpper(p(0)) & p.Substring(1)))

        Return first & rest
    End Function

    ''' <summary>
    ''' Returns a unique filename, appending -2, -3, etc. if file exists
    ''' </summary>
    Public Shared Function UniqueFilenameNoSpaces(outputDir As String, baseName As String, Optional ext As String = ".pdf") As String
        Dim candidate = baseName & ext
        Dim counter = 2

        While File.Exists(Path.Combine(outputDir, candidate))
            candidate = $"{baseName}-{counter}{ext}"
            counter += 1
        End While

        Return candidate
    End Function

    ''' <summary>
    ''' Gets the relative path between two directories
    ''' </summary>
    Private Shared Function GetRelativePath(basePath As String, fullPath As String) As String
        basePath = Path.GetFullPath(basePath)
        fullPath = Path.GetFullPath(fullPath)

        If fullPath = basePath Then
            Return "."
        End If

        Dim baseUri = New Uri(basePath & Path.DirectorySeparatorChar)
        Dim fullUri = New Uri(fullPath)
        Return baseUri.MakeRelativeUri(fullUri).ToString().Replace("/", Path.DirectorySeparatorChar)
    End Function

    Private Sub InitializePublisher()
        Try
            publisher = New Application With {
                .Visible = False
            }
        Catch ex As Exception
            logger.LogException("Error initializing Publisher", ex)
        End Try
    End Sub

    Private Sub ReleasePublisher()
        If publisher IsNot Nothing Then
            Try
                publisher.Quit()
            Catch ex As Exception
                logger.LogException("Error quitting Publisher application", ex)
            End Try
        End If
    End Sub
End Class

''' <summary>
''' Logger interface for extensibility
''' </summary>
Public Interface ILogger
    Sub Log(message As String)
    Sub LogWarning(message As String)
    Sub LogError(message As String)
    Sub LogException(message As String, ex As Exception)
End Interface

''' <summary>
''' Simple console-based logger
''' </summary>
Public Class ConsoleLogger
    Implements ILogger

    Public Sub Log(message As String) Implements ILogger.Log
        Console.WriteLine($"[INFO] {DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}")
    End Sub

    Public Sub LogWarning(message As String) Implements ILogger.LogWarning
        Console.WriteLine($"[WARNING] {DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}")
    End Sub

    Public Sub LogError(message As String) Implements ILogger.LogError
        Console.WriteLine($"[ERROR] {DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}")
    End Sub

    Public Sub LogException(message As String, ex As Exception) Implements ILogger.LogException
        Console.WriteLine($"[ERROR] {DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}")
        Console.WriteLine($"Exception: {ex.Message}")
        Console.WriteLine($"StackTrace: {ex.StackTrace}")
    End Sub
End Class