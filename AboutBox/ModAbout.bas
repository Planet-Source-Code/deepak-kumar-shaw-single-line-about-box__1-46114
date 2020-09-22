Attribute VB_Name = "ModAbout"
Option Explicit
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

''** hWnd
''Identifies a parent window. This parameter can be NULL.
''
''** szApp
''Points to text that the function displays in the title bar of the Shell About dialog box and on the first line of the dialog box after the text “Microsoft Windows” or “Microsoft Windows NT.” If the text contains a “#” separator, dividing it into two parts, the function displays the first part in the title bar, and the second part on the first line after the text “Microsoft Windows” or “Microsoft Windows NT.”
''
''** szOtherStuff
''Points to text that the function displays in the dialog box after the version and copyright information.
''
''** hIcon
''Identifies an icon that the function displays in the dialog box. If this parameter is NULL, the function displays the Microsoft Windows or Microsoft Windows NT icon.
