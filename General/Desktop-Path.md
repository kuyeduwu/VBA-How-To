```vb
Public Function DskPath() As String
    Let DskPath = CStr(CreateObject("WSCript.shell").SpecialFolders("Desktop"))
End Function
```

The `SpecialFolders` property could return the path for many folders in Windows system, including:
- AllUsersDesktop
- AllUsersStartMenu
- AllUsersPrograms
- AllUsersStartup
- AppData
- Desktop
- Favorites
- Fonts
- MyDocuments
- NetHood
- PrintHood
- Programs
- Recent
- SendTo
- StartMenu
- Startup
- Templates
