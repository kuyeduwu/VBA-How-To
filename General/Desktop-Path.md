```vb
Sub DskPath()
    Debug.Print CreateObject("WSCript.shell").SpecialFolders("Desktop")
End Sub
```

The `SpecialFolders` property could return the path for may folders in Windows system, including:
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
