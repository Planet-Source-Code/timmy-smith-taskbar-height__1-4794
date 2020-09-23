<div align="center">

## TaskBar Height


</div>

### Description

The purpose of the code is to determine the height of the taskbar so that you can display a form at the bottom right hand corner of the screen everytime.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Timmy Smith](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/timmy-smith.md)
**Level**          |Unknown
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/timmy-smith-taskbar-height__1-4794/archive/master.zip)

### API Declarations

```
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
```


### Source Code

```
' this code will display a form at the bottom right had corner everytime.
   dim WindowRect as RECT
   SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
   FrmMain.Top = WindowRect.Bottom * Screen.TwipsPerPixelY - FrmMain.Height
   FrmMain.Left = WindowRect.Right * Screen.TwipsPerPixelX - FrmMain.Width
```

