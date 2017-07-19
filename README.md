# -VBA
#生成菜单
*****************************************
Public Sub Workbook_Open()
    Dim NewMenu As CommandBarPopup
    Dim MenuItem As CommandBarControl
    On Error Resume Next
    '如果菜单已存在,则删除该菜单
    Application.CommandBars(1).Controls("价格标签").Delete
    Set NewMenu = Application.CommandBars(1).Controls.Add(Type:=msoControlPopup, Before:=11)
     
    '添加菜单标题并指定热键
    NewMenu.Caption = "价格标签"
    
    '添加第一个菜单项
    Set MenuItem = NewMenu.Controls.Add _
      (Type:=msoControlButton)
    With MenuItem
        .Caption = "生成表"
        .OnAction = "生成表"
    End With
    
        
        Set MenuItem = NewMenu.Controls.Add(Type:=msoControlButton)
    With MenuItem
    .Caption = "生成标签"
    .OnAction = "价格标签生成"
    End With
    
    
        Set MenuItem = NewMenu.Controls.Add(Type:=msoControlButton)
    With MenuItem
    .Caption = "打印布局"
    .OnAction = "打印"

    End With
End Sub
##模块代码
*************************************************************
Sub 填充颜色()
    Range("B3:D4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10043393
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B3:H3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10043393
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B5:H5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10043393
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
   Range("B13:H13").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10043393
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B15:H15").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10043393
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("F5:F12").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10043393
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H3:H15").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10043393
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B3:B15").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10043393
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("G6:G12").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 61951
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub 输入文字()
    Range("C4:D4").Select
    ActiveCell.FormulaR1C1 = "品 牌"
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "品名："
        Range("C8").Select
    ActiveCell.FormulaR1C1 = "型号："
        Range("C9").Select
    ActiveCell.FormulaR1C1 = "产地："
        Range("C10").Select
    ActiveCell.FormulaR1C1 = "规格："
        Range("C11").Select
    ActiveCell.FormulaR1C1 = "单位："
    Range("C14:G14").Select
    ActiveCell.FormulaR1C1 = "全国24小时统一服务热线：8008288988"
        Range("G7").Select
    ActiveCell.FormulaR1C1 = "全国统一零售价："
  End Sub
Sub 合并单元格()
   Range("C14:G14").Select
    Selection.Merge
     Range("D7:E7").Select
    Selection.Merge
    Range("D8:E8").Select
    Selection.Merge
    Range("D9:E9").Select
    Selection.Merge
    Range("D10:E10").Select
    Selection.Merge
    Range("D11:E11").Select
    Selection.Merge
        Range("E4:G4").Select
    Selection.Merge
    Range("C4:D4").Select
    Selection.Merge
    Range("G8:G10").Select
    Selection.Merge
    End Sub
Sub 设置字体()
    Range("C4:D4").Select
    With Selection.Font
    .Bold = True
        .Name = "黑体"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("C7:C11").Select
     With Selection.Font
    .Bold = False
        .Name = "微软雅黑"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
   
    Range("E7:E11").Select
     With Selection.Font
    .Bold = False
        .Name = "宋体"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
   
     Range("G7").Select
     With Selection.Font
    .Bold = False
        .Name = "微软雅黑"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
     Range("C14:G14").Select
    With Selection.Font
        .Name = "微软雅黑"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
     Range("G8:G10").Select
    With Selection.Font
        .Name = "微软雅黑"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.NumberFormatLocal = "￥#,##0;￥-#,##0"
End Sub
Sub 调行列距()
Rows("3:3").Select
    Selection.RowHeight = 7
    Rows("4:4").Select
    Selection.RowHeight = 28
    Rows("5:6").Select
    Selection.RowHeight = 3
    Rows("7:11").Select
    Selection.RowHeight = 44
    Rows("12:13").Select
    Selection.RowHeight = 3
    Rows("14:14").Select
    Selection.RowHeight = 26.5
    Rows("15:15").Select
    Selection.RowHeight = 7
    Rows("16:16").Select
    Selection.RowHeight = 1.5
 Columns("B:B").Select
    Selection.ColumnWidth = 0.77
    Columns("C:C").Select
    Selection.ColumnWidth = 4.38
    Columns("D:D").Select
    Selection.ColumnWidth = 1.88
    Columns("E:E").Select
    Selection.ColumnWidth = 9.25
    Columns("F:F").Select
    Selection.ColumnWidth = 0.23
    Columns("G:G").Select
    Selection.ColumnWidth = 14.63
    Columns("H:H").Select
    Selection.ColumnWidth = 0.77
    Columns("I:I").Select
    Selection.ColumnWidth = 0.15
End Sub
Sub 生成排版()
On Error Resume Next
h = Application.WorksheetFunction.CountA(Sheets("价格表").Range("A:A"))
i = Application.WorksheetFunction.RoundUp(h / 3, 0) * 14 + 2

'每排四标签
    Columns("B:I").Select
    Selection.Copy
    Columns("J:AF").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

 插入logo
  '复制内容第一排
  

    Range("C4:G14").Select
    Selection.Copy
    Range("K4").Select
    ActiveSheet.Paste
    Range("S4").Select
    ActiveSheet.Paste
    Range("AA4").Select
    ActiveSheet.Paste
录入价格

   
 '复制格式
Rows("3:16").Copy
    Rows("17:" & i).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.DisplayAlerts = False

'复制内容多页
Rows("3:16").Select
    Selection.Copy
    For x = 17 To i Step 14
   Rows(x & ":" & i).Select
    ActiveSheet.Paste
    Next x
Application.CutCopyMode = False
If h < 4 Then
Rows("16:17").Select
    Selection.Delete Shift:=xlUp
End If


'零值不显示
'Cells.Select
'Selection.NumberFormatLocal = "[=0]"""";G/通用格式"
Range("A1").Select
End Sub

Sub 录入价格()
Range("D7").FormulaR1C1 = "=INDIRECT(""价格表!A""&ROUNDDOWN(ROW()/14,0)*4+2)"
Range("D8").FormulaR1C1 = "=INDIRECT(""价格表!B""&ROUNDDOWN(ROW()/14,0)*4+2)"
Range("D9").FormulaR1C1 = "=INDIRECT(""价格表!C""&ROUNDDOWN(ROW()/14,0)*4+2)"
Range("D10").FormulaR1C1 = "=INDIRECT(""价格表!D""&ROUNDDOWN(ROW()/14,0)*4+2)"
Range("D11").FormulaR1C1 = "=INDIRECT(""价格表!E""&ROUNDDOWN(ROW()/14,0)*4+2)"
Range("G8").FormulaR1C1 = "=INDIRECT(""价格表!F""&ROUNDDOWN(ROW()/14,0)*4+2)"

Range("L7").FormulaR1C1 = "=INDIRECT(""价格表!A""&ROUNDDOWN(ROW()/14,0)*4+3)"
Range("L8").FormulaR1C1 = "=INDIRECT(""价格表!B""&ROUNDDOWN(ROW()/14,0)*4+3)"
Range("L9").FormulaR1C1 = "=INDIRECT(""价格表!C""&ROUNDDOWN(ROW()/14,0)*4+3)"
Range("L10").FormulaR1C1 = "=INDIRECT(""价格表!D""&ROUNDDOWN(ROW()/14,0)*4+3)"
Range("L11").FormulaR1C1 = "=INDIRECT(""价格表!E""&ROUNDDOWN(ROW()/14,0)*4+3)"
Range("O8").FormulaR1C1 = "=INDIRECT(""价格表!F""&ROUNDDOWN(ROW()/14,0)*4+3)"

Range("T7").FormulaR1C1 = "=INDIRECT(""价格表!A""&ROUNDDOWN(ROW()/14,0)*4+4)"
Range("T8").FormulaR1C1 = "=INDIRECT(""价格表!B""&ROUNDDOWN(ROW()/14,0)*4+4)"
Range("T9").FormulaR1C1 = "=INDIRECT(""价格表!C""&ROUNDDOWN(ROW()/14,0)*4+4)"
Range("T10").FormulaR1C1 = "=INDIRECT(""价格表!D""&ROUNDDOWN(ROW()/14,0)*4+4)"
Range("T11").FormulaR1C1 = "=INDIRECT(""价格表!E""&ROUNDDOWN(ROW()/14,0)*4+4)"
Range("W8").FormulaR1C1 = "=INDIRECT(""价格表!F""&ROUNDDOWN(ROW()/14,0)*4+4)"

Range("AB7").FormulaR1C1 = "=INDIRECT(""价格表!A""&ROUNDDOWN(ROW()/14,0)*4+5)"
Range("AB8").FormulaR1C1 = "=INDIRECT(""价格表!B""&ROUNDDOWN(ROW()/14,0)*4+5)"
Range("AB9").FormulaR1C1 = "=INDIRECT(""价格表!C""&ROUNDDOWN(ROW()/14,0)*4+5)"
Range("AB10").FormulaR1C1 = "=INDIRECT(""价格表!D""&ROUNDDOWN(ROW()/14,0)*4+5)"
Range("AB11").FormulaR1C1 = "=INDIRECT(""价格表!E""&ROUNDDOWN(ROW()/14,0)*4+5)"
Range("AE8").FormulaR1C1 = "=INDIRECT(""价格表!F""&ROUNDDOWN(ROW()/14,0)*4+5)"

End Sub
Sub 生成表()
On Error Resume Next
Sheets.Add after:=Sheets(Sheets.Count)
ActiveSheet.Name = "价格表"
Sheets("价格表").Range("A1:F1") = Array("品名", "型号", "产地", "规格", "单位", "价格")
'cells(2,1) = "=IF(LEFT(B2)="J","快速燃气热水器",IF(MID(B2,2,1)="R","家用净水机",IF(LEFT(B2)="H","空气源热水器",IF(OR(LEFT(B2)="C",LEFT(B2)="E"),"壁挂电热水器",IF(LEFT(B2,4)="ACWP","中央净水机",IF(LEFT(B2,3)="RSE","软水机","*"))))))"
Sheets.Add after:=Sheets(Sheets.Count)
ActiveSheet.Name = "标签表"
Sheets("标签表").Active
End Sub
Sub 插入logo()
Set p = ActiveSheet.Pictures.Insert(ThisWorkbook.Path & "\aologo.png")
 Range("E4").Activate
        With p
            .Top = ActiveCell.Top + 2
            .Left = ActiveCell.Left
        End With
Debug.Print TypeName(p)
End Sub
Sub 清除()
Cells.Select
Selection.Delete
ActiveSheet.Shapes.SelectAll
Selection.Delete
End Sub

Sub 价格标签生成()
清除
调行列距
填充颜色
输入文字
合并单元格
设置字体
生成排版
End Sub
Sub 打印()
Columns("A:A").Select
Selection.ColumnWidth = 0.77
Rows("1:2").Select
Selection.RowHeight = 0
On Error Resume Next
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    ActiveWindow.View = xlPageBreakPreview
    ym = 0
    For pr = 16 To 1000 Step 14
    ym = ym + 1
    Set ActiveSheet.HPageBreaks(ym).Location = Range("A" & pr)
    Next pr
End Sub
