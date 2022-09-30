Sub savebysheetname()

myDateTime = Now()
myDateYear = Right(Format(Year(myDateTime), "@"), 2)
myDateMonth = Month(myDateTime)
myDateDay = Day(myDateTime)
myDateHour = Hour(myDateTime)
myDateMinute = Minute(myDateTime)
myDateSecond = Second(myDateTime)

myDateTimeText = myDateYear & "_" & myDateMonth & "_" & myDateDay & "_" & myDateHour & "_" & myDateMinute & "_" & myDateSecond

'원하는 폴더
myRootFolder = FolderDialog()

If Dir(myRootFolder & "\" & myDateTimeText, vbDirectory) = "" Then
       MkDir (myRootFolder & "\" & myDateTimeText)
   Else
       MsgBox "이미 해당 경로에 폴더가 있습니다."
End If


Dim twb As Workbook

Set twb = ActiveWorkbook
shCount = twb.Worksheets.Count
twb.Save

Dim wb As Workbook

Application.ScreenUpdating = False

For i = 1 To shCount
  
    myfilename = myRootFolder & "\" & myDateTimeText & "\" & myDateTimeText & "_" & twb.Worksheets(i).Name & ".xlsx"

    Set wb = Application.Workbooks.Add
        
    twb.Worksheets(i).Copy before:=wb.Worksheets(1)
    wb.SaveAs (myfilename)
    wb.Close

Next i

Application.ScreenUpdating = True

End Sub



Function FolderDialog(Optional Title As String = "폴더를 선택하세요", Optional InitialFolder As String = "", _
Optional InitialView As MsoFileDialogView = msoFileDialogViewList)
 
Dim sFolder As String
Dim Selected As Long
 
With Application.FileDialog(msoFileDialogFolderPicker)
.Title = Title
.InitialView = InitialView
.InitialFileName = InitialFolder
Selected = .Show
 
    If Selected = -1 Then
        sFolder = .SelectedItems(1)
        FolderDialog = sFolder
    Else
        MsgBox "선택된 폴더가 없어 프로그램을 종료합니다."
        End
    End If
End With
 
End Function