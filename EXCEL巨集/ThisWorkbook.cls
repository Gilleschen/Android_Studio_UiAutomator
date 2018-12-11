VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo skipp
    Application.CommandBars("功能說明").Delete
    
    
skipp:
    ActiveWorkbook.Save
End Sub

Private Sub Workbook_Open()
Dim myNewBar As CommandBar           '宣告工具列物件
    Dim ScriptCreator, RunScript '宣告工具列按鈕物件

    Set myNewBar = Application.CommandBars.Add      '新增一個工具列
    On Error GoTo skipp

    myNewBar.Name = "功能說明" '工具列命名
    myNewBar.Position = msoBarTop
   

    With myNewBar
        Set RunScript = .Controls.Add(msoControlButton)
    
        With RunScript

            '.Style = msoButtonCaption   '只顯示文字 底下這3種型式選一種
            '.Style = msoButtonIcon       '只顯示小圖示
            '同時顯示文字和小圖示
            .Style = msoButtonIconAndCaption 'msoButtonIconAndCaption 'msoButtonIconAndWrapCaption 'msoButtonIconAndWrapCaptionBelow 'msoButtonIconAndCaptionBelow 'msoButtonIconAndCaption ' msoCombolabel
            .BeginGroup = True
            .Caption = "執行腳本"                  '顯示在工具列上的按鈕文字
            .TooltipText = "執行腳本" ' "TUTK QA團隊製作"  '滑鼠移過去時，所顯示的提示文字
            .FaceId = 156  '1778                       '小圖示
            .Tag = "Jhih Chen"
            .OnAction = "RunScript"     '設定按下此鍵時所要執行的巨集
        End With
        
    
        'FaceId :http://juve.tistory.com/archive/20131205

        .Position = msoBarTop             '工具列擺放在上層
        .Visible = True
    End With
    'Call loadingAPPAndDevice
skipp:
End Sub
