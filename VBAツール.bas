Attribute VB_Name = "Module1"
Sub スポット用処理()
    
    Dim i As Integer, cnt As Integer, buf As Integer, num As Integer, time As String, dum As Integer
    Dim sh_s1 As Worksheet, sh_s2 As Worksheet
    Dim dialogueResult As String
    Dim myPath As String
    Dim csv_p As String, kikaku As String, receiveDate As String
    Dim targetRange As Range
    Dim wb As Workbook
    Dim fso As Object
    
    Set sh_s1 = Sheets("登録表(スポットセンター用)")
    Set sh_s2 = Sheets("③出力フォーマット(新規センター用)")
    
    sh_s2.Range("A5:A" & Rows.Count).EntireRow.Clear
    
    i = sh_s1.Cells(Rows.Count, 4).End(xlUp).Row        '登録表シートの対応件数の最終行
    
    buf = 5                                             'sh_s2シートの5行目から転記を始めるための変数
    
    num = 1                                             '受注Noの転記で使う変数
    
    sh_s2.Range("C:C").NumberFormatLocal = "@"
    
    For cnt = 7 To i
      time = Format("00:00:00", "hh:mm:ss")
      dum = 1
      
      For cnt2 = 1 To sh_s1.Cells(cnt, 4)
       

       If sh_s1.Cells(cnt, 2) = "受注" Or sh_s1.Cells(cnt, 2) = "引上" Then
         sh_s2.Cells(buf, 1) = num                                              '受注Noを転記
         num = num + 1
       End If
    
       sh_s2.Cells(buf, 2) = sh_s1.Cells(2, 3).Text                             'データ登録日を転記
                      
       sh_s2.Cells(buf, 3) = time                                               'データ登録時間を転記
       time = Format(DateAdd("s", 1, time), "hh:mm:ss")
       
       With sh_s2
          .Cells(buf, 4) = sh_s1.Cells(cnt, 3)                           'FDを転記
          .Cells(buf, 5) = sh_s1.Cells(3, 3).Value                            '企画を転記
          .Cells(buf, 6) = sh_s1.Cells(4, 3).Value                            'センター名を転記
          .Cells(buf, 7).Value = "通常応答"                                   '対応区分を転記
          .Cells(buf, 8).Value = "マルチ"                                     '席区分を転記
       End With
       
       sh_s2.Cells(buf, 11).Value = "dummy" + CStr(dum)                         'OPIDを転記
       dum = dum + 1
       
       If sh_s1.Cells(cnt, 2) = "受注" Or sh_s1.Cells(cnt, 2) = "引上" Then
         sh_s2.Cells(buf, 12).Value = "不明枠ID"                                '枠IDを転記
       End If
       
       dialogueResult = sh_s1.Cells(cnt, 2)
       
       Select Case dialogueResult
          Case "引上"
              sh_s2.Cells(buf, 13).Value = "受注"
          Case "新規問合せ"
              sh_s2.Cells(buf, 13).Value = "新規顧客"                           '対話結果1を転記
          Case "既存問合せ"
              sh_s2.Cells(buf, 13).Value = "既存顧客"
          Case Else
              sh_s2.Cells(buf, 13).Value = sh_s1.Cells(cnt, 2)
       End Select
       
       Select Case dialogueResult
          Case "受注"
            sh_s2.Cells(buf, 14).Value = "非引上"
          Case "引上"
            sh_s2.Cells(buf, 14).Value = "引上"
          Case "新規問合せ"
            sh_s2.Cells(buf, 14).Value = "購入前その他"                         '対話結果２を転記
          Case "既存問合せ"
            sh_s2.Cells(buf, 14).Value = "購入前その他"
          Case "ロス"
            sh_s2.Cells(buf, 14).Value = "非接続"
          Case "阻害"
            sh_s2.Cells(buf, 14).Value = "直切れ"
          Case "販促NG"
            sh_s2.Cells(buf, 14).Value = "その他"
       End Select
          
       Select Case dialogueResult
          Case "受注"
            sh_s2.Cells(buf, 15).Value = "その他"
          Case "引上"
            sh_s2.Cells(buf, 15).Value = "-"                                     '対話結果３を転記
          Case "新規問合せ"
            sh_s2.Cells(buf, 15).Value = "その他"
       End Select
         
       With sh_s2
          .Cells(buf, 16).Value = "0"                                          'エスカレーションを転記
          .Cells(buf, 17).Value = "0"                                          '資料請求を転記
          .Cells(buf, 18).Value = "0"                                          'FD案内を転記
          .Cells(buf, 19).Value = "0"                                          '店舗案内を転記
       End With
       
       If sh_s1.Cells(cnt, 2) = "受注" Then
         With sh_s2
            .Cells(buf, 20).Value = "不明"                                     '性別を転記
            .Cells(buf, 22).Value = "不明"                                     '年代を転記
            .Cells(buf, 25).Value = "297_91"                                   '商品ID1を転記
            .Cells(buf, 26).Value = "受注_簡易履歴登録用_単品"                 '受注商品1を転記
            .Cells(buf, 27).Value = "1"                                        '数量1を転記
          End With
       ElseIf sh_s1.Cells(cnt, 2) = "引上" Then
         With sh_s2
            .Cells(buf, 20).Value = "不明"
            .Cells(buf, 22).Value = "不明"
            .Cells(buf, 25).Value = "297_92"
            .Cells(buf, 26).Value = "受注_簡易履歴登録用_定期"
            .Cells(buf, 27).Value = "1"
         End With
       End If
       
       buf = buf + 1
       
      Next
      
    Next
 
    'ファイル名にするセルを変数へ格納
    csv_p = sh_s1.Range("F3")                                                                '保存したいパス
    kikaku = sh_s1.Cells(3, 3)                                                               '企画名
    receiveDate = Format(sh_s1.Cells(2, 3), "yyyymmdd")                                      '受電日

    myPath = csv_p & "\" & "op_" & kikaku & "_" & receiveDate & ".csv"                             'ファイルパス
       
    Set targetRange = sh_s2.Range("A4:BB" & sh_s2.Cells(Rows.Count, 2).End(xlUp).Row)        'CSVファイルへ出力する範囲を指定

    Set wb = Workbooks.Add                                                                   '新規ブックを作成
    
    targetRange.Copy wb.Worksheets(1).Range("A1")                                            'CSVファイルへ出力する範囲を新規ブックへコピー

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(myPath) Then                                                           '出力するCSVファイルが既に存在する場合は削除
        fso.deleteFile (myPath)
    End If
    
    wb.SaveAs Filename:=myPath, FileFormat:=xlCSV, Local:=True                               '新規ブックをCSVファイルとして出力
    
    wb.Close SaveChanges:=False                                                              '新規ブックを保存せずに閉じる

    Set fso = Nothing                                                                        '後片付け

    MsgBox "保存が完了しました"
    
End Sub
