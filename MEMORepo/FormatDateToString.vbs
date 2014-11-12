'******************************************************************************
' 関数名： FormatYYYYMMDD
' 機　能： 引数の変数をYYYYMMDDで整形する
' 引　数： ARG1 - 日付型の変数
' 戻り値： フォーマット後の文字列（String）
' 補　足： 
'******************************************************************************
Private Function FormatYYYYMMDD(psDate)
  FormatYYYYMMDD = _
       Right("0000" & Year(psDate) , 4 ) & _
       Right("0" & Month(psDate) , 2) & _
       Right("0" & Day(psDate) , 2) 
End Function

'******************************************************************************
' 関数名： FormatHHMMSS
' 機　能： 引数の変数をHHMMSSで整形する
' 引　数： ARG1 - 日付型の変数
' 戻り値： フォーマット後の文字列（String）
' 補　足： 
'******************************************************************************
Private Function FormatHHMMSS(psDate)
  FormatHHMMSS = _
       Right("0" & Hour(psDate) , 2) & _
       Right("0" & Minute(psDate) , 2) & _
       Right("0" & Second(psDate) , 2) 
End Function
