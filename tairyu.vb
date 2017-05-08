Option Explicit
'------考え方と設計------------------------------------------------>
'回収予定日　今日の日付　入金予定日 から滞留月数を返す関数をつくる
'大雑把な計算式　（滞留月数） ＝（今日の月）ー（回収予定日の月）
'具体例）　回収予定日2017年04月10日 今日の日付　2017年05月01日 入金予定日　2017年07月10日
'　滞留月数　＝　今日の月である５　マイナス　回収予定日の月である４　→　滞留月数　＝　１　になり滞留月数は１超
'考慮する事項（滞留を定義する上で考える事項）
'：年数をまたぐ場合月だけでは計算できない
'例）2017年01月　と　2016年12月　で計算すると　マイナスになる
'：滞留の定義上入金予定日が今日の月と同じ場合は滞留ではない
'：入金予定日がきまってない場合があるので予定日を構築する必要がある
'：今日が月末で回収予定日も月末だと滞留となる
'：今日の日付が回収予定日の日付よりも大きい場合は同じ月でも滞留になる
'--------------------------------------------------------------------->
'/**
'*下記関数パラメータに関しての記述
'*滞留を返す関数です
'*2016sato-yoshitaka@akt-g.jp
'*WorkSheetFunction "A4" = tairyu("A1","A2","A3")
'*@pram {Date(Variant)} enddate　回収予定日、セル指定でも直接指定でもOK
'*@pram {Date(Variant)} pivdate 軸になる日付、セル指定でも直接指定でもOK
'*@pram {Date(Variant)} scdate 入金予定日、セル指定でも直接指定でもOK
'*@return {Number or String} tairyu 滞留月数かエラーメッセージ
'*/

Public Function tairyu(ByVal enddate As Variant, ByVal pivdate As Variant, ByVal scdate As Variant) As Variant
    
    Dim result  As Variant '返す結果用変数をvariant型で定義して数値も文字列も返せるようにする
   
   'セル指定でも直接指定でもOKの為引数がRange型かどうか判定して
   
   '引数がセルの場合はセル内容を取り出し再代入
    
    If TypeName(enddate) = "Range" Then
        enddate = enddate.Value
        pivdate = pivdate.Value
        scdate = scdate.Value
    End If
    
    '回収予定日が日付かどうか判定する為の分岐
    '日付以外はエラーメッセージ表示
    If Not IsDate(enddate) Then
        result = "回収予定日は日付を指定してください"
        GoTo endResult
    End If
    
    '軸になる日付が日付かどうか判定する為の分岐
    '日付以外はエラーメッセージ表示
    If Not IsDate(pivdate) Then
    result = "軸になる日付の指定は日付を指定してください"
        GoTo endResult
    End If
    
    '入金予定日が日付もしくは空白を判定
    If IsDate(scdate) = False And scdate <> "" Then
        result = "入金予定日は日付もしくは空白を指定してください"
        GoTo endResult
    End If

'エラーメッセージ表示部分が終わったので滞留月数にとりあえず０を代入
'各数字データを各変数へ代入
    result = 0  '結果用変数を初期化
    Dim pivflag As Boolean '軸になる日付用フラグ
    pivflag = False  'フラグを初期化
    Dim endflag As Boolean '回収予定日用フラグ
    endflag = False 'フラグ
     Dim YY As Long  '軸になる日付（今日）の年用の変数
    YY = Year(pivdate)  '軸になる日付（今日）の年（20170401→2017）
    Dim frYY As Long  '回収予定日の年用の変数
    frYY = Year(enddate) '回収予定日の年
    Dim MM As Long  '軸になる日付（今日）の月用の変数
    MM = Month(pivdate) '軸になる日付（今日）の月（20170401→ 4）
    Dim frMM As Long '回収予定日の月用の変数
    frMM = Month(enddate) '回収予定日の月
 
 '考慮事項：年が違っている場合の処理
 'select分で年数比較してそれぞれで分岐　年数が違う場合は月に加算します
 '例）2017年１月と2016年１２月の場合　月　MMは１＋１２＝１３　frMM = １２　になり　１超になる
    Select Case YY - frYY
        Case 1
            MM = MM + 12
        Case -1
            frMM = frMM + 12
    End Select
    
'考慮事項：入金予定日が決まっていない場合の処理
'仮の入金予定日を構築
'軸になる年の翌月でかつ回収予定日の同日
'例）今日が20170401回収予定日が　20170531の場合　仮の入金予定日は20170531になる
    If scdate = "" Then '入金予定日が空白の場合
        scdate = DateSerial(YY, MM + 1, Day(enddate))
    End If

'今日の日付＝入金予定日の場合は滞留にならないので０を返す
    If Month(pivdate) = Month(scdate) And Day(pivdate) < Day(scdate) Then
        result = 0
        GoTo endResult
    End If

'今日の日付の月から回収予定日の月を引いて０より小さい場合滞留にはならないので０を返す
    If (MM - frMM) < 0 Then
        result = 0
        GoTo endResult
    End If

'今日の日付の月から回収予定の月を引いて０より大きい場合は滞留になるので滞留用変数に引いた月数を足す
    If (MM - frMM) > 0 Then
        result = result + (MM - frMM)
    End If

'考慮事項
'軸になる日付も回収予定日も月末の場合は滞留のため月末を取得して処理
    Dim endday As String '回収予定日の月末日用変数
    endday = DateSerial(Year(enddate), Month(enddate) + 1, 0)
    Dim pivday As String '軸になる日付の月末日用変数
    pivday = DateSerial(Year(pivdate), Month(pivdate) + 1, 0)

'回収予定日が月末日かどうか判定して月末日だった場合回収予定日用フラグをtrueにする
    If Day(endday) = Day(enddate) Then
        endflag = True
    End If
 
 '軸になる日付が月末日かどうか判定して月末日だった場合軸になる日付用フラグをtrueにする
    If Day(pivday) = Day(pivdate) Then
        pivflag = True
    End If
 
 '回収予定日用フラグと軸になる日付用フラグが両方ともtrueの場合は滞留になるので滞留月数に１を足す
    If endflag And pivflag Then
        result = result + 1

'考慮事項
'軸になる日付の月と回収予定日の月が一緒の場合軸になる日付の日が回収予定日の日よりも大きい場合滞留になる
'日付計算で分岐処理
    ElseIf MM - frMM >= 0 And Day(pivdate) - Day(enddate) >= 0 Then
        result = result + 1
    End If
endResult:
'滞留データのresultを滞留関数へ返す処理
tairyu = result
End Function
