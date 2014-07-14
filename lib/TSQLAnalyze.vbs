
Option Explicit
'*
'* TSQLAnalyzerのオブジェクトを作成・取得する.
'* @return TSQLAnalyzer
'*
Function getTSQLAnalyzer()
	set getTSQLAnalyzer = new TSQLAnalyzer
End Function


'! @class SQLAnalyzer
'!
class TSQLAnalyzer
	Private m_dicObjects	'* オブジェクトの可能性のあるものを抜き出す。

	'*
	'* SQL文の解析を行う
	'* @param[in] sSQL SQL
	'* @retrun True 成功 False:失敗
	'*
	Public Function AnalyzeTSQL( Byval sSQL )
		Set m_dicObjects = Nothing
		Set m_dicObjects = WScript.CreateObject("Scripting.Dictionary")
		Dim objLine
		Set objLine = SplitRegExp( sSQL, "\r|\n" )
		Dim i
		Dim sLine
		Dim iFind
		Dim iFindEnd
		Dim bCommentOut

		bCommentOut = False
		For i = 0 To objLine.Count - 1
			sLine = objLine.Item(i)
			iFind = InStr( sLine, "--" )
			IF iFind > 0 Then
				sLine = Left( sLine, iFind - 1 )
			End If
			iFind = InStr( sLine, "/*" )
			iFindEnd = InStr( sLine, "*/" )
			If iFind > 0 And iFindEnd > 0 Then
				sLine = Right( sLine, Len(sLine) - iFindEnd -1 )
			ElseIF iFind > 0 And iFindEnd = 0 Then
				sLine = Left( sLine, iFind -1  )
				bCommentOut = True
			ElseIF iFindEnd > 0 And iFind = 0 Then
				sLine = Right( sLine, Len(sLine) - iFindEnd -1 )
				bCommentOut = False
			End If

			If Not bCommentOut Then
				Call AnalyzeLine( sLine )
			End If
		Next
		Set objLine = Nothing
	End Function
	
	'* オブジェクトの可能性のある文字を取得
	Public Function GetObjects
		Set GetObjects = m_dicObjects
	End Function
	
	'* 行を解析する
	'* @param[in] 行
	Private Function AnalyzeLine( Byval sLine )
		Dim objTerm
		Set objTerm = SplitRegExp( sLine, " |\t|\,|\.|\(|\)|\[|\]|""|'" )
		Dim i
		For i = 0 To objTerm.Count - 1
			Call AnalyzeTerm( objTerm.Item(i) )
		Next
		Set objTerm = Nothing
	End Function
	
	'* 文字を解析する。
	'* オブジェクトになりそうな文字を抜き出している。
	'* @param[in] 検査対象の文字
	Private Function AnalyzeTerm( Byval sTerm )
		' ごめんなさい、手抜いてます。ここ。
		If IsMatch( sTerm , "^max$|^min$|^sum$|^group$|^begin$|^end$|^;$|^select$|^delete$|^where$|^insert$|^into$|^order$|^from$|^by$|^as$|^nocount$|^on$|^off$|^create$|^update$|^values$|^numeric$|^set$|^table$|^proc$|^varchar$|^datetime$|^exec$|^\=$|^\+$|^and$|^or$|^@") Then
			Exit Function
		End If
		If IsNumeric( sTerm ) Then
			Exit Function
		End If
		If IsExistObjects(sTerm) Then
			Exit Function
		End If

		m_dicObjects.Item(m_dicObjects.Count) = sTerm
	End Function
	
	'* オブジェクト名の候補に追加済みかしらべる
	'* @param[in] 名前
	'* @return Ture 追加済み
	Private Function IsExistObjects( Byval sTerm )
		IsExistObjects = True
		Dim i
		For i = 0 To m_dicObjects.Count - 1
			If  UCase(m_dicObjects.Item(i)) =  UCase(sTerm) Then
				Exit Function
			End If
		Next
		IsExistObjects = False
	End Function
	'*
	Private Function IsMatch( Byval sText , Byval sPattern )
		Dim objRegExp
		Set objRegExp =New RegExp
		objRegExp.Pattern = sPattern
		objRegExp.IgnoreCase = True
		objRegExp.Global = True
		IsMatch = objRegExp.Test(sText)
		Set objRegExp = Nothing
	End Function
	
	'* 正規表現によるSplit
	'* @param[in] sText  検査対象
	'* @param[in] パターン
	'* @return    連想配列による結果 Dictionary
	Private Function SplitRegExp( Byval sText , Byval sPattern )
		Dim objDictRet
		Set objDictRet = WScript.CreateObject("Scripting.Dictionary")
		
		Dim objRegExp
		Set objRegExp =New RegExp
		objRegExp.Pattern = sPattern
		objRegExp.IgnoreCase = True
		objRegExp.Global = True

		'// 結果の文字列を作成
		Dim Match
		Dim Matches
		Set Matches = objRegExp.Execute(sText) 
		Dim iStart
		Dim iCnt
		iStart = 1
		iCnt = 0
		For Each Match in Matches
			IF iStart <=  Match.FirstIndex THEN
			    objDictRet.Item(iCnt) =  Mid( sText,iStart,Match.FirstIndex-iStart + 1 )
			    iCnt = iCnt + 1
		    End If
		    iStart = Match.FirstIndex + Len(Match.Value) + 1
		Next
	    objDictRet.Item(iCnt) = Mid( sText,iStart)
		Set objRegExp = Nothing
		Set SplitRegExp = objDictRet
	End Function

	
End Class

