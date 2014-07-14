
Option Explicit
'*
'* TSQLAnalyzer�̃I�u�W�F�N�g���쐬�E�擾����.
'* @return TSQLAnalyzer
'*
Function getTSQLAnalyzer()
	set getTSQLAnalyzer = new TSQLAnalyzer
End Function


'! @class SQLAnalyzer
'!
class TSQLAnalyzer
	Private m_dicObjects	'* �I�u�W�F�N�g�̉\���̂�����̂𔲂��o���B

	'*
	'* SQL���̉�͂��s��
	'* @param[in] sSQL SQL
	'* @retrun True ���� False:���s
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
	
	'* �I�u�W�F�N�g�̉\���̂��镶�����擾
	Public Function GetObjects
		Set GetObjects = m_dicObjects
	End Function
	
	'* �s����͂���
	'* @param[in] �s
	Private Function AnalyzeLine( Byval sLine )
		Dim objTerm
		Set objTerm = SplitRegExp( sLine, " |\t|\,|\.|\(|\)|\[|\]|""|'" )
		Dim i
		For i = 0 To objTerm.Count - 1
			Call AnalyzeTerm( objTerm.Item(i) )
		Next
		Set objTerm = Nothing
	End Function
	
	'* ��������͂���B
	'* �I�u�W�F�N�g�ɂȂ肻���ȕ����𔲂��o���Ă���B
	'* @param[in] �����Ώۂ̕���
	Private Function AnalyzeTerm( Byval sTerm )
		' ���߂�Ȃ����A�蔲���Ă܂��B�����B
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
	
	'* �I�u�W�F�N�g���̌��ɒǉ��ς݂�����ׂ�
	'* @param[in] ���O
	'* @return Ture �ǉ��ς�
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
	
	'* ���K�\���ɂ��Split
	'* @param[in] sText  �����Ώ�
	'* @param[in] �p�^�[��
	'* @return    �A�z�z��ɂ�錋�� Dictionary
	Private Function SplitRegExp( Byval sText , Byval sPattern )
		Dim objDictRet
		Set objDictRet = WScript.CreateObject("Scripting.Dictionary")
		
		Dim objRegExp
		Set objRegExp =New RegExp
		objRegExp.Pattern = sPattern
		objRegExp.IgnoreCase = True
		objRegExp.Global = True

		'// ���ʂ̕�������쐬
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

