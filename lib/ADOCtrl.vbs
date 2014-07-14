Option Explicit

'*
'* ADOCtrl�̃I�u�W�F�N�g���쐬�E�擾����.
'* @return ADOCtrl
'*
Function getADOCtrl()
	set getADOCtrl = new ADOCtrl
End Function

'! @class ADOCtrl
'!
class ADOCtrl
	Private m_objADOCnn
	
	'* DB�ɐڑ�
	'* @param[in] host �z�X�g��
	'* @param[in] db   �f�[�^�x�[�X��
	'* @param[in] user ���[�U��
	'* @param[in] pass �p�X���[�h
	'*
	Public Function ConnectSqlServer(Byval host, Byval db, Byval user, Byval pass )
		Set m_objADOCnn = CreateObject("ADODB.Connection")
		
		Dim cnn
		
		cnn = "PROVIDER=SQLOLEDB" & _
              ";SERVER=" & host & _
              ";DATABASE=" & db & _
              ";UID=" & user & _ 
              ";PWD=" & pass & ";"
		
		Call m_objADOCnn.Open( cnn )
		If m_objADOCnn.Errors.Count = 0 Then
			ConnectSqlServer = True
		Else
			ConnectSqlServer = False
		End If
	End Function
	
	
	'*
	'* DB�̐ڑ���ؒf����B
	'*
	Public Sub Close
		If Not m_objADOCnn Is Nothing Then
			m_objADOCnn.Close
			Set m_objADOCnn = Nothing
		End If
	End Sub
	
	'* 
	'* SQL�̎��s
	'* @param[in]  SQL��
	'* @param[out] �o�͌���
	'*
	Public Function ExcuteSQL( ByVal sSQL, ByRef outRet )
		Dim objRS
		ExcuteSQL = False
		
		Set objRS = CreateObject("ADODB.Recordset")
		Call objRS.Open( sSQL, m_objADOCnn, 0, 1, 1 )
		If m_objADOCnn.Errors.Count > 0 Then
			Set objRS = NoThing
			Exit Function
		End If

		If objRS.EOF Then
			Set objRS = NoThing
			Exit Function
		End If
		
		outRet = objRS.GetRows()
		objRS.Close
		
		Set objRS = NoThing
		ExcuteSQL = True
		
	End Function
End Class