Option Strict Off
Option Explicit On
Imports CoreLib
Imports CoreNS

Module std<PROGRAM_ID>
	' ****************************************************************************
	' ŠT—v  : <FORM_NAME>
	' ****************************************************************************
	'

	Public dbCmdCMM As CoreLib.ADODB.Command
	Public dbRecCMM As CoreLib.ADODB.Recordset

	Public dbCmdINS As CoreLib.ADODB.Command
	Public dbCmdUPD As CoreLib.ADODB.Command

	Public Sub Main()
		Dim pDebug As Boolean

#If DEBUG Then
		pDebug = True
#Else
        pDebug = False
#End If

		Try
			Using mainform As New frm<PROGRAM_ID>()
				MainLogic("<FORM_NAME>", mainform, pDebug)
			End Using
		Finally
			If Not dbCmdCMM Is Nothing Then dbCmdCMM.Dispose()
			If Not dbRecCMM Is Nothing Then dbRecCMM.Close()

			If Not dbCmdINS Is Nothing Then dbCmdINS.Dispose()
			If Not dbCmdUPD Is Nothing Then dbCmdUPD.Dispose()

			dbCmdCMM = Nothing
			dbRecCMM = Nothing

			dbCmdINS = Nothing
			dbCmdUPD = Nothing
		End Try

	End Sub

End Module