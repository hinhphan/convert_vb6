Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPK9OSCM004
	Inherits System.Windows.Forms.Form
	' ****************************************************************************
	' 概要  : 評価項目マスター保守 （ＰＫ９ＯＳＣＭ００４）
	' 　　  :
	' 履歴  : 2010.02.12  REV.0007  植田  医学部のＳＰの時に、「はい、中間、いいえ」を「３，２，１，０」に変更
	' 　　  : 2009.06.03  REV.0006  植田  薬学部対応
	' 　　  : 2009.03.04  REV.0005  保延  事前登録データからのコピー機能を追加。
	'       : 2006.02.20  REV.0004  植田  解答欄の４つ目に対応
	' 　　  : 2006.01.16  REV.0003  吉田  問題番号の項目追加。
	' 　　  : 2006.01.06  REV.0002  吉田  模擬患者の時、評価項目を最大２６行に変更。
	' 　　  : 2005.12.12  REV.0001  吉田  新規作成。
	' ****************************************************************************
	'
	
	Private mMsgText As String
	Private mMBOXTitle As String
	Private mLostFocusCheck As Boolean
	Private mInputMode As String
	
	Private mSW_CellFocusEvent As Boolean
	Private mSW_CellKeyPress As Boolean
	Private mCellClickCol As Integer
	Private mSW_EnterKeyPress As Boolean
	Private mActive_PGrid As Short
	
	'''グリッド行切り替え
	''Private Const mGRGYOH              As Long = 7      '模擬患者
	''Private Const mGRGYOM              As Long = 15     '
	
	'画面入力項目
	Private mNENDO As Integer
	Private mCDGK As Integer
	Private mKBSK As Integer
	Private mYEAR As Integer
	Private mKBHY As Integer
	Private mCDST As Integer
	
	'設定済み情報-----------------------
	Private mSUKJ As Integer
	Private mSUST As Integer
	Private mSUJU As Integer
	Private mSUBN As Integer
	'
	Private mSUKJM As Integer
	Private mSUSTM As Integer
	''Private mSUJUM               As Long
	
	Private mTIME As String
	
	
	' *******************************************************************************
	' 概要    : 数字セルの有効範囲のチェックを行う。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : Row, I, Long, 列。
	'         : Col, I, Long, 行。
	'         : 戻り値, O, Integer, True=正常。
	'         :                    False=エラー。
	'         :
	' 説明    :
	'         :
	' 履歴    : 1998.12.18  REV.0001  深津  新規作成。
	' *******************************************************************************
	'
	Private Function CellCheck_Numeric(ByVal pPGrid As AxPGRIDLib.AxPerfectGrid, ByVal Row As Integer, ByVal Col As Integer) As Short
		
		Dim pRetVal As Short
		Dim pBytes As Integer
		
		CellCheck_Numeric = False
		
		pPGrid.set_CellText(Row, Col, Trim(pPGrid.get_CellText(Row, Col)))
		
		'UPGRADE_ISSUE: 定数 vbFromUnicode はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"' をクリックしてください。
		'UPGRADE_ISSUE: LenB 関数はサポートされません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"' をクリックしてください。
		pBytes = LenB(StrConv(pPGrid.get_CellText(Row, Col), vbFromUnicode))
		
		If (pBytes <> 0) And (Not pPGrid.get_CellIsValue(Row, Col)) Then
			pPGrid.set_CellText(Row, Col, "")
			Exit Function
		End If
		
		If (pPGrid.get_CellValue(Row, Col) < pPGrid.get_ColMinValue(Col)) Or (pPGrid.get_CellValue(Row, Col) > pPGrid.get_ColMaxValue(Col)) Then
			Exit Function
		End If
		
		CellCheck_Numeric = True
		
	End Function
	
	Private Sub cboCDGK_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCDGK.ClickEvent
		
		Dim pKBMEK As Integer
		
		''    If Not mLostFocusCheck Then
		''        Exit Sub
		''    End If
		''
		''    If Not OSC_CMM2_Read(cboCDGK.ItemData(cboCDGK.ListIndex), inpYEAR.Value, pKBMEK) Then
		''         optKBKJ(1).Enabled = False
		''         optKBKJ(0).Value = True
		''    Else
		''        If pKBMEK = 0 Then
		''            optKBKJ(1).Enabled = False
		''            optKBKJ(0).Value = True
		''        End If
		''    End If
		
	End Sub
	
	Private Sub cboCDGK_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCDGK.Enter
		
		mLostFocusCheck = True
		
	End Sub
	
	Private Sub cboCDGK_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCDGK.Leave
		
		Dim pRetVal As Short
		
		Select Case FocusMove_HEAD(cboCDGK)
			Case "H"
				If Not mLostFocusCheck Then
					Exit Sub
				End If
				
				mLostFocusCheck = False
				Exit Sub
				
			Case "I"
				If InputMode_Check() Then
					Exit Sub
				End If
		End Select
		
	End Sub
	
	'********************************************************************************
	' 概要    : ヘッダー部内でのフォーカスの移動方向を判断する。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : pCtrl, I, Control, コントロール。
	'         : 戻り値, O, String, ""=移動不可。
	'         :                   "H"=ヘッダー部内の移動。
	'         :                   "I"=アイテム部内への移動。
	'         :                   "C"=アイテム部内のコマンドボタンへの移動。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2005.12.12  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function FocusMove_HEAD(ByVal pCtrl As System.Windows.Forms.Control) As String
		
		FocusMove_HEAD = ""
		'
		' ピクチャーボックス／ラベルへの移動
		'
		If (TypeOf Me.ActiveControl Is System.Windows.Forms.PictureBox) Or (TypeOf Me.ActiveControl Is AxxLabelLib.AxxLabel) Then
			pCtrl.Focus()
			Exit Function
		End If
		'
		' コマンドボタンへの移動
		'
		If (TypeOf Me.ActiveControl Is AxxCBtnLib.AxxCmdBtn) Then
			If Me.ActiveControl Is cmdCDST Then
				FocusMove_HEAD = "C"
				Exit Function
			End If
		End If
		'
		' ツールバーへの移動
		'
		If (Me.ActiveControl Is Toolbar1 Or Me.ActiveControl Is mnuFILE Or Me.ActiveControl Is pCtrl) Then
			FocusMove_HEAD = "C"
			Exit Function
		End If
		
		If Me.ActiveControl Is cboCDGK Or Me.ActiveControl Is optKBSK(0) Or Me.ActiveControl Is optKBSK(1) Or Me.ActiveControl Is inpYEAR Or Me.ActiveControl Is inpCDST Or Me.ActiveControl Is optKBHY(0) Or Me.ActiveControl Is optKBHY(1) Then
			FocusMove_HEAD = "H"
		Else
			FocusMove_HEAD = "I"
		End If
		
	End Function
	
	' *******************************************************************************
	' 概要    : 画面を初期状態にする。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2005.12.12  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Public Sub Screen_Clear()
		
		Dim pRTN As Short
		
		picHEAD.Enabled = True
		picITEM.Enabled = True
		
		picGAIRYAKU.Visible = False
		picMONDAI.Visible = False
		
		Toolbar1.Items.Item("EXEC").Enabled = False
		Toolbar1.Items.Item("CANCEL").Enabled = False
		Toolbar1.Items.Item("PRINT").Enabled = True
		Toolbar1.Items.Item("EXIT").Enabled = True
		Toolbar1.Items.Item("ROWINSERT").Enabled = False
		Toolbar1.Items.Item("ROWDELETE").Enabled = False
		Toolbar1.Items.Item("COPY").Enabled = True
		
		mnuFILEItem(0).Enabled = False
		mnuFILEItem(1).Enabled = False
		mnuFILEItem(2).Enabled = False
		mnuFILEItem(5).Enabled = True
		mnuFILEItem(9).Enabled = True
		mnuEDITItem(0).Enabled = False
		mnuEDITItem(1).Enabled = False
		mnuEDITItem(9).Enabled = True
		
		Call GridClear()
		
		mLostFocusCheck = True
		mActive_PGrid = 0
		
	End Sub
	
	Private Function CDGP_Check(ByVal pPGrid As AxPGRIDLib.AxPerfectGrid, ByVal pROW As Integer, ByVal pCDGP As Integer) As Integer
		
		Dim pCDGP2 As Integer
		'Dim pNN     As String
		Dim pIX As Integer
		
		CDGP_Check = False
		
		If pROW = 0 Then
			If pCDGP <> 1 Then
				Exit Function
			Else
				CDGP_Check = True
				Exit Function
			End If
		End If
		
		For pIX = pROW - 1 To 0 Step -1
			If pPGrid.get_CellCheckedByName(pIX, "KBKR") = False Then '空行でない
				'前の行のＣＤＧＰ
				pCDGP2 = pPGrid.get_CellValueByName(pIX, "CDGP")
				Exit For
			End If
		Next pIX
		
		'前の行より小さい値は入力不可
		If pCDGP < pCDGP2 Then
			Exit Function
		End If
		
		CDGP_Check = True
		
	End Function
	
	Private Sub cmdCDST_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCDST.ClickEvent
		
		Dim pRetVal As Short
		Dim pSW_Select As Boolean
		Dim pINCODE As Integer
		Dim pIX As Short
		Dim pNAME As String
		Dim pKBSP As Integer
		
		With frmINQ_OSC_STM
			.OSC_STMNENDO = mNENDO
			.OSC_STMCDGK = cboCDGK.get_ItemData(cboCDGK.ListIndex)
			.OSC_STMKBSK = IIf(optKBSK(0).Value = True, 1, 2)
			.OSC_STMYEAR = inpYEAR.Value
			.OSC_STMCDST = inpCDST.Value
			.ShowDialog()
			
			pINCODE = .OSC_STMCDST
			pSW_Select = .SELECTION
		End With
		
		'UPGRADE_NOTE: オブジェクト frmINQ_OSC_STM をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		frmINQ_OSC_STM = Nothing
		
		inpCDST.Focus()
		System.Windows.Forms.Application.DoEvents()
		
		If pSW_Select Then
			inpCDST.Value = pINCODE
			
			If Not OSC_STM_READ(inpCDST.Value, pNAME, False, pKBSP) Then
				lblNMST.Text = pNAME
				Exit Sub
			Else
				lblNMST.Text = pNAME
			End If
			If pKBSP = 0 Then
				optKBHY(1).Enabled = False
				optKBHY(0).Value = True
			Else
				optKBHY(1).Enabled = True
			End If
			System.Windows.Forms.SendKeys.Send("{tab}")
		End If
		
	End Sub
	
	Private Sub cmdCLOSE_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCLOSE.ClickEvent
		Dim Index As Short = cmdCLOSE.GetIndex(eventSender)
		
		picGAIRYAKU.Visible = False
		picMONDAI.Visible = False
		
		
	End Sub
	'ADD 2009.03.04
	Private Sub cmdCOPY_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCOPY.ClickEvent
		
		Dim pRetVal As Short
		Dim pSW_Select As Boolean
		Dim pINCODE As Integer
		Dim pIX As Short
		Dim pNAME As String
		Dim pKBSP As Integer
		Dim pGYO As Short
		Dim pROW As Integer
		
		With frmINQ_OSC_TOROKU
			.OSC_STMNENDO = mNENDO
			.OSC_STMCDGK = cboCDGK.get_ItemData(cboCDGK.ListIndex)
			.OSC_STMNOTR = 0
			.OSC_STMCDST = mCDST
			.OSC_STMYEAR = mYEAR
			.OSC_STMSWKB = True
			.OSC_STMKBHY = mKBHY
			.ShowDialog()
			
			pINCODE = .OSC_STMNOTR
			pSW_Select = .SELECTION
		End With
		
		'UPGRADE_NOTE: オブジェクト frmINQ_OSC_TOROKU をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		frmINQ_OSC_TOROKU = Nothing
		
		PGrid.Focus()
		System.Windows.Forms.Application.DoEvents()
		
		If pSW_Select Then
			
			'グリッドのクリア
			''評価者区分によって、最大行を切り替える(評価)----------2006.01.06 add
			If mKBHY = 0 Then
				pGYO = 50
			ElseIf mKBHY = 1 Then 
				pGYO = 26
			End If
			
			PGrid.RefreshLater = True
			pROW = PGrid.Items
			PGrid.RemoveItems(0, pROW)
			PGrid.TextAtAddItem = ""
			PGrid.AddItems(0, pGYO)
			
			For pIX = 0 To (PGrid.Items - 1)
				PGrid.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
			Next pIX
			
			PGrid.RefreshLater = False
			'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
			PGrid.CtlRefresh()
			
			''評価者区分によって、最大行を切り替える(概評・問題点)----------
			If mKBHY = 0 Then
				pGYO = 7
			ElseIf mKBHY = 1 Then 
				pGYO = 15
			End If
			
			PGrid2.RefreshLater = True
			pROW = PGrid2.Items
			PGrid2.RemoveItems(0, pROW)
			PGrid2.TextAtAddItem = ""
			PGrid2.AddItems(0, pGYO)
			
			For pIX = 0 To (PGrid2.Items - 1)
				PGrid2.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
			Next pIX
			
			PGrid2.RefreshLater = False
			'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
			PGrid2.CtlRefresh()
			
			PGrid3.RefreshLater = True
			pROW = PGrid3.Items
			PGrid3.RemoveItems(0, pROW)
			PGrid3.TextAtAddItem = ""
			PGrid3.AddItems(0, pGYO)
			
			For pIX = 0 To (PGrid3.Items - 1)
				PGrid3.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
			Next pIX
			
			PGrid3.RefreshLater = False
			'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
			PGrid3.CtlRefresh()
			
			''評価者区分によって、タイトルを切り替える----------
			'2009.06.03 upd ---
			If bCDGB = 3 Then
				If mKBHY = 0 Then
					PGrid.set_CellTextByName(-1, "KB1", "")
					PGrid.set_CellTextByName(-1, "KB2", "")
					PGrid.set_CellTextByName(-1, "KB3", "はい")
					PGrid.set_CellTextByName(-1, "KB4", "いいえ")
				ElseIf mKBHY = 1 Then 
					PGrid.set_CellTextByName(-1, "KB1", "")
					PGrid.set_CellTextByName(-1, "KB2", "はい")
					PGrid.set_CellTextByName(-1, "KB3", "中間") '2006.12.21 update（とりあえず固定。メンテできるようにする）
					PGrid.set_CellTextByName(-1, "KB4", "いいえ")
				End If
			Else
				If mKBHY = 0 Then
					PGrid.set_CellTextByName(-1, "KB1", "３")
					PGrid.set_CellTextByName(-1, "KB2", "２")
					PGrid.set_CellTextByName(-1, "KB3", "１")
					PGrid.set_CellTextByName(-1, "KB4", "０")
				ElseIf mKBHY = 1 Then 
					'2010.02.12 del ---
					''PGrid.CellTextByName(-1, "KB1") = ""
					''PGrid.CellTextByName(-1, "KB2") = "はい"
					''PGrid.CellTextByName(-1, "KB3") = "中間"    '2006.12.21 update（とりあえず固定。メンテできるようにする）
					''PGrid.CellTextByName(-1, "KB4") = "いいえ"
					'2010.02.12 add ---
					PGrid.set_CellTextByName(-1, "KB1", "３")
					PGrid.set_CellTextByName(-1, "KB2", "２")
					PGrid.set_CellTextByName(-1, "KB3", "１")
					PGrid.set_CellTextByName(-1, "KB4", "０")
					'------------------
				End If
			End If
			'-----------------------------------------------
			
			Call OSC_HYOKAKMT_Read(pINCODE)
			Call OSC_HYOKAKMGT_READ(pINCODE)
			Call OSC_HYOKAKMMT_READ(pINCODE)
			
			System.Windows.Forms.SendKeys.Send("{tab}")
		End If
		
		
	End Sub
	
	Private Sub cmdDISP2_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDISP2.ClickEvent
		
		picGAIRYAKU.Visible = True
		picMONDAI.Visible = False
		
		
	End Sub
	
	Private Sub cmdDISP3_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDISP3.ClickEvent
		
		picGAIRYAKU.Visible = False
		picMONDAI.Visible = True
		
	End Sub
	
	Private Sub frmPK9OSCM004_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim pRetVal As Short
		Dim pMsgText As String
		Dim pIX As Short
		Dim pCHAR() As String
		Dim pNAME As String
		Dim pYMD As String
		
		'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		Me.Text = FormCaption
		mMBOXTitle = Me.Text
		
		mInputMode = "E"
		
		'' --------------------------------------------------------------
		If Not Qry_Create() Then
			Me.Close()
		End If
		
		If Not ResultSet_Initialize() Then
			Me.Close()
		End If
		
		'ユーザID、ユーザー、学部取得
		If VB.Command() <> "" Then
			pCHAR = Split(VB.Command(), "_")
			On Error Resume Next
			bIDUS = pCHAR(0)
			bNMUS = pCHAR(1)
			bCDGB = CInt(pCHAR(2))
			On Error GoTo 0
		End If
		
		'処理年度の取得
		pRetVal = OSC_CMM_Read()
		mNENDO = CMNENDO
		pYMD = CMNENDO & "/01/01"
		lblNENDO.Text = CMNENDO & "年度"
		
		Call Screen_Clear()
		Call Grid_Resize()
		Call Grid_Resize2()
		Call Grid_Resize3()
		
		''学部マスターをREADしてラベルにセットする
		pRetVal = GBM_xLabel_Set(1, bCDGB, lblNMGB)
		
		''学科をドロップダウンにセットする
		Call GKM_xDrop_Set2(1, bCDGB, False, False, cboCDGK)
		
		inpYEAR.Value = 0
		inpCDST.Value = 0
		lblNMST.Text = ""
		
		picGAIRYAKU.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(picITEM.Top) + 30)
		picGAIRYAKU.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(picITEM.Left) + 30)
		picGAIRYAKU.Width = picITEM.Width
		picMONDAI.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(picITEM.Top) + 30)
		picMONDAI.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(picITEM.Left) + 30)
		picMONDAI.Width = picITEM.Width
		
		'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	Private Sub frmPK9OSCM004_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
		'UPGRADE_WARNING: App プロパティ App.EXEName には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
		Call MenuLOGF_Insert(My.Application.Info.AssemblyName, FormCaption, bEXIT, bIDUS, bNMUS)
		
		' 更新用のコネクション解除
		On Error Resume Next
		dbCon2.Close()
		'UPGRADE_NOTE: オブジェクト dbCon2 をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbCon2 = Nothing
		On Error GoTo 0
		
		Call ProgramEnd()
		
	End Sub
	
	
	Private Sub inpCDST_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles inpCDST.Enter
		
		mLostFocusCheck = True
		
	End Sub
	
	Private Sub inpCDST_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles inpCDST.Leave
		
		Dim pRetVal As Integer
		Dim pNAME As String
		Dim pKBSP As Integer
		
		'UPGRADE_WARNING: オブジェクト OSC_STM_READ() の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
		pRetVal = OSC_STM_READ(inpCDST.Value, pNAME, False, pKBSP)
		lblNMST.Text = pNAME
		If pKBSP = 0 Then
			'SP評価無し
			optKBHY(1).Enabled = False
			optKBHY(0).Value = True
		Else
			optKBHY(1).Enabled = True
		End If
		
		Select Case FocusMove_HEAD(inpCDST)
			Case "H"
				If Not mLostFocusCheck Then
					Exit Sub
				End If
				
				mLostFocusCheck = False
				Exit Sub
				
			Case "I"
				If InputMode_Check() Then
					Exit Sub
				End If
		End Select
		
	End Sub
	
	
	Private Sub inpYEAR_ChangeValue(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles inpYEAR.ChangeValue
		
		Dim pKBMEK As Integer
		
		If Not mLostFocusCheck Then
			Exit Sub
		End If
		
		
	End Sub
	
	Private Sub inpYEAR_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles inpYEAR.Enter
		
		mLostFocusCheck = True
		
	End Sub
	
	Private Sub inpYEAR_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles inpYEAR.Leave
		
		Dim pRetVal As Short
		
		Select Case FocusMove_HEAD(inpYEAR)
			Case "H"
				If Not mLostFocusCheck Then
					Exit Sub
				End If
				
				mLostFocusCheck = False
				Exit Sub
				
			Case "I"
				If InputMode_Check() Then
					Exit Sub
				End If
		End Select
		
	End Sub
	
	Private Sub lblSUKJ_Click()
	End Sub
	
	Public Sub mnuFILEItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFILEItem.Click
		Dim Index As Short = mnuFILEItem.GetIndex(eventSender)
		
		Dim pRetVal As Short
		
		Select Case Index
			Case 0 ' 更新
				If Kosin_Exec() Then
					'UPGRADE_WARNING: App プロパティ App.EXEName には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
					Call MenuLOGF_Insert(My.Application.Info.AssemblyName, FormCaption, bEXEC, bIDUS, bNMUS)
				End If
				
			Case 1 ' ｷｬﾝｾﾙ
				pRetVal = Kosin_Cancel(False)
				
			Case 2 ' 削除
				pRetVal = Kosin_Delete()
				
			Case 3 ' ﾒﾆｭｰの区分線
				
			Case 5 ' 印刷
				Me.Enabled = False
				frmPK9OSCM004pr.lblNMGB.Text = lblNMGB.Text
				frmPK9OSCM004pr.lblNENDO.Text = lblNENDO.Text
				frmPK9OSCM004pr.Show()
				
			Case 8 ' ﾒﾆｭｰの区分線
				
			Case 9 ' 終了
				Me.Close()
		End Select
		
	End Sub
	
	' *******************************************************************************
	' 概要    : ａｄｏＣｏｍｍａｎｄオブジェクトの作成を行う。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=正常。
	'         :                    False=エラー。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2005.12.12  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function Qry_Create() As Short
		
		Dim pRetVal As Short
		
		Qry_Create = False
		
		'
		' 評価項目マスターの追加    '2006.01.16 update
		'
		SqlText = " insert into OSC_HYOKAKM "
		SqlText = SqlText & "(HYCDGA, HYCDGB, HYCDGK, HYNENDO, HYKBSK, "
		SqlText = SqlText & " HYYEAR, HYCDST, HYKBHY, HYNOSQ, HYCDGP, "
		SqlText = SqlText & " HYKBDAI, HYKBCHU, HYKBKR, HYMONDAI, HYSUKA, "
		SqlText = SqlText & " HYKB1, HYKB2, HYKB3, HYNO,HYKB4)"
		SqlText = SqlText & " values(?, ?, ?, ?, ?, "
		SqlText = SqlText & "        ?, ?, ?, ?, ?, "
		SqlText = SqlText & "        ?, ?, ?, ?, ?, "
		SqlText = SqlText & "        ?, ?, ?, ?, ?)"
		
		On Error Resume Next
		dbCmdINS = New ADODB.Command
		dbCmdINS.let_ActiveConnection(dbCon2)
		dbCmdINS.CommandText = SqlText
		dbCmdINS.CommandType = ADODB.CommandTypeEnum.adCmdText
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター追加のコマンド作成でエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description, dbCon2)
			Err.Clear()
			On Error GoTo 0
			Exit Function
		End If
		On Error GoTo 0
		
		'
		' 評価項目マスターの削除
		'
		SqlText = "delete from OSC_HYOKAKM "
		SqlText = SqlText & " where HYCDGA = ? " ' PARM=0
		SqlText = SqlText & "   and HYCDGB = ? " ' PARM=1
		SqlText = SqlText & "   and HYCDGK = ? " ' PARM=2
		SqlText = SqlText & "   and HYNENDO = ? " ' PARM=3
		SqlText = SqlText & "   and HYKBSK = ? " ' PARM=4
		SqlText = SqlText & "   and HYYEAR = ? " ' PARM=5
		SqlText = SqlText & "   and HYCDST = ? " ' PARM=6
		SqlText = SqlText & "   and HYKBHY = ? " ' PARM=7
		
		On Error Resume Next
		dbCmdDEL = New ADODB.Command
		dbCmdDEL.let_ActiveConnection(dbCon2)
		dbCmdDEL.CommandText = SqlText
		dbCmdDEL.CommandType = ADODB.CommandTypeEnum.adCmdText
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター削除のコマンド作成でエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description, dbCon2)
			Err.Clear()
			On Error GoTo 0
			Exit Function
		End If
		On Error GoTo 0
		
		'
		' ステーションマスターの追加（概略評価）
		'
		SqlText = " insert into OSC_HYOKAKMG "
		SqlText = SqlText & "(HGCDGA, HGCDGB, HGCDGK, HGNENDO, HGKBSK, "
		SqlText = SqlText & " HGYEAR, HGCDST, HGKBHY, HGNOSQ, HGCDGP, "
		SqlText = SqlText & " HGKBDAI, HGKBKR, HGHYOKA, HGNAIYO, HGKB1)"
		SqlText = SqlText & " values(?, ?, ?, ?, ?, "
		SqlText = SqlText & "        ?, ?, ?, ?, ?, "
		SqlText = SqlText & "        ?, ?, ?, ?, ?)"
		
		On Error Resume Next
		dbCmdINSG = New ADODB.Command
		dbCmdINSG.let_ActiveConnection(dbCon2)
		dbCmdINSG.CommandText = SqlText
		dbCmdINSG.CommandType = ADODB.CommandTypeEnum.adCmdText
		
		If Err.Number <> 0 Then
			mMsgText = "ステーションマスター（概略評価）追加のコマンド作成でエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description, dbCon2)
			Err.Clear()
			On Error GoTo 0
			Exit Function
		End If
		On Error GoTo 0
		
		'
		' ステーションマスターの削除（概略評価）
		'
		SqlText = "delete from OSC_HYOKAKMG "
		SqlText = SqlText & " where HGCDGA = ? " ' PARM=0
		SqlText = SqlText & "   and HGCDGB = ? " ' PARM=1
		SqlText = SqlText & "   and HGCDGK = ? " ' PARM=2
		SqlText = SqlText & "   and HGNENDO = ? " ' PARM=3
		SqlText = SqlText & "   and HGKBSK = ? " ' PARM=4
		SqlText = SqlText & "   and HGYEAR = ? " ' PARM=5
		SqlText = SqlText & "   and HGCDST = ? " ' PARM=6
		SqlText = SqlText & "   and HGKBHY = ? " ' PARM=7
		
		On Error Resume Next
		dbCmdDELG = New ADODB.Command
		dbCmdDELG.let_ActiveConnection(dbCon2)
		dbCmdDELG.CommandText = SqlText
		dbCmdDELG.CommandType = ADODB.CommandTypeEnum.adCmdText
		
		If Err.Number <> 0 Then
			mMsgText = "ステーションマスター（概略評価）削除のコマンド作成でエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description, dbCon2)
			Err.Clear()
			On Error GoTo 0
			Exit Function
		End If
		On Error GoTo 0
		
		'
		' ステーションマスターの追加（問題点）
		'
		SqlText = " insert into OSC_HYOKAKMM "
		SqlText = SqlText & "(HMCDGA, HMCDGB, HMCDGK, HMNENDO, HMKBSK, "
		SqlText = SqlText & " HMYEAR, HMCDST, HMKBHY, HMNOSQ, HMCDGP, "
		SqlText = SqlText & " HMKBDAI, HMKBKR, HMMONDAI, HMKB1)"
		SqlText = SqlText & " values(?, ?, ?, ?, ?, "
		SqlText = SqlText & "        ?, ?, ?, ?, ?, "
		SqlText = SqlText & "        ?, ?, ?, ?)"
		
		On Error Resume Next
		dbCmdINSM = New ADODB.Command
		dbCmdINSM.let_ActiveConnection(dbCon2)
		dbCmdINSM.CommandText = SqlText
		dbCmdINSM.CommandType = ADODB.CommandTypeEnum.adCmdText
		
		If Err.Number <> 0 Then
			mMsgText = "ステーションマスター（問題点）追加のコマンド作成でエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description, dbCon2)
			Err.Clear()
			On Error GoTo 0
			Exit Function
		End If
		On Error GoTo 0
		
		'
		' ステーションマスターの削除（問題点）
		'
		SqlText = "delete from OSC_HYOKAKMM "
		SqlText = SqlText & " where HMCDGA = ? " ' PARM=0
		SqlText = SqlText & "   and HMCDGB = ? " ' PARM=1
		SqlText = SqlText & "   and HMCDGK = ? " ' PARM=2
		SqlText = SqlText & "   and HMNENDO = ? " ' PARM=3
		SqlText = SqlText & "   and HMKBSK = ? " ' PARM=4
		SqlText = SqlText & "   and HMYEAR = ? " ' PARM=5
		SqlText = SqlText & "   and HMCDST = ? " ' PARM=6
		SqlText = SqlText & "   and HMKBHY = ? " ' PARM=7
		
		On Error Resume Next
		dbCmdDELM = New ADODB.Command
		dbCmdDELM.let_ActiveConnection(dbCon2)
		dbCmdDELM.CommandText = SqlText
		dbCmdDELM.CommandType = ADODB.CommandTypeEnum.adCmdText
		
		If Err.Number <> 0 Then
			mMsgText = "ステーションマスター（問題点）削除のコマンド作成でエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description, dbCon2)
			Err.Clear()
			On Error GoTo 0
			Exit Function
		End If
		On Error GoTo 0
		
		Qry_Create = True
		
	End Function
	
	' *******************************************************************************
	' 概要    : Ｒｅｑｕｅｒｙ用にＲｅｃｏｒｄｓｅｔを生成する。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=正常終了。
	'         :                    False=異常終了。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2005.12.02  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function ResultSet_Initialize() As Short
		
		Dim pRetVal As Short
		
		ResultSet_Initialize = False
		
		
		
		ResultSet_Initialize = True
		
	End Function
	
	' *******************************************************************************
	' 概要  : 評価項目マスターを更新する。
	' 　　  :
	' 引数  : 戻り値, O, Integer, True=正常終了。
	' 　　  : 　　　              False=異常終了。
	' 　　  :
	' 説明  :
	' 　　  :
	' 履歴  : 2005.12.12  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function Table_Insert() As Short
		
		Dim pRetVal As Short
		Dim pIX As Integer
		Dim pIIX As Integer
		Dim pIIIX As Integer
		Dim pCOL As Integer
		Dim pSUKA As Integer
		Dim pLEN As Short
		
		Table_Insert = False
		
		''----------------------------
		' トランザクション開始
		''----------------------------
		On Error Resume Next
		dbCon2.Execute("BEGIN TRAN",  , ADODB.CommandTypeEnum.adCmdText + ADODB.ExecuteOptionEnum.adExecuteNoRecords)
		On Error GoTo 0
		
		''-------------
		' 評価項目マスター削除処理
		''-------------
		dbCmdDEL.Parameters(0).Value = 1
		dbCmdDEL.Parameters(1).Value = bCDGB
		dbCmdDEL.Parameters(2).Value = mCDGK
		dbCmdDEL.Parameters(3).Value = mNENDO
		dbCmdDEL.Parameters(4).Value = mKBSK
		dbCmdDEL.Parameters(5).Value = mYEAR
		dbCmdDEL.Parameters(6).Value = mCDST
		dbCmdDEL.Parameters(7).Value = mKBHY
		
		On Error Resume Next
		dbCmdDEL.Execute()
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター削除でエラーが発生しました。"
			'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
			GoSub Table_Insert_Err
		End If
		On Error GoTo 0
		
		''-------------
		' 評価項目マスター削除処理
		''-------------
		dbCmdDEL.Parameters(0).Value = 1
		dbCmdDEL.Parameters(1).Value = bCDGB
		dbCmdDEL.Parameters(2).Value = mCDGK
		dbCmdDEL.Parameters(3).Value = mNENDO
		dbCmdDEL.Parameters(4).Value = mKBSK
		dbCmdDEL.Parameters(5).Value = mYEAR
		dbCmdDEL.Parameters(6).Value = mCDST
		dbCmdDEL.Parameters(7).Value = mKBHY
		
		On Error Resume Next
		dbCmdDEL.Execute()
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター削除でエラーが発生しました。"
			'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
			GoSub Table_Insert_Err
		End If
		On Error GoTo 0
		
		''-------------
		' 評価項目マスター削除処理（G）
		''-------------
		dbCmdDELG.Parameters(0).Value = 1
		dbCmdDELG.Parameters(1).Value = bCDGB
		dbCmdDELG.Parameters(2).Value = mCDGK
		dbCmdDELG.Parameters(3).Value = mNENDO
		dbCmdDELG.Parameters(4).Value = mKBSK
		dbCmdDELG.Parameters(5).Value = mYEAR
		dbCmdDELG.Parameters(6).Value = mCDST
		dbCmdDELG.Parameters(7).Value = mKBHY
		
		On Error Resume Next
		dbCmdDELG.Execute()
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター(G)削除でエラーが発生しました。"
			'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
			GoSub Table_Insert_Err
		End If
		On Error GoTo 0
		
		''-------------
		' 評価項目マスター削除処理（M）
		''-------------
		dbCmdDELM.Parameters(0).Value = 1
		dbCmdDELM.Parameters(1).Value = bCDGB
		dbCmdDELM.Parameters(2).Value = mCDGK
		dbCmdDELM.Parameters(3).Value = mNENDO
		dbCmdDELM.Parameters(4).Value = mKBSK
		dbCmdDELM.Parameters(5).Value = mYEAR
		dbCmdDELM.Parameters(6).Value = mCDST
		dbCmdDELM.Parameters(7).Value = mKBHY
		
		On Error Resume Next
		dbCmdDELM.Execute()
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター(M)削除でエラーが発生しました。"
			'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
			GoSub Table_Insert_Err
		End If
		On Error GoTo 0
		
		''-------------
		' 追加処理
		''-------------
		If mInputMode <> "D" Then
			For pIX = 0 To (PGrid.Items - 1)
				If Not Grid1SpaceGyoCheck(PGrid, pIX) Then
					'If Trim(PGrid.CellTextByName(pIX, "CDGP")) <> "" Then
					'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
					GoSub Table_Insert_INS
					'End If
					''            Else
					''                Exit For
				End If
			Next pIX
			'概評
			For pIX = 0 To (PGrid2.Items - 1)
				If Not Grid1SpaceGyoCheck(PGrid2, pIX) Then
					'If Trim(PGrid2.CellTextByName(pIX, "STTM")) <> "" Then
					'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
					GoSub Table_Insert_INSG
					'End If
					''            Else
					''                Exit For
				End If
			Next pIX
			'問題
			For pIX = 0 To (PGrid3.Items - 1)
				If Not Grid1SpaceGyoCheck(PGrid3, pIX) Then
					'If Trim(PGrid2.CellTextByName(pIX, "STTM")) <> "" Then
					'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
					GoSub Table_Insert_INSM
					'End If
					''            Else
					''                Exit For
				End If
			Next pIX
		End If
		
		''----------------------------
		' コミット
		''----------------------------
		'    If UCase(DBTYPE) = "SQL" Then
		On Error Resume Next
		dbCon2.Execute("COMMIT TRAN",  , ADODB.CommandTypeEnum.adCmdText + ADODB.ExecuteOptionEnum.adExecuteNoRecords)
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター追加のコミットでエラーが発生しました。"
			'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
			GoSub Table_Insert_Err
		End If
		On Error GoTo 0
		
		Table_Insert = True
		Exit Function
		
		' --------------------------------------------------------------------
		' 追加処理
		' --------------------------------------------------------------------
Table_Insert_INS: 
		
		''----------------------
		''評価項目マスターINS
		''----------------------
		dbCmdINS.Parameters(0).Value = 1
		dbCmdINS.Parameters(1).Value = bCDGB
		dbCmdINS.Parameters(2).Value = mCDGK
		dbCmdINS.Parameters(3).Value = mNENDO
		dbCmdINS.Parameters(4).Value = mKBSK
		dbCmdINS.Parameters(5).Value = mYEAR
		dbCmdINS.Parameters(6).Value = mCDST
		dbCmdINS.Parameters(7).Value = mKBHY
		
		'連番
		dbCmdINS.Parameters(8).Value = pIX + 1
		
		If PGrid.get_CellCheckedByName(pIX, "KBKR") = True Then
			''空行----------------------
			'グループ
			dbCmdINS.Parameters(9).Value = 0
			'大問
			dbCmdINS.Parameters(10).Value = 0
			'中門
			dbCmdINS.Parameters(11).Value = 0
			'空
			dbCmdINS.Parameters(12).Value = 1
			'設問
			'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
			dbCmdINS.Parameters(13).Value = System.DBNull.Value
			'解答数
			dbCmdINS.Parameters(14).Value = 0
			'解答有無１
			dbCmdINS.Parameters(15).Value = 0
			'解答有無２
			dbCmdINS.Parameters(16).Value = 0
			'解答有無３
			dbCmdINS.Parameters(17).Value = 0
			'問題番号   '2006.01.16 add
			dbCmdINS.Parameters(18).Value = 0
			'解答有無４
			dbCmdINS.Parameters(19).Value = 0
		Else
			If PGrid.get_CellCheckedByName(pIX, "KBDAI") = True Then
				''大問----------------------
				'グループ
				dbCmdINS.Parameters(9).Value = PGrid.get_CellValueByName(pIX, "CDGP")
				'大問
				dbCmdINS.Parameters(10).Value = 1
				'中門
				dbCmdINS.Parameters(11).Value = 0
				'空
				dbCmdINS.Parameters(12).Value = 0
				'設問
				dbCmdINS.Parameters(13).Value = RTrim(PGrid.get_CellTextByName(pIX, "MONDAI"))
				'解答数
				dbCmdINS.Parameters(14).Value = PGrid.get_CellValueByName(pIX, "SUKA")
				pSUKA = PGrid.get_CellValueByName(pIX, "SUKA")
				'解答有無１
				dbCmdINS.Parameters(15).Value = 0
				'解答有無２
				dbCmdINS.Parameters(16).Value = 0
				'解答有無３
				dbCmdINS.Parameters(17).Value = 0
				'問題番号   '2006.01.16 add
				dbCmdINS.Parameters(18).Value = 0
				'解答有無４
				dbCmdINS.Parameters(19).Value = 0
			ElseIf PGrid.get_CellCheckedByName(pIX, "KBCHU") Then 
				''中問----------------------
				'グループ
				dbCmdINS.Parameters(9).Value = PGrid.get_CellValueByName(pIX, "CDGP")
				'大問
				dbCmdINS.Parameters(10).Value = 0
				'中門
				dbCmdINS.Parameters(11).Value = 1
				'空
				dbCmdINS.Parameters(12).Value = 0
				'設問
				dbCmdINS.Parameters(13).Value = RTrim(PGrid.get_CellTextByName(pIX, "MONDAI"))
				'解答数
				dbCmdINS.Parameters(14).Value = 0
				'解答有無１
				dbCmdINS.Parameters(15).Value = 0
				'解答有無２
				dbCmdINS.Parameters(16).Value = 0
				'解答有無３
				dbCmdINS.Parameters(17).Value = 0
				'問題番号   '2006.01.16 add
				dbCmdINS.Parameters(18).Value = 0
				'解答有無４
				dbCmdINS.Parameters(19).Value = 0
			Else
				''そのた----------------------
				'グループ
				dbCmdINS.Parameters(9).Value = PGrid.get_CellValueByName(pIX, "CDGP")
				'大問
				dbCmdINS.Parameters(10).Value = 0
				'中門
				dbCmdINS.Parameters(11).Value = 0
				'空
				dbCmdINS.Parameters(12).Value = 0
				'設問
				dbCmdINS.Parameters(13).Value = RTrim(PGrid.get_CellTextByName(pIX, "MONDAI"))
				'解答数
				dbCmdINS.Parameters(14).Value = 0
				If pSUKA = 4 Then
					'解答有無１
					dbCmdINS.Parameters(15).Value = IIf(PGrid.get_CellCheckedByName(pIX, "KB1") = True, 1, 0)
					'解答有無２
					dbCmdINS.Parameters(16).Value = IIf(PGrid.get_CellCheckedByName(pIX, "KB2") = True, 1, 0)
					'解答有無３
					dbCmdINS.Parameters(17).Value = IIf(PGrid.get_CellCheckedByName(pIX, "KB3") = True, 1, 0)
					'解答有無４
					dbCmdINS.Parameters(19).Value = IIf(PGrid.get_CellCheckedByName(pIX, "KB4") = True, 1, 0)
				ElseIf pSUKA = 3 Then 
					'解答有無１
					dbCmdINS.Parameters(15).Value = 0
					'解答有無２
					dbCmdINS.Parameters(16).Value = IIf(PGrid.get_CellCheckedByName(pIX, "KB2") = True, 1, 0)
					'解答有無３
					dbCmdINS.Parameters(17).Value = IIf(PGrid.get_CellCheckedByName(pIX, "KB3") = True, 1, 0)
					'解答有無４
					dbCmdINS.Parameters(19).Value = IIf(PGrid.get_CellCheckedByName(pIX, "KB4") = True, 1, 0)
				ElseIf pSUKA = 2 Then 
					'解答有無1
					dbCmdINS.Parameters(15).Value = 0
					'解答有無２
					dbCmdINS.Parameters(16).Value = 0
					'解答有無３
					dbCmdINS.Parameters(17).Value = IIf(PGrid.get_CellCheckedByName(pIX, "KB3") = True, 1, 0)
					'解答有無４
					dbCmdINS.Parameters(19).Value = IIf(PGrid.get_CellCheckedByName(pIX, "KB4") = True, 1, 0)
				ElseIf pSUKA = 1 Then 
					'解答有無1
					dbCmdINS.Parameters(15).Value = 0
					'解答有無２
					dbCmdINS.Parameters(16).Value = 0
					'解答有無３
					dbCmdINS.Parameters(17).Value = 0
					'解答有無４
					dbCmdINS.Parameters(19).Value = IIf(PGrid.get_CellCheckedByName(pIX, "KB4") = True, 1, 0)
				Else
					'解答有無1
					dbCmdINS.Parameters(15).Value = 0
					'解答有無２
					dbCmdINS.Parameters(16).Value = 0
					'解答有無３
					dbCmdINS.Parameters(17).Value = 0
					'解答有無４
					dbCmdINS.Parameters(19).Value = 0
				End If
				'問題番号   '2006.01.16 add
				dbCmdINS.Parameters(18).Value = PGrid.get_CellValueByName(pIX, "NO")
			End If
		End If
		
		On Error Resume Next
		dbCmdINS.Execute()
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター追加でエラーが発生しました。"
			'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
			GoSub Table_Insert_Err
		End If
		On Error GoTo 0
		'UPGRADE_WARNING: Return に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
		Return 
		
Table_Insert_INSG: 
		
		''----------------------
		''評価項目マスターINS(概略評価)
		''----------------------
		dbCmdINSG.Parameters(0).Value = 1
		dbCmdINSG.Parameters(1).Value = bCDGB
		dbCmdINSG.Parameters(2).Value = mCDGK
		dbCmdINSG.Parameters(3).Value = mNENDO
		dbCmdINSG.Parameters(4).Value = mKBSK
		dbCmdINSG.Parameters(5).Value = mYEAR
		dbCmdINSG.Parameters(6).Value = mCDST
		dbCmdINSG.Parameters(7).Value = mKBHY
		
		'連番
		dbCmdINSG.Parameters(8).Value = pIX + 1
		
		If PGrid2.get_CellCheckedByName(pIX, "KBKR") = True Then
			''空行----------------------
			'グループ
			dbCmdINSG.Parameters(9).Value = 0
			'大問
			dbCmdINSG.Parameters(10).Value = 0
			'空
			dbCmdINSG.Parameters(11).Value = 1
			'評価
			dbCmdINSG.Parameters(12).Value = 0
			'内容
			'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
			dbCmdINSG.Parameters(13).Value = System.DBNull.Value
			'解答有無
			dbCmdINSG.Parameters(14).Value = 0
			
		Else
			If PGrid2.get_CellCheckedByName(pIX, "KBDAI") = True Then
				''大問----------------------
				'グループ
				dbCmdINSG.Parameters(9).Value = PGrid2.get_CellValueByName(pIX, "CDGP")
				'大問
				dbCmdINSG.Parameters(10).Value = 1
				'空
				dbCmdINSG.Parameters(11).Value = 0
				'評価
				dbCmdINSG.Parameters(12).Value = 0
				'内容
				If mKBHY = 0 Then
					dbCmdINSG.Parameters(13).Value = MOJIByte_Get(PGrid2.get_CellTextByName(pIX, "NAIYO"), Me, 1, 50, pLEN)
				Else
					dbCmdINSG.Parameters(13).Value = MOJIByte_Get(PGrid2.get_CellTextByName(pIX, "NAIYO"), Me, 1, 80, pLEN)
				End If
				'解答有無
				dbCmdINSG.Parameters(14).Value = 0
			Else
				''そのた----------------------
				'グループ
				dbCmdINSG.Parameters(9).Value = PGrid2.get_CellValueByName(pIX, "CDGP")
				'大問
				dbCmdINSG.Parameters(10).Value = 0
				'空
				dbCmdINSG.Parameters(11).Value = 0
				'評価
				dbCmdINSG.Parameters(12).Value = PGrid2.get_CellValueByName(pIX, "HYOKA")
				'内容
				If mKBHY = 0 Then
					dbCmdINSG.Parameters(13).Value = MOJIByte_Get(PGrid2.get_CellTextByName(pIX, "NAIYO"), Me, 1, 44, pLEN)
				Else
					dbCmdINSG.Parameters(13).Value = MOJIByte_Get(PGrid2.get_CellTextByName(pIX, "NAIYO"), Me, 1, 80, pLEN)
				End If
				'解答有無
				dbCmdINSG.Parameters(14).Value = IIf(PGrid2.get_CellCheckedByName(pIX, "KB1") = True, 1, 0)
			End If
		End If
		On Error Resume Next
		dbCmdINSG.Execute()
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター追加でエラーが発生しました。"
			'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
			GoSub Table_Insert_Err
		End If
		On Error GoTo 0
		'UPGRADE_WARNING: Return に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
		Return 
		
Table_Insert_INSM: 
		
		''----------------------
		''評価項目マスターINS(問題点)
		''----------------------
		dbCmdINSM.Parameters(0).Value = 1
		dbCmdINSM.Parameters(1).Value = bCDGB
		dbCmdINSM.Parameters(2).Value = mCDGK
		dbCmdINSM.Parameters(3).Value = mNENDO
		dbCmdINSM.Parameters(4).Value = mKBSK
		dbCmdINSM.Parameters(5).Value = mYEAR
		dbCmdINSM.Parameters(6).Value = mCDST
		dbCmdINSM.Parameters(7).Value = mKBHY
		
		'連番
		dbCmdINSM.Parameters(8).Value = pIX + 1
		
		If PGrid3.get_CellCheckedByName(pIX, "KBKR") = True Then
			''空行----------------------
			'グループ
			dbCmdINSM.Parameters(9).Value = 0
			'大問
			dbCmdINSM.Parameters(10).Value = 0
			'空
			dbCmdINSM.Parameters(11).Value = 1
			'問題点
			'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
			dbCmdINSM.Parameters(12).Value = System.DBNull.Value
			'解答有無
			dbCmdINSM.Parameters(13).Value = 0
			
		Else
			If PGrid3.get_CellCheckedByName(pIX, "KBDAI") = True Then
				''大問----------------------
				'グループ
				dbCmdINSM.Parameters(9).Value = PGrid3.get_CellValueByName(pIX, "CDGP")
				'大問
				dbCmdINSM.Parameters(10).Value = 1
				'空
				dbCmdINSM.Parameters(11).Value = 0
				'問題点
				If mKBHY = 0 Then
					dbCmdINSM.Parameters(12).Value = MOJIByte_Get(PGrid3.get_CellTextByName(pIX, "MONDAI"), Me, 1, 60, pLEN)
				Else
					dbCmdINSM.Parameters(12).Value = MOJIByte_Get(PGrid3.get_CellTextByName(pIX, "MONDAI"), Me, 1, 80, pLEN)
				End If
				'解答有無
				dbCmdINSM.Parameters(13).Value = 0
			Else
				''そのた----------------------
				'グループ
				dbCmdINSM.Parameters(9).Value = PGrid3.get_CellValueByName(pIX, "CDGP")
				'大問
				dbCmdINSM.Parameters(10).Value = 0
				'空
				dbCmdINSM.Parameters(11).Value = 0
				'問題点
				If mKBHY = 0 Then
					dbCmdINSM.Parameters(12).Value = MOJIByte_Get(PGrid3.get_CellTextByName(pIX, "MONDAI"), Me, 1, 44, pLEN)
				Else
					dbCmdINSM.Parameters(12).Value = MOJIByte_Get(PGrid3.get_CellTextByName(pIX, "MONDAI"), Me, 1, 80, pLEN)
				End If
				'解答有無
				dbCmdINSM.Parameters(13).Value = IIf(PGrid3.get_CellCheckedByName(pIX, "KB1") = True, 1, 0)
			End If
		End If
		
		On Error Resume Next
		dbCmdINSM.Execute()
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター追加でエラーが発生しました。"
			'UPGRADE_ISSUE: GoSub ステートメントはサポートされていません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"' をクリックしてください。
			GoSub Table_Insert_Err
		End If
		On Error GoTo 0
		'UPGRADE_WARNING: Return に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
		Return 
		
		' --------------------------------------------------------------------
		' エラー処理
		' --------------------------------------------------------------------
Table_Insert_Err: 
		
		pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description, dbCon2)
		Err.Clear()
		
		dbCon2.Execute("ROLLBACK TRAN",  , ADODB.CommandTypeEnum.adCmdText + ADODB.ExecuteOptionEnum.adExecuteNoRecords)
		
		On Error GoTo 0
		Exit Function
		
	End Function
	
	' *******************************************************************************
	' 概要    : マスターをＲＥＡＤし、登録モードのチェックを行う。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=正常終了。
	'         :                    False=エラーあり。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2005.12.12  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function InputMode_Check() As Short
		
		Dim pRetVal As Short
		Dim pMENU As System.Windows.Forms.ToolStripMenuItem
		Dim pIX As Short
		Dim pWW As Integer
		Dim pII As Integer
		Dim pKBMEK As Integer
		Dim pGYO As Integer
		Dim pROW As Integer
		Dim pNAME As String
		Dim pKBSP As Integer
		
		InputMode_Check = False
		'
		' キー項目のチェック
		'
		If cboCDGK.ListIndex = -1 Then
			mMsgText = "学科を選択してください。"
			pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, mMBOXTitle)
			''Call GridClear
			cboCDGK.Focus()
			Exit Function
		End If
		mCDGK = cboCDGK.get_ItemData(cboCDGK.ListIndex)
		
		'試験区分
		If optKBSK(0).Value = True Then
			mKBSK = 1
		ElseIf optKBSK(1).Value = True Then 
			mKBSK = 2
		End If
		
		'学年
		If inpYEAR.Value = 0 Then
			mMsgText = "学年を入力してください。"
			pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, mMBOXTitle)
			''Call GridClear
			inpYEAR.Focus()
			Exit Function
		End If
		mYEAR = inpYEAR.Value
		
		'ステーション
		If inpCDST.Value = 0 Then
			mMsgText = "ステーションを入力してください。"
			pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, mMBOXTitle)
			''Call GridClear
			inpCDST.Focus()
			Exit Function
		Else
			If Not OSC_STM_READ(inpCDST.Value, pNAME, False, pKBSP) Then
				lblNMST.Text = ""
				mMsgText = "ステーションマスターに登録されていません。"
				pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, mMBOXTitle)
				inpCDST.Focus()
				Exit Function
			End If
			lblNMST.Text = pNAME
		End If
		mCDST = inpCDST.Value
		
		'評価者
		If optKBHY(0).Value = True Then
			mKBHY = 0
			
		ElseIf optKBHY(1).Value = True Then 
			mKBHY = 1
		End If
		
		If mKBHY = 1 And pKBSP = 0 Then
			mMsgText = "このステーションは、模擬患者による評価対象として登録されていません。"
			pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, mMBOXTitle)
			inpCDST.Focus()
			Exit Function
		End If
		
		'
		' マスターをＲＥＡＤし、登録モードを判断する
		'
		mSW_CellFocusEvent = False
		
		''評価者区分によって、最大行を切り替える(評価)----------2006.01.06 add
		If mKBHY = 0 Then
			pGYO = 50
		ElseIf mKBHY = 1 Then 
			pGYO = 26
		End If
		
		PGrid.RefreshLater = True
		pROW = PGrid.Items
		PGrid.RemoveItems(0, pROW)
		PGrid.TextAtAddItem = ""
		PGrid.AddItems(0, pGYO)
		
		For pIX = 0 To (PGrid.Items - 1)
			PGrid.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
		Next pIX
		
		PGrid.RefreshLater = False
		'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		PGrid.CtlRefresh()
		
		''評価者区分によって、最大行を切り替える(概評・問題点)----------
		If mKBHY = 0 Then
			pGYO = 7
		ElseIf mKBHY = 1 Then 
			pGYO = 15
		End If
		
		PGrid2.RefreshLater = True
		pROW = PGrid2.Items
		PGrid2.RemoveItems(0, pROW)
		PGrid2.TextAtAddItem = ""
		PGrid2.AddItems(0, pGYO)
		
		For pIX = 0 To (PGrid2.Items - 1)
			PGrid2.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
		Next pIX
		
		PGrid2.RefreshLater = False
		'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		PGrid2.CtlRefresh()
		
		PGrid3.RefreshLater = True
		pROW = PGrid3.Items
		PGrid3.RemoveItems(0, pROW)
		PGrid3.TextAtAddItem = ""
		PGrid3.AddItems(0, pGYO)
		
		For pIX = 0 To (PGrid3.Items - 1)
			PGrid3.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
		Next pIX
		
		PGrid3.RefreshLater = False
		'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		PGrid3.CtlRefresh()
		
		''評価者区分によって、タイトルを切り替える----------
		'2009.06.03 upd ---
		If bCDGB = 3 Then
			If mKBHY = 0 Then
				PGrid.set_CellTextByName(-1, "KB1", "")
				PGrid.set_CellTextByName(-1, "KB2", "")
				PGrid.set_CellTextByName(-1, "KB3", "はい")
				PGrid.set_CellTextByName(-1, "KB4", "いいえ")
			ElseIf mKBHY = 1 Then 
				PGrid.set_CellTextByName(-1, "KB1", "")
				PGrid.set_CellTextByName(-1, "KB2", "はい")
				''PGrid.CellTextByName(-1, "KB3") = ""
				PGrid.set_CellTextByName(-1, "KB3", "中間") '2006.12.21 update（とりあえず固定。メンテできるようにする）
				PGrid.set_CellTextByName(-1, "KB4", "いいえ")
			End If
		Else
			If mKBHY = 0 Then
				PGrid.set_CellTextByName(-1, "KB1", "３")
				PGrid.set_CellTextByName(-1, "KB2", "２")
				PGrid.set_CellTextByName(-1, "KB3", "１")
				PGrid.set_CellTextByName(-1, "KB4", "０")
			ElseIf mKBHY = 1 Then 
				'2010.02.12 del ---
				''PGrid.CellTextByName(-1, "KB1") = ""
				''PGrid.CellTextByName(-1, "KB2") = "はい"
				''''PGrid.CellTextByName(-1, "KB3") = ""
				''PGrid.CellTextByName(-1, "KB3") = "中間"    '2006.12.21 update（とりあえず固定。メンテできるようにする）
				''PGrid.CellTextByName(-1, "KB4") = "いいえ"
				'2010.02.12 add ---
				PGrid.set_CellTextByName(-1, "KB1", "３")
				PGrid.set_CellTextByName(-1, "KB2", "２")
				PGrid.set_CellTextByName(-1, "KB3", "１")
				PGrid.set_CellTextByName(-1, "KB4", "０")
				'------------------
			End If
		End If
		'-----------------------------------------------
		
		If OSC_HYOKAKM_Read() Then
			mInputMode = "U"
		Else
			mInputMode = "E"
		End If
		
		If OSC_HYOKAKMG_READ() Then
			mInputMode = "U"
		Else
			''mInputMode = "E"
		End If
		
		If OSC_HYOKAKMM_READ() Then
			mInputMode = "U"
		Else
			''mInputMode = "E"
		End If
		
		
		
		''Call SU_Disp
		
		mSW_CellFocusEvent = True
		
		' ------------------------------------------------------------
		mLostFocusCheck = False
		
		picHEAD.Enabled = False
		Toolbar1.Items.Item("EXEC").Enabled = True
		Toolbar1.Items.Item("CANCEL").Enabled = True
		Toolbar1.Items.Item("PRINT").Enabled = False
		Toolbar1.Items.Item("EXIT").Enabled = False
		Toolbar1.Items.Item("ROWINSERT").Enabled = True
		Toolbar1.Items.Item("ROWDELETE").Enabled = True
		Toolbar1.Items.Item("COPY").Enabled = False
		
		mnuFILEItem(0).Enabled = True
		mnuFILEItem(1).Enabled = True
		mnuFILEItem(5).Enabled = False
		mnuFILEItem(9).Enabled = False
		mnuEDITItem(0).Enabled = True
		mnuEDITItem(1).Enabled = True
		mnuEDITItem(9).Enabled = False
		
		If mInputMode = "U" Then
			mnuFILEItem(2).Enabled = True
		End If
		
		If PGrid.get_CellCheckedByName(0, "KBKR") = False Then
			pRetVal = PGrid.SelectCell(0, PGrid.get_ColOfColName("CDGP"))
		Else
			pRetVal = PGrid.SelectCell(0, PGrid.get_ColOfColName("KBKR"))
		End If
		
		' --------------------------------------
		InputMode_Check = True
		
	End Function
	
	' *******************************************************************************
	' 概要    : 評価項目マスターのＲＥＡＤを行う。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=READ成功。
	'         :                    False=READ失敗。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2005.12.12  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function OSC_HYOKAKM_Read() As Short
		
		Dim pRetVal As Short
		Dim pROW As Integer
		Dim pNM As String
		Dim dbRec As ADODB.Recordset
		Dim pMAXLen As Short
		Dim pIX As Short
		Dim pIndex As Short
		Dim pFMT As String
		Dim pMVL As String
		Dim pCOL As Integer
		
		OSC_HYOKAKM_Read = False
		
		dbRec = New ADODB.Recordset
		
		'
		' 評価項目マスターＲＥＡＤ
		'
		SqlText = "select * "
		SqlText = SqlText & " from OSC_HYOKAKM "
		SqlText = SqlText & " where HYCDGA  = 1 "
		SqlText = SqlText & "   and HYCDGB  = " & CStr(bCDGB)
		SqlText = SqlText & "   and HYCDGK  = " & CStr(mCDGK)
		SqlText = SqlText & "   and HYNENDO = " & CStr(mNENDO)
		SqlText = SqlText & "   and HYKBSK  = " & CStr(mKBSK)
		SqlText = SqlText & "   and HYYEAR  = " & CStr(mYEAR)
		SqlText = SqlText & "   and HYCDST  = " & CStr(mCDST)
		SqlText = SqlText & "   and HYKBHY  = " & CStr(mKBHY)
		SqlText = SqlText & " order by HYCDGA, HYCDGB, HYCDGK,HYNENDO, HYKBSK, HYYEAR, HYCDST, HYKBHY, HYNOSQ "
		
		On Error Resume Next
		dbRec.Open(SqlText, dbCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスターのＲＥＡＤでエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
			Err.Clear()
			On Error GoTo 0
			GoTo RecordSet_Close
		End If
		On Error GoTo 0
		
		If dbRec.EOF = True Then
			GoTo RecordSet_Close
		End If
		
		pROW = -1
		Do While dbRec.EOF = False
			
			pROW = pROW + 1
			PGrid.set_CellValueByName(pROW, "CDGP", dbRec.Fields("HYCDGP").Value)
			
			PGrid.set_CellCheckedByName(pROW, "KBDAI", IIf(dbRec.Fields("HYKBDAI").Value = 1, True, False))
			PGrid.set_CellCheckedByName(pROW, "KBCHU", IIf(dbRec.Fields("HYKBCHU").Value = 1, True, False))
			PGrid.set_CellCheckedByName(pROW, "KBKR", IIf(dbRec.Fields("HYKBKR").Value = 1, True, False))
			'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
			PGrid.set_CellTextByName(pROW, "MONDAI", IIf(IsDbNull(dbRec.Fields("HYMONDAI").Value), "", dbRec.Fields("HYMONDAI").Value))
			'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
			PGrid.set_CellTextByName(pROW, "NO", IIf(IsDbNull(dbRec.Fields("HYNO").Value), "", dbRec.Fields("HYNO").Value)) '2006.01.16 add
			PGrid.set_CellValueByName(pROW, "SUKA", dbRec.Fields("HYSUKA").Value)
			PGrid.set_CellCheckedByName(pROW, "KB1", IIf(dbRec.Fields("HYKB1").Value = 1, True, False))
			PGrid.set_CellCheckedByName(pROW, "KB2", IIf(dbRec.Fields("HYKB2").Value = 1, True, False))
			PGrid.set_CellCheckedByName(pROW, "KB3", IIf(dbRec.Fields("HYKB3").Value = 1, True, False))
			PGrid.set_CellCheckedByName(pROW, "KB4", IIf(dbRec.Fields("HYKB4").Value = 1, True, False)) '2006.02.20 add
			
			Call ENABLE_Change(pROW)
			
			
			On Error Resume Next
			dbRec.MoveNext()
			
			If Err.Number <> 0 Then
				mMsgText = "評価項目マスターのＲＥＡＤでエラーが発生しました。"
				pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
				Err.Clear()
				On Error GoTo 0
				GoTo RecordSet_Close
			End If
			On Error GoTo 0
		Loop 
		
		OSC_HYOKAKM_Read = True
		
RecordSet_Close: 
		On Error Resume Next
		dbRec.Close()
		'UPGRADE_NOTE: オブジェクト dbRec をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbRec = Nothing
		On Error GoTo 0
		
	End Function
	
	' *******************************************************************************
	' 概要    : 評価項目マスターのＲＥＡＤを行う。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=READ成功。
	'         :                    False=READ失敗。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2005.12.12  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function OSC_HYOKAKMG_READ() As Short
		
		Dim pRetVal As Short
		Dim pROW As Integer
		Dim pNM As String
		Dim dbRec As ADODB.Recordset
		Dim pMAXLen As Short
		Dim pIX As Short
		Dim pIndex As Short
		Dim pFMT As String
		Dim pMVL As String
		Dim pCOL As Integer
		
		OSC_HYOKAKMG_READ = False
		
		
		
		
		dbRec = New ADODB.Recordset
		
		'
		' 評価項目マスターＲＥＡＤ（概略評価）
		'
		SqlText = "select * "
		SqlText = SqlText & " from OSC_HYOKAKMG "
		SqlText = SqlText & " where HGCDGA  = 1 "
		SqlText = SqlText & "   and HGCDGB  = " & CStr(bCDGB)
		SqlText = SqlText & "   and HGCDGK  = " & CStr(mCDGK)
		SqlText = SqlText & "   and HGNENDO = " & CStr(mNENDO)
		SqlText = SqlText & "   and HGKBSK  = " & CStr(mKBSK)
		SqlText = SqlText & "   and HGYEAR  = " & CStr(mYEAR)
		SqlText = SqlText & "   and HGCDST  = " & CStr(mCDST)
		SqlText = SqlText & "   and HGKBHY  = " & CStr(mKBHY)
		SqlText = SqlText & " order by HGCDGA, HGCDGB, HGCDGK,HGNENDO, HGKBSK, HGYEAR, HGCDST, HGKBHY, HGNOSQ "
		
		On Error Resume Next
		dbRec.Open(SqlText, dbCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター（概略評価）のＲＥＡＤでエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
			Err.Clear()
			On Error GoTo 0
			GoTo RecordSet_Close
		End If
		On Error GoTo 0
		
		If dbRec.EOF = True Then
			GoTo RecordSet_Close
		End If
		
		pROW = -1
		Do While dbRec.EOF = False
			
			pROW = pROW + 1
			PGrid2.set_CellValueByName(pROW, "CDGP", dbRec.Fields("HGCDGP").Value)
			
			PGrid2.set_CellCheckedByName(pROW, "KBDAI", IIf(dbRec.Fields("HGKBDAI").Value = 1, True, False))
			PGrid2.set_CellCheckedByName(pROW, "KBKR", IIf(dbRec.Fields("HGKBKR").Value = 1, True, False))
			PGrid2.set_CellValueByName(pROW, "HYOKA", dbRec.Fields("HGHYOKA").Value)
			'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
			PGrid2.set_CellTextByName(pROW, "NAIYO", IIf(IsDbNull(dbRec.Fields("HGNAIYO").Value), "", dbRec.Fields("HGNAIYO").Value))
			PGrid2.set_CellCheckedByName(pROW, "KB1", IIf(dbRec.Fields("HGKB1").Value = 1, True, False))
			
			Call ENABLE_Change2(pROW)
			
			On Error Resume Next
			dbRec.MoveNext()
			
			If Err.Number <> 0 Then
				mMsgText = "評価項目マスターのＲＥＡＤでエラーが発生しました。"
				pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
				Err.Clear()
				On Error GoTo 0
				GoTo RecordSet_Close
			End If
			On Error GoTo 0
		Loop 
		
		OSC_HYOKAKMG_READ = True
		
RecordSet_Close: 
		On Error Resume Next
		dbRec.Close()
		'UPGRADE_NOTE: オブジェクト dbRec をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbRec = Nothing
		On Error GoTo 0
		
	End Function
	
	
	' *******************************************************************************
	' 概要    : 評価項目マスターのＲＥＡＤを行う。(問題点)
	'         :
	' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=READ成功。
	'         :                    False=READ失敗。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2005.12.12  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function OSC_HYOKAKMM_READ() As Short
		
		Dim pRetVal As Short
		Dim pROW As Integer
		Dim pNM As String
		Dim dbRec As ADODB.Recordset
		Dim pMAXLen As Short
		Dim pIX As Short
		Dim pIndex As Short
		Dim pFMT As String
		Dim pMVL As String
		Dim pCOL As Integer
		
		OSC_HYOKAKMM_READ = False
		
		
		
		
		dbRec = New ADODB.Recordset
		
		'
		' 評価項目マスターＲＥＡＤ(問題点)
		'
		SqlText = "select * "
		SqlText = SqlText & " from OSC_HYOKAKMM "
		SqlText = SqlText & " where HMCDGA  = 1 "
		SqlText = SqlText & "   and HMCDGB  = " & CStr(bCDGB)
		SqlText = SqlText & "   and HMCDGK  = " & CStr(mCDGK)
		SqlText = SqlText & "   and HMNENDO = " & CStr(mNENDO)
		SqlText = SqlText & "   and HMKBSK  = " & CStr(mKBSK)
		SqlText = SqlText & "   and HMYEAR  = " & CStr(mYEAR)
		SqlText = SqlText & "   and HMCDST  = " & CStr(mCDST)
		SqlText = SqlText & "   and HMKBHY  = " & CStr(mKBHY)
		SqlText = SqlText & " order by HMCDGA, HMCDGB, HMCDGK, HMNENDO, HMKBSK, HMYEAR, HMCDST, HMKBHY, HMNOSQ "
		
		On Error Resume Next
		dbRec.Open(SqlText, dbCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター（概略評価）のＲＥＡＤでエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
			Err.Clear()
			On Error GoTo 0
			GoTo RecordSet_Close
		End If
		On Error GoTo 0
		
		If dbRec.EOF = True Then
			GoTo RecordSet_Close
		End If
		
		pROW = -1
		Do While dbRec.EOF = False
			
			pROW = pROW + 1
			PGrid3.set_CellValueByName(pROW, "CDGP", dbRec.Fields("HMCDGP").Value)
			
			PGrid3.set_CellCheckedByName(pROW, "KBDAI", IIf(dbRec.Fields("HMKBDAI").Value = 1, True, False))
			PGrid3.set_CellCheckedByName(pROW, "KBKR", IIf(dbRec.Fields("HMKBKR").Value = 1, True, False))
			'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
			PGrid3.set_CellTextByName(pROW, "MONDAI", IIf(IsDbNull(dbRec.Fields("HMMONDAI").Value), "", dbRec.Fields("HMMONDAI").Value))
			PGrid3.set_CellCheckedByName(pROW, "KB1", IIf(dbRec.Fields("HMKB1").Value = 1, True, False))
			
			
			Call ENABLE_Change3(pROW)
			
			On Error Resume Next
			dbRec.MoveNext()
			
			If Err.Number <> 0 Then
				mMsgText = "評価項目マスターのＲＥＡＤでエラーが発生しました。"
				pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
				Err.Clear()
				On Error GoTo 0
				GoTo RecordSet_Close
			End If
			On Error GoTo 0
		Loop 
		
		OSC_HYOKAKMM_READ = True
		
RecordSet_Close: 
		On Error Resume Next
		dbRec.Close()
		'UPGRADE_NOTE: オブジェクト dbRec をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbRec = Nothing
		On Error GoTo 0
		
	End Function
	
	
	' *******************************************************************************
	' 概要    : 学部別名称マスターのＲＥＡＤを行う。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=READ成功。
	'         :                    False=READ失敗。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2005.12.02  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function KYT_NAMEMGB_Read(ByVal pSYS As Integer, ByVal pKB As Integer, ByVal pCD As Integer, ByRef pNM As String) As Short
		
		Dim pRetVal As Short
		Dim dbRec As ADODB.Recordset
		Dim pMAXLen As Short
		Dim pIX As Short
		
		KYT_NAMEMGB_Read = False
		pNM = ""
		
		dbRec = New ADODB.Recordset
		
		'
		' 学部別名称マスターＲＥＡＤ
		'
		SqlText = "select * "
		SqlText = SqlText & " from KYT_NAMEMGB "
		SqlText = SqlText & " where NMCDGB = " & CStr(bCDGB)
		SqlText = SqlText & "   and NMSYS = " & CStr(pSYS)
		SqlText = SqlText & "   and NMKB = " & CStr(pKB)
		SqlText = SqlText & "   and NMCD = " & CStr(pCD)
		SqlText = SqlText & " order by NMCDGB, NMSYS, NMKB, NMCD "
		
		On Error Resume Next
		dbRec.Open(SqlText, dbCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Err.Number <> 0 Then
			mMsgText = "学部別名称マスターのＲＥＡＤでエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
			Err.Clear()
			On Error GoTo 0
			GoTo RecordSet_Close
		End If
		On Error GoTo 0
		
		If dbRec.EOF = True Then
			GoTo RecordSet_Close
		End If
		
		'略称
		'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
		pNM = IIf(IsDbNull(dbRec.Fields("NMNMR").Value), "", RTrim(dbRec.Fields("NMNMR").Value))
		
		KYT_NAMEMGB_Read = True
		
RecordSet_Close: 
		On Error Resume Next
		dbRec.Close()
		'UPGRADE_NOTE: オブジェクト dbRec をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbRec = Nothing
		On Error GoTo 0
		
	End Function
	
	Public Sub mnuHELPItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHELPItem.Click
		Dim Index As Short = mnuHELPItem.GetIndex(eventSender)
		
		Select Case Index
			Case 0
				'UPGRADE_WARNING: App プロパティ App.EXEName には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
				frmHelpScreen.HelpFileName = My.Application.Info.AssemblyName
				frmHelpScreen.OptionMode = CStr(True)
				frmHelpScreen.Show()
		End Select
		
	End Sub
	
	
	
	Private Sub optKBHY_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optKBHY.Enter
		Dim Index As Short = optKBHY.GetIndex(eventSender)
		
		mLostFocusCheck = True
		
	End Sub
	
	
	Private Sub optKBHY_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optKBHY.Leave
		Dim Index As Short = optKBHY.GetIndex(eventSender)
		
		Dim pRetVal As Short
		
		Select Case FocusMove_HEAD(optKBHY(Index))
			Case "H"
				If Not mLostFocusCheck Then
					Exit Sub
				End If
				
				mLostFocusCheck = False
				Exit Sub
				
			Case "I"
				If InputMode_Check() Then
					Exit Sub
				End If
		End Select
		
	End Sub
	
	
	Private Sub optKBSK_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optKBSK.Enter
		Dim Index As Short = optKBSK.GetIndex(eventSender)
		
		mLostFocusCheck = True
		
	End Sub
	
	
	Private Sub optKBSK_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optKBSK.Leave
		Dim Index As Short = optKBSK.GetIndex(eventSender)
		
		Dim pRetVal As Short
		
		Select Case FocusMove_HEAD(optKBSK(Index))
			Case "H"
				If Not mLostFocusCheck Then
					Exit Sub
				End If
				
				mLostFocusCheck = False
				Exit Sub
				
			Case "I"
				If InputMode_Check() Then
					Exit Sub
				End If
		End Select
		
	End Sub
	
	
	'問題点----------------------------------------------------------------
	Private Sub PGrid3_CellClick(ByVal eventSender As System.Object, ByVal eventArgs As AxPGRIDLib._DPGridEvents_CellClickEvent) Handles PGrid3.CellClick
		
		Dim pRetVal As Short
		Dim pSW_ERROR As Boolean
		Dim pSW_Select As Boolean
		Dim pINCODE As String
		Dim pNMKN As String
		Dim pNMCH As String
		Dim pNMIK As String
		
		'
		' PrefectGrid のサイズ設定中はこのイベントを無視する
		'
		If Not mSW_CellFocusEvent Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		'
		' 行を移動した場合
		'
		If PGrid3.NextRow > (PGrid3.Items - 1) Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If PGrid3.NextRow < 0 Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If eventArgs.Row > (PGrid3.Items - 1) Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If eventArgs.Row < 0 Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If (eventArgs.Row <> PGrid3.NextRow) Then
			'
			' 移動先の行の前の行が空白行かチェックする
			'
			''If Grid1SpaceGyoCheck(PGrid3, PGrid3.NextRow - 1) Then
			If Grid1SpaceGyoCheck(PGrid3, eventArgs.Row) Then
				PGrid3.NextRow = eventArgs.Row
				PGrid3.NextCol = eventArgs.Col
				GoTo PGrid_CellLostFocus_Exit
			End If
		End If
		'
		' 各セルをチェックする
		'
		pSW_ERROR = False
		
		Select Case eventArgs.Col
			Case PGrid3.get_ColOfColName("KBDAI")
				''        '大・中のどちらか
				''PGrid3.CellCheckedByName(Row, "KBCHU") = False
				Call ENABLE_Change3(eventArgs.Row)
				''    Case PGrid3.ColOfColName("KBCHU")
				''        '大・中のどちらか
				''        PGrid3.CellCheckedByName(Row, "KBDAI") = False
				''        Call Button_Change
			Case PGrid3.get_ColOfColName("KBKR")
				Call ENABLE_Change3(eventArgs.Row)
		End Select
		
		' 現在入力中の行の背景色を変更する。 ---------------------------------
		PGrid3.set_RowBackColor(eventArgs.Row, System.Convert.ToUInt32(System.Drawing.Color.White))
		PGrid3.set_RowBackColor(PGrid3.NextRow, System.Convert.ToUInt32(System.Drawing.ColorTranslator.FromOle(&HC0E0FF)))
		
		
		
		'
		' フォーカスをもとに戻す
		'
		mCellClickCol = -1
		
		If pSW_ERROR Then
			PGrid3.NextRow = eventArgs.Row
			PGrid3.NextCol = eventArgs.Col
		End If
		
		GoTo PGrid_CellLostFocus_Exit
		
		'---------------------------------------------
PGrid_CellLostFocus_Exit: 
		
		mSW_CellKeyPress = False
		
	End Sub
	
	
	'
	Private Sub PGrid3_CellLostFocus(ByVal eventSender As System.Object, ByVal eventArgs As AxPGRIDLib._DPGridEvents_CellLostFocusEvent) Handles PGrid3.CellLostFocus
		
		Dim pRetVal As Short
		Dim pSW_ERROR As Boolean
		Dim pCOL As Integer
		Dim pCellText As String
		Dim pJU As String
		Dim pNMKN As String
		Dim pNMCH As String
		Dim pNMIK As String
		Dim pTIME As String
		Dim pLEN As Short
		
		
		'
		' PrefectGrid のサイズ設定中はこのイベントを無視する
		'
		If Not mSW_CellFocusEvent Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		'
		' 行を移動した場合
		'
		If PGrid3.NextRow > (PGrid3.Items - 1) Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If (eventArgs.Row <> PGrid3.NextRow) Then
			'
			' 移動先の行の前の行が空白行かチェックする
			'
			If Grid1SpaceGyoCheck(PGrid3, PGrid3.NextRow - 1) Then
				PGrid3.NextRow = eventArgs.Row
				PGrid3.NextCol = eventArgs.Col
				GoTo PGrid_CellLostFocus_Exit
			End If
		End If
		'
		' 各セルをチェックする
		'
		pSW_ERROR = False
		
		' 現在入力中の行の背景色を変更する。 ---------------------------------
		PGrid3.set_RowBackColor(eventArgs.Row, System.Convert.ToUInt32(System.Drawing.Color.White))
		'PGrid3.RowBackColor(PGrid3.NextRow) = &HC0E0FF
		
		Select Case eventArgs.Col
			'' ｺｰﾄﾞ -----------------------------------------------------------
			Case PGrid3.get_ColOfColName("CDGP")
				If PGrid3.get_CellTextByName(eventArgs.Row, "CDGP") <> "" And PGrid3.get_CellCheckedByName(eventArgs.Row, "KBKR") = False Then
					If Not CellCheck_Numeric(PGrid3, eventArgs.Row, eventArgs.Col) Then
						pSW_ERROR = True
					End If
					'前の行との大小をcheck
					If Not CDGP_Check(PGrid3, eventArgs.Row, PGrid3.get_CellValue(eventArgs.Row, eventArgs.Col)) Then
						pSW_ERROR = True
					End If
				End If
				''        ' 評価 -----------------------------------------------------------
				''        Case PGrid3.ColOfColName("HYOKA")
				''            If PGrid3.CellTextByName(Row, "HYOKA") <> "" And PGrid3.CellCheckedByName(Row, "KBKR") = False Then
				''                If Not CellCheck_Numeric(PGrid, Row, Col) Then
				''                    pSW_ERROR = True
				''                End If
				''
				''            End If
				'' 内容-----------------------------------------------------------
			Case PGrid3.get_ColOfColName("MONDAI")
				If mKBHY = 0 Then
					If PGrid3.get_CellCheckedByName(eventArgs.Row, "KBDAI") = True Then
						PGrid3.set_CellText(eventArgs.Row, eventArgs.Col, MOJIByte_Get(PGrid3.get_CellText(eventArgs.Row, eventArgs.Col), Me, 1, 60, pLEN))
					Else
						PGrid3.set_CellText(eventArgs.Row, eventArgs.Col, MOJIByte_Get(PGrid3.get_CellText(eventArgs.Row, eventArgs.Col), Me, 1, 44, pLEN))
					End If
				ElseIf mKBHY = 1 Then 
					PGrid3.set_CellText(eventArgs.Row, eventArgs.Col, MOJIByte_Get(PGrid3.get_CellText(eventArgs.Row, eventArgs.Col), Me, 1, 80, pLEN))
				End If
				
		End Select
		'
		' フォーカスをもとに戻す
		'
		mCellClickCol = -1
		
		If pSW_ERROR Then
			PGrid3.NextRow = eventArgs.Row
			PGrid3.NextCol = eventArgs.Col
		End If
		
		GoTo PGrid_CellLostFocus_Exit
		
		'---------------------------------------------
PGrid_CellLostFocus_Exit: 
		
		mSW_CellKeyPress = False
		
	End Sub
	
	
	
	
	Private Sub Toolbar1_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _Toolbar1_Button1.Click, _Toolbar1_Button2.Click, _Toolbar1_Button3.Click, _Toolbar1_Button4.Click, _Toolbar1_Button5.Click, _Toolbar1_Button6.Click, _Toolbar1_Button7.Click, _Toolbar1_Button8.Click, _Toolbar1_Button9.Click
		Dim Button As System.Windows.Forms.ToolStripItem = CType(eventSender, System.Windows.Forms.ToolStripItem)
		
		Dim pRetVal As Short
		
		Select Case Button.Name
			Case "EXEC"
				Call mnuFILEItem_Click(mnuFILEItem.Item(0), New System.EventArgs())
				
			Case "CANCEL"
				Call mnuFILEItem_Click(mnuFILEItem.Item(1), New System.EventArgs())
				
			Case "PRINT"
				Call mnuFILEItem_Click(mnuFILEItem.Item(5), New System.EventArgs())
				
			Case "EXIT"
				Call mnuFILEItem_Click(mnuFILEItem.Item(9), New System.EventArgs())
				
			Case "ROWINSERT"
				Call mnuEDITItem_Click(mnuEDITItem.Item(0), New System.EventArgs())
				
			Case "ROWDELETE"
				Call mnuEDITItem_Click(mnuEDITItem.Item(1), New System.EventArgs())
				
			Case "COPY"
				Call mnuEDITItem_Click(mnuEDITItem.Item(9), New System.EventArgs())
				
		End Select
		
	End Sub
	
	' *******************************************************************************
	' 概要  : 保存ボタンが押されたときの処理を行う。
	' 　　  :
	' 引数  : 戻り値, O, Integer, True =正常。
	' 　　  : 　　　              False=異常。
	' 　　  :
	' 説明  :
	' 　　  :
	' 履歴  : 2005.12.12  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function Kosin_Exec() As Short
		
		Dim pRetVal As Short
		Dim pROW As Integer
		Dim pROW2 As Integer
		Dim pIX As Integer
		Dim pIIX As Integer
		Dim pCODE As String
		Dim pCODE2 As String
		Dim pCDTA1 As Short
		Dim pCDTA2 As Short
		Dim pHIT As Boolean
		Dim pNM As String
		Dim pYEAR As String
		Dim pYEAR2 As String
		Dim pCDKJ As Integer
		Dim pCDKJ2 As Integer
		
		Dim pTIME As String
		Dim pTIME2 As String
		
		Dim pCDGP As Integer
		Dim pCDGP2 As Integer
		
		Kosin_Exec = False
		
		''----------------------------
		' マスターファイルを更新する
		''----------------------------
		'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		
		pHIT = False
		pCDGP = 0
		pCDGP2 = 0
		
		'メイングリッド=======================================================
		For pROW = 0 To (PGrid.Items - 1)
			If Not Grid1SpaceGyoCheck(PGrid, pROW) Then
				If PGrid.get_CellCheckedByName(pROW, "KBKR") = False Then
					
					pRetVal = PGrid.SelectCell(pROW, PGrid.get_ColOfColName("CDGP"))
					
					'' ----- 入力ﾁｪｯｸ
					If PGrid.get_CellTextByName(pROW, "CDGP") = "" Then
						mMsgText = CStr(pROW + 1) & "行目のグループが入力されていません。" & Chr(10)
						mMsgText = mMsgText & "再度確認して下さい。"
						pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Information, mMBOXTitle)
						Call cmdCLOSE_ClickEvent(cmdCLOSE.Item(0), New System.EventArgs())
						pRetVal = PGrid.SelectCell(pROW, PGrid.get_ColOfColName("CDGP"))
						'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
						System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
						Exit Function
					End If
					
					'' ----- 数字ﾁｪｯｸ
					If Not CellCheck_Numeric(PGrid, pROW, PGrid.get_ColOfColName("CDGP")) Then
						mMsgText = CStr(pROW + 1) & "行目のグループの入力が正しくありません。" & Chr(10)
						mMsgText = mMsgText & "再度確認してください。"
						pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Information, mMBOXTitle)
						Call cmdCLOSE_ClickEvent(cmdCLOSE.Item(0), New System.EventArgs())
						pRetVal = PGrid.SelectCell(pROW, PGrid.get_ColOfColName("CDGP"))
						'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
						System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
						Exit Function
					End If
					
					pCDGP = PGrid.get_CellValueByName(pROW, "CDGP")
					
					
					'前の行より小さいとerror
					If pROW <> 0 And pCDGP <> 0 Then
						If pCDGP < pCDGP2 Then
							mMsgText = CStr(pROW + 1) & "行目のグループの入力に誤りがあります。" & Chr(10)
							mMsgText = mMsgText & "前のグループ以上の値を入力して下さい。"
							pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Information, mMBOXTitle)
							Call cmdCLOSE_ClickEvent(cmdCLOSE.Item(0), New System.EventArgs())
							pRetVal = PGrid.SelectCell(pROW, PGrid.get_ColOfColName("CDGP"))
							'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
							System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
							Exit Function
						End If
					End If
					
					pHIT = True
					pCDGP2 = pCDGP
					
					'解答数の数字チェック
					If Not CellCheck_Numeric(PGrid, pROW, PGrid.get_ColOfColName("SUKA")) Then
						mMsgText = CStr(pROW + 1) & "行目の解答数の入力が正しくありません。" & Chr(10)
						mMsgText = mMsgText & "再度確認してください。"
						pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Information, mMBOXTitle)
						Call cmdCLOSE_ClickEvent(cmdCLOSE.Item(0), New System.EventArgs())
						pRetVal = PGrid.SelectCell(pROW, PGrid.get_ColOfColName("SUKA"))
						'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
						System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
						Exit Function
					End If
				End If
			End If
		Next pROW
		
		pHIT = False
		pCDGP = 0
		pCDGP2 = 0
		
		'概略評価グリッド=======================================================
		For pROW = 0 To (PGrid2.Items - 1)
			If Not Grid1SpaceGyoCheck(PGrid2, pROW) Then
				If PGrid2.get_CellCheckedByName(pROW, "KBKR") = False Then
					
					pRetVal = PGrid2.SelectCell(pROW, PGrid2.get_ColOfColName("CDGP"))
					
					'' ----- 入力ﾁｪｯｸ
					If PGrid2.get_CellTextByName(pROW, "CDGP") = "" Then
						mMsgText = "概略評価　" & CStr(pROW + 1) & "行目のグループが入力されていません。" & Chr(10)
						mMsgText = mMsgText & "再度確認して下さい。"
						pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Information, mMBOXTitle)
						Call cmdDISP2_ClickEvent(cmdDISP2, New System.EventArgs())
						pRetVal = PGrid2.SelectCell(pROW, PGrid2.get_ColOfColName("CDGP"))
						'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
						System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
						Exit Function
					End If
					
					'' ----- 数字ﾁｪｯｸ
					If Not CellCheck_Numeric(PGrid2, pROW, PGrid2.get_ColOfColName("CDGP")) Then
						mMsgText = "概略評価　" & CStr(pROW + 1) & "行目のグループの入力が正しくありません。" & Chr(10)
						mMsgText = mMsgText & "再度確認してください。"
						pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Information, mMBOXTitle)
						Call cmdDISP2_ClickEvent(cmdDISP2, New System.EventArgs())
						pRetVal = PGrid2.SelectCell(pROW, PGrid2.get_ColOfColName("CDGP"))
						'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
						System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
						Exit Function
					End If
					
					pCDGP = PGrid2.get_CellValueByName(pROW, "CDGP")
					
					
					'前の行より小さいとerror
					If pROW <> 0 And pCDGP <> 0 Then
						If pCDGP < pCDGP2 Then
							mMsgText = "概略評価　" & CStr(pROW + 1) & "行目のグループの入力に誤りがあります。" & Chr(10)
							mMsgText = mMsgText & "前のグループ以上の値を入力して下さい。"
							pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Information, mMBOXTitle)
							Call cmdDISP2_ClickEvent(cmdDISP2, New System.EventArgs())
							pRetVal = PGrid2.SelectCell(pROW, PGrid2.get_ColOfColName("CDGP"))
							'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
							System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
							Exit Function
						End If
					End If
					
					pHIT = True
					pCDGP2 = pCDGP
					
					'評価の数字チェック
					If Not CellCheck_Numeric(PGrid2, pROW, PGrid2.get_ColOfColName("HYOKA")) Then
						mMsgText = "概略評価　" & CStr(pROW + 1) & "行目の評価の入力が正しくありません。" & Chr(10)
						mMsgText = mMsgText & "再度確認してください。"
						pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Information, mMBOXTitle)
						Call cmdDISP2_ClickEvent(cmdDISP2, New System.EventArgs())
						pRetVal = PGrid2.SelectCell(pROW, PGrid2.get_ColOfColName("HYOKA"))
						'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
						System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
						Exit Function
					End If
				End If
			End If
		Next pROW
		
		pHIT = False
		pCDGP = 0
		pCDGP2 = 0
		
		'問題点グリッド=======================================================
		For pROW = 0 To (PGrid3.Items - 1)
			If Not Grid1SpaceGyoCheck(PGrid3, pROW) Then
				If PGrid3.get_CellCheckedByName(pROW, "KBKR") = False Then
					
					pRetVal = PGrid3.SelectCell(pROW, PGrid3.get_ColOfColName("CDGP"))
					
					'' ----- 入力ﾁｪｯｸ
					If PGrid3.get_CellTextByName(pROW, "CDGP") = "" Then
						mMsgText = "問題点　" & CStr(pROW + 1) & "行目のグループが入力されていません。" & Chr(10)
						mMsgText = mMsgText & "再度確認して下さい。"
						pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Information, mMBOXTitle)
						Call cmdDISP3_ClickEvent(cmdDISP3, New System.EventArgs())
						pRetVal = PGrid3.SelectCell(pROW, PGrid3.get_ColOfColName("CDGP"))
						'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
						System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
						Exit Function
					End If
					
					'' ----- 数字ﾁｪｯｸ
					If Not CellCheck_Numeric(PGrid3, pROW, PGrid3.get_ColOfColName("CDGP")) Then
						mMsgText = "問題点　" & CStr(pROW + 1) & "行目のグループの入力が正しくありません。" & Chr(10)
						mMsgText = mMsgText & "再度確認してください。"
						pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Information, mMBOXTitle)
						Call cmdDISP3_ClickEvent(cmdDISP3, New System.EventArgs())
						pRetVal = PGrid3.SelectCell(pROW, PGrid3.get_ColOfColName("CDGP"))
						'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
						System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
						Exit Function
					End If
					
					pCDGP = PGrid3.get_CellValueByName(pROW, "CDGP")
					
					'前の行より小さいとerror
					If pROW <> 0 And pCDGP <> 0 Then
						If pCDGP < pCDGP2 Then
							mMsgText = "問題点　" & CStr(pROW + 1) & "行目のグループの入力に誤りがあります。" & Chr(10)
							mMsgText = mMsgText & "前のグループ以上の値を入力して下さい。"
							pRetVal = MsgBox(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Information, mMBOXTitle)
							Call cmdDISP3_ClickEvent(cmdDISP3, New System.EventArgs())
							pRetVal = PGrid3.SelectCell(pROW, PGrid3.get_ColOfColName("CDGP"))
							'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
							System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
							Exit Function
						End If
					End If
					
					pHIT = True
					pCDGP2 = pCDGP
					
				End If
			End If
		Next pROW
		
		Select Case mInputMode
			Case "E", "U" : pRetVal = Table_Insert()
			Case Else : pRetVal = False
		End Select
		
		If Not pRetVal Then
			'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			Exit Function
		End If
		
		Call Screen_Clear()
		inpCDST.Focus()
		
		'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
		Kosin_Exec = True
		
	End Function
	
	
	Public Sub mnuEDITItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEDITItem.Click
		Dim Index As Short = mnuEDITItem.GetIndex(eventSender)
		
		Dim pROW As Integer
		Dim pRetVal As Short
		Dim pIX As Integer
		
		Select Case Index
			Case 0 ''行挿入
				
				
				If picGAIRYAKU.Visible = True Then
					pROW = PGrid2.Row
					
					pRetVal = MsgBox(CStr(pROW + 1) & "行目に挿入しますか？", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation + MsgBoxStyle.DefaultButton2, mMBOXTitle)
					
					If pRetVal <> MsgBoxResult.Yes Then
						Exit Sub
					End If
					
					If Not Grid1SpaceGyoCheck(PGrid2, PGrid2.Items - 1) Then
						mMsgText = "最終行のデータが削除されます。"
						mMsgText = mMsgText & Chr(10) & "行を挿入しますか？"
						pRetVal = MsgBox(mMsgText, MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation + MsgBoxStyle.DefaultButton2, mMBOXTitle)
						
						If pRetVal = MsgBoxResult.No Then
							Exit Sub
						End If
						
					End If
					
					PGrid2.RefreshLater = True
					
					PGrid2.RemoveItems(PGrid2.Items - 1, 1)
					'UPGRADE_NOTE: Text は CtlText にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
					PGrid2.CtlText = PGrid2.get_CellText(PGrid2.Row, PGrid2.Col)
					
					PGrid2.TextAtAddItem = ""
					PGrid2.AddItems(pROW, 1)
					
					For pIX = 0 To (PGrid2.Items - 1)
						PGrid2.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
					Next pIX
					
					PGrid2.NextRow = pROW
					PGrid2.NextCol = PGrid2.Col
					
					PGrid2.RefreshLater = False
					'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
					PGrid2.CtlRefresh()
					
					Call ENABLE_Change2(pROW)
					
				ElseIf picMONDAI.Visible = True Then 
					pROW = PGrid3.Row
					
					pRetVal = MsgBox(CStr(pROW + 1) & "行目に挿入しますか？", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation + MsgBoxStyle.DefaultButton2, mMBOXTitle)
					
					If pRetVal <> MsgBoxResult.Yes Then
						Exit Sub
					End If
					
					If Not Grid1SpaceGyoCheck(PGrid3, PGrid3.Items - 1) Then
						mMsgText = "最終行のデータが削除されます。"
						mMsgText = mMsgText & Chr(10) & "行を挿入しますか？"
						pRetVal = MsgBox(mMsgText, MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation + MsgBoxStyle.DefaultButton2, mMBOXTitle)
						
						If pRetVal = MsgBoxResult.No Then
							Exit Sub
						End If
						
					End If
					
					PGrid3.RefreshLater = True
					
					PGrid3.RemoveItems(PGrid3.Items - 1, 1)
					'UPGRADE_NOTE: Text は CtlText にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
					PGrid3.CtlText = PGrid3.get_CellText(PGrid3.Row, PGrid3.Col)
					
					PGrid3.TextAtAddItem = ""
					PGrid3.AddItems(pROW, 1)
					
					For pIX = 0 To (PGrid3.Items - 1)
						PGrid3.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
					Next pIX
					
					PGrid3.NextRow = pROW
					PGrid3.NextCol = PGrid3.Col
					
					PGrid3.RefreshLater = False
					'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
					PGrid3.CtlRefresh()
					
					Call ENABLE_Change3(pROW)
				Else
					pROW = PGrid.Row
					
					pRetVal = MsgBox(CStr(pROW + 1) & "行目に挿入しますか？", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation + MsgBoxStyle.DefaultButton2, mMBOXTitle)
					
					If pRetVal <> MsgBoxResult.Yes Then
						Exit Sub
					End If
					
					If Not Grid1SpaceGyoCheck(PGrid, PGrid.Items - 1) Then
						mMsgText = "最終行のデータが削除されます。"
						mMsgText = mMsgText & Chr(10) & "行を挿入しますか？"
						pRetVal = MsgBox(mMsgText, MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation + MsgBoxStyle.DefaultButton2, mMBOXTitle)
						
						If pRetVal = MsgBoxResult.No Then
							Exit Sub
						End If
						
					End If
					
					PGrid.RefreshLater = True
					
					PGrid.RemoveItems(PGrid.Items - 1, 1)
					'UPGRADE_NOTE: Text は CtlText にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
					PGrid.CtlText = PGrid.get_CellText(PGrid.Row, PGrid.Col)
					
					PGrid.TextAtAddItem = ""
					PGrid.AddItems(pROW, 1)
					
					For pIX = 0 To (PGrid.Items - 1)
						PGrid.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
					Next pIX
					
					PGrid.NextRow = pROW
					PGrid.NextCol = PGrid.Col
					
					PGrid.RefreshLater = False
					'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
					PGrid.CtlRefresh()
					
					Call ENABLE_Change(pROW)
				End If
				
			Case 1 ' 行削除
				''            If picGAIRYAKU.Visible = True Then
				''                pPGrid = PGrid2
				''            ElseIf picMONDAI.Visible = True Then
				''                pPGrid = PGrid3
				''            Else
				''                pPGrid = PGrid
				''            End If
				
				If picGAIRYAKU.Visible = True Then
					pROW = PGrid2.Row
					
					pRetVal = MsgBox(CStr(pROW + 1) & "行目を削除しますか？", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.Exclamation, mMBOXTitle)
					
					If pRetVal = MsgBoxResult.Yes Then
						pRetVal = PGrid2.RemoveItems(PGrid2.Row, 1)
						If PGrid2.Row <> 32700 Then '次の行が空行だとこうなる。
							'UPGRADE_NOTE: Text は CtlText にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
							PGrid2.CtlText = PGrid2.get_CellText(PGrid2.Row, PGrid2.Col)
						End If
						pRetVal = PGrid2.AddItems(PGrid2.Items, 1)
						
						
						For pIX = 0 To (PGrid2.Items - 1)
							PGrid2.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
							'                    PGrid.CellTextByName(pIX, "cmdNORM") = "参照"
							'                    PGrid.CellTextByName(pIX, "cmdCDTO") = "参照"
						Next pIX
						
						'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
						PGrid2.CtlRefresh()
						If PGrid2.get_CellEnabled(pROW, PGrid2.Col) = True Then
							pRetVal = PGrid2.SelectCell(pROW, PGrid2.Col)
						End If
						Call ENABLE_Change2(pROW)
					End If
					
					
				ElseIf picMONDAI.Visible = True Then 
					pROW = PGrid3.Row
					
					pRetVal = MsgBox(CStr(pROW + 1) & "行目を削除しますか？", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.Exclamation, mMBOXTitle)
					
					If pRetVal = MsgBoxResult.Yes Then
						pRetVal = PGrid3.RemoveItems(PGrid3.Row, 1)
						If PGrid3.Row <> 32700 Then '次の行が空行だとこうなる。
							'UPGRADE_NOTE: Text は CtlText にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
							PGrid3.CtlText = PGrid3.get_CellText(PGrid3.Row, PGrid3.Col)
						End If
						pRetVal = PGrid3.AddItems(PGrid3.Items, 1)
						
						For pIX = 0 To (PGrid3.Items - 1)
							PGrid3.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
							'                    PGrid.CellTextByName(pIX, "cmdNORM") = "参照"
							'                    PGrid.CellTextByName(pIX, "cmdCDTO") = "参照"
						Next pIX
						
						'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
						PGrid3.CtlRefresh()
						If PGrid3.get_CellEnabled(pROW, PGrid3.Col) = True Then
							pRetVal = PGrid3.SelectCell(pROW, PGrid3.Col)
						End If
						Call ENABLE_Change3(pROW)
					End If
					
				Else
					pROW = PGrid.Row
					
					pRetVal = MsgBox(CStr(pROW + 1) & "行目を削除しますか？", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.Exclamation, mMBOXTitle)
					
					If pRetVal = MsgBoxResult.Yes Then
						pRetVal = PGrid.RemoveItems(PGrid.Row, 1)
						
						If PGrid.Row <> 32700 Then '次の行が空行だとこうなる。
							'UPGRADE_NOTE: Text は CtlText にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
							PGrid.CtlText = PGrid.get_CellText(PGrid.Row, PGrid.Col)
						End If
						pRetVal = PGrid.AddItems(PGrid.Items, 1)
						
						For pIX = 0 To (PGrid.Items - 1)
							PGrid.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
							'                    PGrid.CellTextByName(pIX, "cmdNORM") = "参照"
							'                    PGrid.CellTextByName(pIX, "cmdCDTO") = "参照"
						Next pIX
						
						'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
						PGrid.CtlRefresh()
						If PGrid.get_CellEnabled(pROW, PGrid.Col) = True Then
							pRetVal = PGrid.SelectCell(pROW, PGrid.Col)
						End If
						Call ENABLE_Change(pROW)
					End If
					
				End If
				
			Case 9 ' 一括ｺﾋﾟｰ
				
				Me.Enabled = False
				frmPK9OSCM004mk.lblNMGB.Text = lblNMGB.Text
				frmPK9OSCM004mk.inpSTNENDO.Value = mNENDO
				frmPK9OSCM004mk.Show()
				
				
		End Select
		
	End Sub
	
	' *******************************************************************************
	' 概要    : キャンセルボタンが押されたときの処理を行う。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : pSW_DEL I,Boolean, 削除ＳＷ　True=削除モード。
	'         :                             False=追加、登録モード。
	'         : 戻り値, O, Integer, True=正常。
	'         :                    False=異常。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2003.05.09  REV.0001  植田  新規作成。
	' *******************************************************************************
	'
	Private Function Kosin_Cancel(ByVal pSW_DEL As Boolean) As Short
		
		Dim pRetVal As Short
		
		Kosin_Cancel = False
		
		Select Case pSW_DEL
			Case False : pRetVal = MsgBox("入力がキャンセルされました。", MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, mMBOXTitle)
			Case True : pRetVal = MsgBox("削除がキャンセルされました。", MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, mMBOXTitle)
		End Select
		
		Call Screen_Clear()
		inpCDST.Focus()
		
		Kosin_Cancel = True
		
	End Function
	
	' *******************************************************************************
	' 概要    : 削除ボタンが押されたときの処理を行う。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=正常。
	'         :                    False=異常。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2005.12.02  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function Kosin_Delete() As Short
		
		Dim pRetVal As Short
		
		Kosin_Delete = False
		
		pRetVal = MsgBox("データを削除します。よろしいですか？", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation + MsgBoxStyle.DefaultButton2, mMBOXTitle)
		
		If pRetVal <> MsgBoxResult.Yes Then
			pRetVal = Kosin_Cancel(True)
			Kosin_Delete = True
			Exit Function
		End If
		'
		' マスターファイルを削除する
		'
		'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		
		mInputMode = "D"
		
		If Not Table_Insert() Then
			'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			Exit Function
		End If
		
		Call Screen_Clear()
		inpCDST.Focus()
		
		'UPGRADE_WARNING: App プロパティ App.EXEName には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
		Call MenuLOGF_Insert(My.Application.Info.AssemblyName, FormCaption, bDEL, bIDUS, bNMUS)
		
		'UPGRADE_WARNING: Screen プロパティ Screen.MousePointer には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
		Kosin_Delete = True
		
	End Function
	
	
	' *******************************************************************************
	' 概要    : ＰｅｒｆｅｃｔＧｒｉｄの行が空白かどうかをチェックする。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : pRow, I, Long, 行番号。
	'         : 戻り値, O, Integer, True=空白あり。
	'         :                     False=空白なし。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2003.05.09  REV.0001  植田  新規作成。
	' *******************************************************************************
	'
	Private Function Grid1SpaceGyoCheck(ByVal pPGrid As AxPGRIDLib.AxPerfectGrid, ByVal pROW As Short) As Short
		
		Dim pCOL As Integer
		
		Grid1SpaceGyoCheck = False
		
		If (pROW < 0) Or (pROW > pPGrid.Items - 1) Then
			Exit Function
		End If
		
		For pCOL = 0 To (pPGrid.Cols - 1)
			
			If pPGrid.get_ColStyle(pCOL) <> PGRIDLib.ColStyleConstants.pgcs_コマンドボタン Then
				If pPGrid.get_ColStyle(pCOL) = PGRIDLib.ColStyleConstants.pgcs_チェックボックス Then '2005.12.01 add
					If pPGrid.get_CellChecked(pROW, pCOL) = True Then
						Exit Function
					End If
				Else
					
					'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
					If Not (IsDbNull(pPGrid.get_CellText(pROW, pCOL)) Or Trim(pPGrid.get_CellText(pROW, pCOL)) = "") Then
						Exit Function
					End If
				End If
			End If
		Next pCOL
		
		Grid1SpaceGyoCheck = True
		
	End Function
	
	
	' *******************************************************************************
	' 概要    : ＰｅｒｆｅｃｔＧｒｉｄのアイテムをクリアーする。
	'         :
	' ﾊﾟﾗﾒｰﾀ  :
	'         :
	' 説明    :
	'         :
	' 履歴    : 2004.01.16  REV.0001  赤池  新規作成。
	' *******************************************************************************
	'
	Private Sub GridClear()
		
		Dim pROW As Integer
		Dim pIX As Integer
		
		PGrid.RefreshLater = True
		pROW = PGrid.Items
		PGrid.RemoveItems(0, pROW)
		PGrid.TextAtAddItem = ""
		PGrid.AddItems(0, pROW)
		
		For pIX = 0 To (PGrid.Items - 1)
			PGrid.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
		Next pIX
		
		PGrid.RefreshLater = False
		'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		PGrid.CtlRefresh()
		
	End Sub
	
	
	
	' *******************************************************************************
	' 概要    : ＰｅｒｆｅｃｔＧｒｉｄのアイテムをクリアーする。
	'         :
	' ﾊﾟﾗﾒｰﾀ  :
	'         :
	' 説明    :
	'         :
	' 履歴    : 2004.01.16  REV.0001  赤池  新規作成。
	' *******************************************************************************
	'
	Private Sub GridClear2()
		
		Dim pROW As Integer
		Dim pIX As Integer
		
		PGrid2.RefreshLater = True
		pROW = PGrid2.Items
		PGrid2.RemoveItems(0, pROW)
		PGrid2.TextAtAddItem = ""
		PGrid2.AddItems(0, pROW)
		
		For pIX = 0 To (PGrid2.Items - 1)
			PGrid2.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
		Next pIX
		
		PGrid2.RefreshLater = False
		'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		PGrid2.CtlRefresh()
		
	End Sub
	
	' *******************************************************************************
	' 概要    : ＰｅｒｆｅｃｔＧｒｉｄのアイテムをクリアーする。
	'         :
	' ﾊﾟﾗﾒｰﾀ  :
	'         :
	' 説明    :
	'         :
	' 履歴    : 2004.01.16  REV.0001  赤池  新規作成。
	' *******************************************************************************
	'
	Private Sub GridClear3()
		
		Dim pROW As Integer
		Dim pIX As Integer
		
		PGrid.RefreshLater = True
		pROW = PGrid3.Items
		PGrid3.RemoveItems(0, pROW)
		PGrid3.TextAtAddItem = ""
		PGrid3.AddItems(0, pROW)
		
		For pIX = 0 To (PGrid3.Items - 1)
			PGrid3.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
		Next pIX
		
		PGrid3.RefreshLater = False
		'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		PGrid3.CtlRefresh()
		
	End Sub
	
	
	' *******************************************************************************
	' 概要 : Ｐｅｒｆｅｃｔ Ｇｒｉｄの設定。
	' 　　 :
	' 説明 :
	' 　　 :
	' 履歴 : 2005.12.02  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Sub Grid_Resize()
		
		Dim pRetVal As Short
		Dim pFONT As System.Drawing.Font = System.Windows.Forms.Control.DefaultFont.Clone()
		Dim pII As Integer
		Dim pCOL As Integer
		Dim pCol2 As Integer
		Dim pRTN As Short
		Dim pIX As Short
		Dim pWW As Integer
		
		pFONT = VB6.FontChangeName(pFONT, "ＭＳ ゴシック")
		pFONT = VB6.FontChangeSize(pFONT, 8)
		
		' 全体プロパティの設定
		PGrid.BorderWidth = 1 '境界線あり
		PGrid.ExitCol = 1 '右列へ
		PGrid.VertBorderWidth = 1 '列境界線幅１
		PGrid.AdjustColWidth = False '列幅補正なし
		PGrid.OddRowBorderWidth = 1 ' 遇数行と奇数行の境界
		PGrid.OddRowBorder = 3 ' 点線
		PGrid.OddRowBorderColor = &HFFC0C0 ' 淡い紫
		PGrid.OddRowMeansLogical = False ' 奇数行は表示上
		PGrid.SepAlwaysDraw = True
		'    PGrid.ForeColor = vbBlue
		
		'    PGrid.FoldCol = 6
		PGrid.ThruStartShift = 1
		PGrid.ThruStartVKeyName = "VK_INSERT"
		
		PGrid.AllowSelCol = False
		PGrid.AllowSelRow = False
		PGrid.HeightSizing = 0
		PGrid.WidthSizing = 0
		
		PGrid.set_ColWidth(-1, 30) '連番
		PGrid.set_CellText(-1, -1, "行")
		PGrid.set_ColAlignmentH(-1, PGRIDLib.ColAlignmentHConstants.pgcah_右揃え)
		'セルプロパティの設定
		
		pCOL = 0
		PGrid.set_CellText(-1, pCOL, "ｸﾞﾙｰﾌﾟ")
		PGrid.set_CellAlignmentH(-1, pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid.set_ColName(pCOL, "CDGP")
		PGrid.set_ColWidth(pCOL, 50)
		PGrid.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_右揃え)
		PGrid.set_ColMaxLengthB(pCOL, 2)
		PGrid.set_ColMinValue(pCOL, 0)
		PGrid.set_ColMaxValue(pCOL, 99)
		PGrid.set_ColFormatString(pCOL, "00")
		PGrid.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		PGrid.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_テキスト)
		
		pCOL = pCOL + 1
		PGrid.set_CellText(-1, pCOL, "大")
		PGrid.set_ColName(pCOL, "KBDAI")
		PGrid.set_ColWidth(pCOL, 20)
		PGrid.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_チェックボックス)
		PGrid.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		PGrid.set_CellText(-1, pCOL, "中")
		PGrid.set_ColName(pCOL, "KBCHU")
		PGrid.set_ColWidth(pCOL, 20)
		PGrid.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_チェックボックス)
		PGrid.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		PGrid.set_CellText(-1, pCOL, "空")
		PGrid.set_ColName(pCOL, "KBKR")
		PGrid.set_ColWidth(pCOL, 20)
		PGrid.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_チェックボックス)
		PGrid.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		PGrid.set_CellText(-1, pCOL, "設  問")
		PGrid.set_CellAlignmentH(-1, pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_左揃え)
		PGrid.set_ColName(pCOL, "MONDAI")
		PGrid.set_ColWidth(pCOL, 610)
		PGrid.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_左揃え)
		PGrid.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_テキスト)
		PGrid.set_ColMaxLengthB(pCOL, 80)
		PGrid.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_全角ひらがな)
		
		pCOL = pCOL + 1
		PGrid.set_CellText(-1, pCOL, "番号")
		PGrid.set_CellAlignmentH(-1, pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid.set_ColName(pCOL, "NO")
		PGrid.set_ColWidth(pCOL, 30)
		PGrid.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_右揃え)
		PGrid.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_テキスト)
		PGrid.set_ColMaxLengthB(pCOL, 2)
		'PGrid.ColMinValue(pCOL) = 0
		'PGrid.ColMaxValue(pCOL) = 3
		PGrid.set_ColFormatString(pCOL, "#")
		PGrid.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		PGrid.set_CellText(-1, pCOL, "解答数")
		PGrid.set_CellAlignmentH(-1, pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid.set_ColName(pCOL, "SUKA")
		PGrid.set_ColWidth(pCOL, 50)
		PGrid.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_右揃え)
		PGrid.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_テキスト)
		PGrid.set_ColMaxLengthB(pCOL, 2)
		PGrid.set_ColMinValue(pCOL, 0)
		'2009.06.03 upd ---
		If bCDGB = 3 Then
			PGrid.set_ColMaxValue(pCOL, 2)
		Else
			PGrid.set_ColMaxValue(pCOL, 4)
		End If
		'------------------
		PGrid.set_ColFormatString(pCOL, "#")
		PGrid.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		'2009.06.03 upd ---
		If bCDGB = 3 Then
			PGrid.set_CellText(-1, pCOL, "")
			PGrid.set_ColWidth(pCOL, 0)
		Else
			PGrid.set_CellText(-1, pCOL, "３")
			PGrid.set_ColWidth(pCOL, 35)
		End If
		'------------------
		PGrid.set_ColName(pCOL, "KB1")
		PGrid.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_チェックボックス)
		PGrid.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		'2009.06.03 upd ---
		If bCDGB = 3 Then
			PGrid.set_CellText(-1, pCOL, "")
			PGrid.set_ColWidth(pCOL, 0)
		Else
			PGrid.set_CellText(-1, pCOL, "２")
			PGrid.set_ColWidth(pCOL, 35)
		End If
		'------------------
		PGrid.set_ColName(pCOL, "KB2")
		PGrid.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_チェックボックス)
		PGrid.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		'2009.06.03 upd ---
		If bCDGB = 3 Then
			PGrid.set_CellText(-1, pCOL, "はい")
			PGrid.set_ColWidth(pCOL, 70)
		Else
			PGrid.set_CellText(-1, pCOL, "１")
			PGrid.set_ColWidth(pCOL, 35)
		End If
		'------------------
		PGrid.set_ColName(pCOL, "KB3")
		PGrid.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_チェックボックス)
		PGrid.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		'2009.06.03 upd ---
		If bCDGB = 3 Then
			PGrid.set_CellText(-1, pCOL, "いいえ")
			PGrid.set_ColWidth(pCOL, 70)
		Else
			PGrid.set_CellText(-1, pCOL, "０")
			PGrid.set_ColWidth(pCOL, 35)
		End If
		'------------------
		PGrid.set_ColName(pCOL, "KB4")
		PGrid.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_チェックボックス)
		PGrid.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		
		PGrid.set_ColExitMode(pCOL, 2) ' Exit動作 行相対/列絶対
		PGrid.set_ColExitCol(pCOL, 0) ' 第０列へ
		PGrid.set_ColExitRow(pCOL, 1) ' 次行へ
		
		PGrid.TextAtAddItem = ""
		PGrid.AddItems(0, PGrid.Rows)
		
		For pIX = 0 To (PGrid.Items - 1)
			PGrid.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
		Next pIX
		
		' --------------------------------------
		If PGrid.get_ColWidth(-1) <= 0 Then
			pWW = pWW + PGrid.DefWidth
		Else
			pWW = pWW + PGrid.get_ColWidth(-1)
		End If
		
		For pII = 0 To PGrid.Cols - 1
			pWW = pWW + PGrid.get_ColWidth(pII)
		Next pII
		' ｽｸﾛｰﾙﾊﾞｰの幅 = 20
		PGrid.Width = VB6.TwipsToPixelsX((pWW + 20) * VB6.TwipsPerPixelX)
		
		'UPGRADE_NOTE: オブジェクト pFONT をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		pFONT = Nothing
		mSW_CellFocusEvent = True
		
		PGrid.ExitCol = 1
		
	End Sub
	
	
	' *******************************************************************************
	' 概要 : Ｐｅｒｆｅｃｔ Ｇｒｉｄの設定。（概略評価）
	' 　　 :
	' 説明 :
	' 　　 :
	' 履歴 : 2005.12.12  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Sub Grid_Resize2()
		
		Dim pRetVal As Short
		Dim pFONT As System.Drawing.Font = System.Windows.Forms.Control.DefaultFont.Clone()
		Dim pII As Integer
		Dim pCOL As Integer
		Dim pCol2 As Integer
		Dim pRTN As Short
		Dim pIX As Short
		Dim pWW As Integer
		
		pFONT = VB6.FontChangeName(pFONT, "ＭＳ ゴシック")
		pFONT = VB6.FontChangeSize(pFONT, 8)
		
		' 全体プロパティの設定
		PGrid2.BorderWidth = 1 '境界線あり
		PGrid2.ExitCol = 1 '右列へ
		PGrid2.VertBorderWidth = 1 '列境界線幅１
		PGrid2.AdjustColWidth = False '列幅補正なし
		PGrid2.OddRowBorderWidth = 1 ' 遇数行と奇数行の境界
		PGrid2.OddRowBorder = 3 ' 点線
		PGrid2.OddRowBorderColor = &HFFC0C0 ' 淡い紫
		PGrid2.OddRowMeansLogical = False ' 奇数行は表示上
		PGrid2.SepAlwaysDraw = True
		'    PGrid2.ForeColor = vbBlue
		
		'    PGrid2.FoldCol = 6
		PGrid2.ThruStartShift = 1
		PGrid2.ThruStartVKeyName = "VK_INSERT"
		
		PGrid2.AllowSelCol = False
		PGrid2.AllowSelRow = False
		PGrid2.HeightSizing = 0
		PGrid2.WidthSizing = 0
		
		PGrid2.set_ColWidth(-1, 30) '連番
		PGrid2.set_CellText(-1, -1, "行")
		PGrid2.set_ColAlignmentH(-1, PGRIDLib.ColAlignmentHConstants.pgcah_右揃え)
		'セルプロパティの設定
		
		pCOL = 0
		PGrid2.set_CellText(-1, pCOL, "ｸﾞﾙｰﾌﾟ")
		PGrid2.set_CellAlignmentH(-1, pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid2.set_ColName(pCOL, "CDGP")
		PGrid2.set_ColWidth(pCOL, 50)
		PGrid2.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_右揃え)
		PGrid2.set_ColMaxLengthB(pCOL, 2)
		PGrid2.set_ColMinValue(pCOL, 0)
		PGrid2.set_ColMaxValue(pCOL, 99)
		PGrid2.set_ColFormatString(pCOL, "00")
		PGrid2.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		PGrid2.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_テキスト)
		
		pCOL = pCOL + 1
		PGrid2.set_CellText(-1, pCOL, "大")
		PGrid2.set_ColName(pCOL, "KBDAI")
		PGrid2.set_ColWidth(pCOL, 25)
		PGrid2.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid2.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_チェックボックス)
		PGrid2.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		PGrid2.set_CellText(-1, pCOL, "空")
		PGrid2.set_ColName(pCOL, "KBKR")
		PGrid2.set_ColWidth(pCOL, 25)
		PGrid2.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid2.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_チェックボックス)
		PGrid2.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		PGrid2.set_CellText(-1, pCOL, "解答欄")
		PGrid2.set_CellAlignmentH(-1, pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid2.set_ColName(pCOL, "KB1")
		PGrid2.set_ColWidth(pCOL, 60)
		PGrid2.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid2.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_チェックボックス)
		PGrid2.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		PGrid2.set_CellText(-1, pCOL, "評価")
		PGrid2.set_CellAlignmentH(-1, pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid2.set_ColName(pCOL, "HYOKA")
		PGrid2.set_ColWidth(pCOL, 60)
		PGrid2.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_右揃え)
		PGrid2.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_テキスト)
		PGrid2.set_ColMaxLengthB(pCOL, 1)
		PGrid2.set_ColMinValue(pCOL, 0)
		PGrid2.set_ColMaxValue(pCOL, 9)
		PGrid2.set_ColFormatString(pCOL, "#")
		PGrid2.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		PGrid2.set_CellText(-1, pCOL, "内  容")
		PGrid2.set_CellAlignmentH(-1, pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_左揃え)
		PGrid2.set_ColName(pCOL, "NAIYO")
		PGrid2.set_ColWidth(pCOL, 645)
		PGrid2.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_左揃え)
		PGrid2.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_テキスト)
		PGrid2.set_ColMaxLengthB(pCOL, 80)
		PGrid2.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_全角ひらがな)
		
		PGrid2.set_ColExitMode(pCOL, 2) ' Exit動作 行相対/列絶対
		PGrid2.set_ColExitCol(pCOL, 0) ' 第０列へ
		PGrid2.set_ColExitRow(pCOL, 1) ' 次行へ
		
		PGrid2.TextAtAddItem = ""
		PGrid2.AddItems(0, PGrid2.Rows)
		
		For pIX = 0 To (PGrid2.Items - 1)
			PGrid2.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
		Next pIX
		
		' --------------------------------------
		If PGrid2.get_ColWidth(-1) <= 0 Then
			pWW = pWW + PGrid2.DefWidth
		Else
			pWW = pWW + PGrid2.get_ColWidth(-1)
		End If
		
		For pII = 0 To PGrid2.Cols - 1
			pWW = pWW + PGrid2.get_ColWidth(pII)
		Next pII
		' ｽｸﾛｰﾙﾊﾞｰの幅 = 20
		PGrid2.Width = VB6.TwipsToPixelsX((pWW + 4) * VB6.TwipsPerPixelX)
		
		'UPGRADE_NOTE: オブジェクト pFONT をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		pFONT = Nothing
		mSW_CellFocusEvent = True
		
		PGrid2.ExitCol = 1
		
	End Sub
	
	
	
	' *******************************************************************************
	' 概要 : Ｐｅｒｆｅｃｔ Ｇｒｉｄの設定。（問題点）
	' 　　 :
	' 説明 :
	' 　　 :
	' 履歴 : 2005.12.02  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Sub Grid_Resize3()
		
		Dim pRetVal As Short
		Dim pFONT As System.Drawing.Font = System.Windows.Forms.Control.DefaultFont.Clone()
		Dim pII As Integer
		Dim pCOL As Integer
		Dim pCol2 As Integer
		Dim pRTN As Short
		Dim pIX As Short
		Dim pWW As Integer
		
		pFONT = VB6.FontChangeName(pFONT, "ＭＳ ゴシック")
		pFONT = VB6.FontChangeSize(pFONT, 8)
		
		' 全体プロパティの設定
		PGrid3.BorderWidth = 1 '境界線あり
		PGrid3.ExitCol = 1 '右列へ
		PGrid3.VertBorderWidth = 1 '列境界線幅１
		PGrid3.AdjustColWidth = False '列幅補正なし
		PGrid3.OddRowBorderWidth = 1 ' 遇数行と奇数行の境界
		PGrid3.OddRowBorder = 3 ' 点線
		PGrid3.OddRowBorderColor = &HFFC0C0 ' 淡い紫
		PGrid3.OddRowMeansLogical = False ' 奇数行は表示上
		PGrid3.SepAlwaysDraw = True
		'    PGrid3.ForeColor = vbBlue
		
		'    PGrid3.FoldCol = 6
		PGrid3.ThruStartShift = 1
		PGrid3.ThruStartVKeyName = "VK_INSERT"
		
		PGrid3.AllowSelCol = False
		PGrid3.AllowSelRow = False
		PGrid3.HeightSizing = 0
		PGrid3.WidthSizing = 0
		
		PGrid3.set_ColWidth(-1, 30) '連番
		PGrid3.set_CellText(-1, -1, "行")
		PGrid3.set_ColAlignmentH(-1, PGRIDLib.ColAlignmentHConstants.pgcah_右揃え)
		'セルプロパティの設定
		
		pCOL = 0
		PGrid3.set_CellText(-1, pCOL, "ｸﾞﾙｰﾌﾟ")
		PGrid3.set_CellAlignmentH(-1, pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid3.set_ColName(pCOL, "CDGP")
		PGrid3.set_ColWidth(pCOL, 50)
		PGrid3.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_右揃え)
		PGrid3.set_ColMaxLengthB(pCOL, 2)
		PGrid3.set_ColMinValue(pCOL, 0)
		PGrid3.set_ColMaxValue(pCOL, 99)
		PGrid3.set_ColFormatString(pCOL, "00")
		PGrid3.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		PGrid3.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_テキスト)
		
		pCOL = pCOL + 1
		PGrid3.set_CellText(-1, pCOL, "大")
		PGrid3.set_ColName(pCOL, "KBDAI")
		PGrid3.set_ColWidth(pCOL, 25)
		PGrid3.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid3.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_チェックボックス)
		PGrid3.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		PGrid3.set_CellText(-1, pCOL, "空")
		PGrid3.set_ColName(pCOL, "KBKR")
		PGrid3.set_ColWidth(pCOL, 25)
		PGrid3.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid3.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_チェックボックス)
		PGrid3.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		PGrid3.set_CellText(-1, pCOL, "解答欄")
		PGrid3.set_CellAlignmentH(-1, pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid3.set_ColName(pCOL, "KB1")
		PGrid3.set_ColWidth(pCOL, 60)
		PGrid3.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_中央揃え)
		PGrid3.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_チェックボックス)
		PGrid3.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_オフ固定)
		
		pCOL = pCOL + 1
		PGrid3.set_CellText(-1, pCOL, "内  容")
		PGrid3.set_CellAlignmentH(-1, pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_左揃え)
		PGrid3.set_ColName(pCOL, "MONDAI")
		PGrid3.set_ColWidth(pCOL, 645)
		PGrid3.set_ColAlignmentH(pCOL, PGRIDLib.ColAlignmentHConstants.pgcah_左揃え)
		PGrid3.set_ColStyle(pCOL, PGRIDLib.ColStyleConstants.pgcs_テキスト)
		PGrid3.set_ColMaxLengthB(pCOL, 80)
		PGrid3.set_ColIMEMode(pCOL, PGRIDLib.ColIMEModeConstants.pgcim_全角ひらがな)
		
		PGrid3.set_ColExitMode(pCOL, 2) ' Exit動作 行相対/列絶対
		PGrid3.set_ColExitCol(pCOL, 0) ' 第０列へ
		PGrid3.set_ColExitRow(pCOL, 1) ' 次行へ
		
		PGrid3.TextAtAddItem = ""
		PGrid3.AddItems(0, PGrid3.Rows)
		
		For pIX = 0 To (PGrid3.Items - 1)
			PGrid3.set_CellText(pIX, -1, CStr(pIX + 1) & " ")
		Next pIX
		
		' --------------------------------------
		If PGrid3.get_ColWidth(-1) <= 0 Then
			pWW = pWW + PGrid3.DefWidth
		Else
			pWW = pWW + PGrid3.get_ColWidth(-1)
		End If
		
		For pII = 0 To PGrid3.Cols - 1
			pWW = pWW + PGrid3.get_ColWidth(pII)
		Next pII
		' ｽｸﾛｰﾙﾊﾞｰの幅 = 20
		PGrid3.Width = VB6.TwipsToPixelsX((pWW + 5) * VB6.TwipsPerPixelX)
		
		'UPGRADE_NOTE: オブジェクト pFONT をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		pFONT = Nothing
		mSW_CellFocusEvent = True
		
		PGrid3.ExitCol = 1
		
	End Sub
	
	' *******************************************************************************
	' 概要  : ＯＳＣＥコントロールマスター２のＲＥＡＤを行う。(学年のチェック)
	' 　　  :
	' 引数  : 戻り値, O, Integer, True=READ成功。
	' 　　  : 　　　              False=READ失敗。
	' 　　  :
	' 説明  :
	' 　　  :
	' 履歴  : 2005.12.01  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function OSC_CMM2_Read(ByVal pCDGK As Integer, ByVal pYEAR As Integer, ByRef pKBMEK As Integer) As Short
		
		Dim pRetVal As Short
		Dim pROW As Integer
		Dim pNAME As String
		Dim pIX As Integer
		Dim pIY As Integer
		
		OSC_CMM2_Read = False
		
		'''------------------------------------------
		dbCmdCMM2.Parameters(0).Value = 1
		dbCmdCMM2.Parameters(1).Value = bCDGB
		dbCmdCMM2.Parameters(2).Value = pCDGK
		dbCmdCMM2.Parameters(3).Value = pYEAR
		
		On Error Resume Next
		dbRecCMM2.Requery()
		
		If Err.Number <> 0 Then
			mMsgText = "ＯＳＣＥコントロールマスターのＲＥＡＤでエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
			Err.Clear()
			On Error GoTo 0
			Exit Function
		End If
		On Error GoTo 0
		
		'''------------------------------------------
		If dbRecCMM2.EOF = False Then
			pKBMEK = dbRecCMM2.Fields("CM2KBMEK").Value
			'データあり→ＯＫ
			OSC_CMM2_Read = True
		End If
		
	End Function
	
	
	
	
	Private Sub PGrid_CellClick(ByVal eventSender As System.Object, ByVal eventArgs As AxPGRIDLib._DPGridEvents_CellClickEvent) Handles PGrid.CellClick
		
		Dim pRetVal As Short
		Dim pSW_ERROR As Boolean
		Dim pSW_Select As Boolean
		Dim pINCODE As String
		Dim pNMKN As String
		Dim pNMCH As String
		Dim pNMIK As String
		
		'
		' PrefectGrid のサイズ設定中はこのイベントを無視する
		'
		If Not mSW_CellFocusEvent Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		'
		' 行を移動した場合
		'
		If PGrid.NextRow > (PGrid.Items - 1) Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If PGrid.NextRow < 0 Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If eventArgs.Row > (PGrid.Items - 1) Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If eventArgs.Row < 0 Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If (eventArgs.Row <> PGrid.NextRow) Then
			'
			' 移動先の行の前の行が空白行かチェックする
			'
			''If Grid1SpaceGyoCheck(PGrid, PGrid.NextRow - 1) Then
			If Grid1SpaceGyoCheck(PGrid, eventArgs.Row) Then
				PGrid.NextRow = eventArgs.Row
				PGrid.NextCol = eventArgs.Col
				GoTo PGrid_CellLostFocus_Exit
			End If
		End If
		'
		' 各セルをチェックする
		'
		pSW_ERROR = False
		
		Select Case eventArgs.Col
			Case PGrid.get_ColOfColName("KBDAI")
				'大・中のどちらか
				PGrid.set_CellCheckedByName(eventArgs.Row, "KBCHU", False)
				Call ENABLE_Change(eventArgs.Row)
				
			Case PGrid.get_ColOfColName("KBCHU")
				'大・中のどちらか
				PGrid.set_CellCheckedByName(eventArgs.Row, "KBDAI", False)
				Call ENABLE_Change(eventArgs.Row)
			Case PGrid.get_ColOfColName("KBKR")
				Call ENABLE_Change(eventArgs.Row)
			Case PGrid.get_ColOfColName("KB1"), PGrid.get_ColOfColName("KB2"), PGrid.get_ColOfColName("KB3"), PGrid.get_ColOfColName("KB4") '2006.01.16 add
				Call ENABLE_Change(eventArgs.Row)
		End Select
		
		' 現在入力中の行の背景色を変更する。 ---------------------------------
		PGrid.set_RowBackColor(eventArgs.Row, System.Convert.ToUInt32(System.Drawing.Color.White))
		PGrid.set_RowBackColor(PGrid.NextRow, System.Convert.ToUInt32(System.Drawing.ColorTranslator.FromOle(&HC0E0FF)))
		
		
		
		'
		' フォーカスをもとに戻す
		'
		mCellClickCol = -1
		
		If pSW_ERROR Then
			PGrid.NextRow = eventArgs.Row
			PGrid.NextCol = eventArgs.Col
		End If
		
		GoTo PGrid_CellLostFocus_Exit
		
		'---------------------------------------------
PGrid_CellLostFocus_Exit: 
		
		mSW_CellKeyPress = False
		
	End Sub
	
	'概略評価----------------------------------------------------------------
	Private Sub PGrid2_CellClick(ByVal eventSender As System.Object, ByVal eventArgs As AxPGRIDLib._DPGridEvents_CellClickEvent) Handles PGrid2.CellClick
		
		Dim pRetVal As Short
		Dim pSW_ERROR As Boolean
		Dim pSW_Select As Boolean
		Dim pINCODE As String
		Dim pNMKN As String
		Dim pNMCH As String
		Dim pNMIK As String
		
		'
		' PrefectGrid のサイズ設定中はこのイベントを無視する
		'
		If Not mSW_CellFocusEvent Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		'
		' 行を移動した場合
		'
		If PGrid2.NextRow > (PGrid2.Items - 1) Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If PGrid2.NextRow < 0 Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If eventArgs.Row > (PGrid2.Items - 1) Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If eventArgs.Row < 0 Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If (eventArgs.Row <> PGrid2.NextRow) Then
			'
			' 移動先の行の前の行が空白行かチェックする
			'
			'If Grid1SpaceGyoCheck(PGrid2, PGrid2.NextRow - 1) Then
			If Grid1SpaceGyoCheck(PGrid2, eventArgs.Row) Then
				PGrid2.NextRow = eventArgs.Row
				PGrid2.NextCol = eventArgs.Col
				GoTo PGrid_CellLostFocus_Exit
			End If
		End If
		'
		' 各セルをチェックする
		'
		pSW_ERROR = False
		
		Select Case eventArgs.Col
			Case PGrid2.get_ColOfColName("KBDAI")
				''        '大・中のどちらか
				''PGrid2.CellCheckedByName(Row, "KBCHU") = False
				Call ENABLE_Change2(eventArgs.Row)
				''    Case PGrid2.ColOfColName("KBCHU")
				''        '大・中のどちらか
				''        PGrid2.CellCheckedByName(Row, "KBDAI") = False
				''        Call Button_Change
			Case PGrid2.get_ColOfColName("KBKR")
				Call ENABLE_Change2(eventArgs.Row)
		End Select
		
		' 現在入力中の行の背景色を変更する。 ---------------------------------
		PGrid2.set_RowBackColor(eventArgs.Row, System.Convert.ToUInt32(System.Drawing.Color.White))
		PGrid2.set_RowBackColor(PGrid2.NextRow, System.Convert.ToUInt32(System.Drawing.ColorTranslator.FromOle(&HC0E0FF)))
		
		
		
		'
		' フォーカスをもとに戻す
		'
		mCellClickCol = -1
		
		If pSW_ERROR Then
			PGrid2.NextRow = eventArgs.Row
			PGrid2.NextCol = eventArgs.Col
		End If
		
		GoTo PGrid_CellLostFocus_Exit
		
		'---------------------------------------------
PGrid_CellLostFocus_Exit: 
		
		mSW_CellKeyPress = False
		
	End Sub
	
	
	Private Sub PGrid2_CellLostFocus(ByVal eventSender As System.Object, ByVal eventArgs As AxPGRIDLib._DPGridEvents_CellLostFocusEvent) Handles PGrid2.CellLostFocus
		
		Dim pRetVal As Short
		Dim pSW_ERROR As Boolean
		Dim pCOL As Integer
		Dim pCellText As String
		Dim pJU As String
		Dim pNMKN As String
		Dim pNMCH As String
		Dim pNMIK As String
		Dim pTIME As String
		Dim pLEN As Short
		
		'
		' PrefectGrid のサイズ設定中はこのイベントを無視する
		'
		If Not mSW_CellFocusEvent Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		'
		' 行を移動した場合
		'
		If PGrid2.NextRow > (PGrid2.Items - 1) Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If (eventArgs.Row <> PGrid2.NextRow) Then
			'
			' 移動先の行の前の行が空白行かチェックする
			'
			If Grid1SpaceGyoCheck(PGrid2, PGrid2.NextRow - 1) Then
				PGrid2.NextRow = eventArgs.Row
				PGrid2.NextCol = eventArgs.Col
				GoTo PGrid_CellLostFocus_Exit
			End If
		End If
		'
		' 各セルをチェックする
		'
		pSW_ERROR = False
		
		' 現在入力中の行の背景色を変更する。 ---------------------------------
		PGrid2.set_RowBackColor(eventArgs.Row, System.Convert.ToUInt32(System.Drawing.Color.White))
		'PGrid2.RowBackColor(PGrid2.NextRow) = &HC0E0FF
		
		Select Case eventArgs.Col
			'' ｺｰﾄﾞ -----------------------------------------------------------
			Case PGrid2.get_ColOfColName("CDGP")
				If PGrid2.get_CellTextByName(eventArgs.Row, "CDGP") <> "" And PGrid2.get_CellCheckedByName(eventArgs.Row, "KBKR") = False Then
					If Not CellCheck_Numeric(PGrid2, eventArgs.Row, eventArgs.Col) Then
						pSW_ERROR = True
					End If
					'前の行との大小をcheck
					If Not CDGP_Check(PGrid2, eventArgs.Row, PGrid2.get_CellValue(eventArgs.Row, eventArgs.Col)) Then
						pSW_ERROR = True
					End If
				End If
				'' 評価 -----------------------------------------------------------
			Case PGrid2.get_ColOfColName("HYOKA")
				If PGrid2.get_CellTextByName(eventArgs.Row, "HYOKA") <> "" And PGrid2.get_CellCheckedByName(eventArgs.Row, "KBKR") = False Then
					If Not CellCheck_Numeric(PGrid2, eventArgs.Row, eventArgs.Col) Then
						pSW_ERROR = True
					End If
					
				End If
				'' 内容-----------------------------------------------------------
			Case PGrid2.get_ColOfColName("NAIYO")
				If mKBHY = 0 Then
					If PGrid2.get_CellCheckedByName(eventArgs.Row, "KBDAI") = True Then
						PGrid2.set_CellText(eventArgs.Row, eventArgs.Col, MOJIByte_Get(PGrid2.get_CellText(eventArgs.Row, eventArgs.Col), Me, 1, 50, pLEN))
					Else
						PGrid2.set_CellText(eventArgs.Row, eventArgs.Col, MOJIByte_Get(PGrid2.get_CellText(eventArgs.Row, eventArgs.Col), Me, 1, 44, pLEN))
					End If
				ElseIf mKBHY = 1 Then 
					PGrid2.set_CellText(eventArgs.Row, eventArgs.Col, MOJIByte_Get(PGrid2.get_CellText(eventArgs.Row, eventArgs.Col), Me, 1, 80, pLEN))
				End If
				
				
		End Select
		'
		' フォーカスをもとに戻す
		'
		mCellClickCol = -1
		
		If pSW_ERROR Then
			PGrid2.NextRow = eventArgs.Row
			PGrid2.NextCol = eventArgs.Col
		End If
		
		GoTo PGrid_CellLostFocus_Exit
		
		'---------------------------------------------
PGrid_CellLostFocus_Exit: 
		
		mSW_CellKeyPress = False
		
	End Sub
	
	
	'大・空の区分により、入力可否を切り替える（概略評価グリッド用）
	Private Function ENABLE_Change2(ByVal pROW As Integer) As Integer
		
		Dim pCOL As Integer
		Dim pIX As Integer
		
		'空にチェックあり？？----------------------------------
		If PGrid2.get_CellCheckedByName(pROW, "KBKR") = True Then
			'空行区分以外の項目入力不可
			For pIX = 0 To PGrid2.Cols - 1
				If pIX <> PGrid2.get_ColOfColName("KBKR") Then
					PGrid2.set_CellEnabled(pROW, pIX, False)
				Else
					PGrid2.set_CellEnabled(pROW, pIX, True)
				End If
			Next pIX
			Exit Function
		Else
			For pIX = 0 To PGrid2.Cols - 1
				PGrid2.set_CellEnabled(pROW, pIX, True)
			Next pIX
		End If
		
		'大にチェックあり？？----------------------------------
		If PGrid2.get_CellCheckedByName(pROW, "KBDAI") = True Then
			'空行区分・解答欄・評価入力不可
			For pIX = 0 To PGrid2.Cols - 1
				If pIX = PGrid2.get_ColOfColName("KBKR") Or pIX = PGrid2.get_ColOfColName("KB1") Or pIX = PGrid2.get_ColOfColName("HYOKA") Then
					PGrid2.set_CellEnabled(pROW, pIX, False)
				End If
			Next pIX
			Exit Function
		Else
			'上の反対
			For pIX = 0 To PGrid2.Cols - 1
				If pIX = PGrid2.get_ColOfColName("KBKR") Or pIX = PGrid2.get_ColOfColName("KB1") Or pIX = PGrid2.get_ColOfColName("HYOKA") Then
					PGrid2.set_CellEnabled(pROW, pIX, True)
				End If
			Next pIX
		End If
		
	End Function
	
	'大・空の区分により、入力可否を切り替える（問題点グリッド用）
	Private Function ENABLE_Change3(ByVal pROW As Integer) As Integer
		
		Dim pCOL As Integer
		Dim pIX As Integer
		
		'空にチェックあり？？----------------------------------
		If PGrid3.get_CellCheckedByName(pROW, "KBKR") = True Then
			'空行区分以外の項目入力不可
			For pIX = 0 To PGrid3.Cols - 1
				If pIX <> PGrid3.get_ColOfColName("KBKR") Then
					PGrid3.set_CellEnabled(pROW, pIX, False)
				Else
					PGrid3.set_CellEnabled(pROW, pIX, True)
				End If
			Next pIX
			Exit Function
		Else
			For pIX = 0 To PGrid3.Cols - 1
				PGrid3.set_CellEnabled(pROW, pIX, True)
			Next pIX
		End If
		
		'大にチェックあり？？----------------------------------
		If PGrid3.get_CellCheckedByName(pROW, "KBDAI") = True Then
			'空行区分・解答欄入力不可
			For pIX = 0 To PGrid3.Cols - 1
				If pIX = PGrid3.get_ColOfColName("KBKR") Or pIX = PGrid3.get_ColOfColName("KB1") Then
					PGrid3.set_CellEnabled(pROW, pIX, False)
				End If
			Next pIX
			Exit Function
		Else
			'上の反対
			For pIX = 0 To PGrid3.Cols - 1
				If pIX = PGrid3.get_ColOfColName("KBKR") Or pIX = PGrid3.get_ColOfColName("KB1") Then
					PGrid3.set_CellEnabled(pROW, pIX, True)
				End If
			Next pIX
		End If
		
	End Function
	
	'大・中・空の区分により、入力可否を切り替える
	Private Function ENABLE_Change(ByVal pROW As Integer) As Integer
		
		Dim pCOL As Integer
		Dim pIX As Integer
		
		'空にチェックあり？？----------------------------------
		If PGrid.get_CellCheckedByName(pROW, "KBKR") = True Then
			'空行区分以外の項目入力不可
			For pIX = 0 To PGrid.Cols - 1
				If pIX <> PGrid.get_ColOfColName("KBKR") Then
					PGrid.set_CellEnabled(pROW, pIX, False)
				Else
					PGrid.set_CellEnabled(pROW, pIX, True)
				End If
			Next pIX
			'exit Function
		Else
			For pIX = 0 To PGrid.Cols - 1
				PGrid.set_CellEnabled(pROW, pIX, True)
			Next pIX
			
			'大にチェックあり？？----------------------------------
			If PGrid.get_CellCheckedByName(pROW, "KBDAI") = True Then
				'解答数入力可、空行区分・解答有無入力不可
				For pIX = 0 To PGrid.Cols - 1
					If pIX = PGrid.get_ColOfColName("SUKA") Then
						PGrid.set_CellEnabled(pROW, pIX, True)
					End If
					If pIX = PGrid.get_ColOfColName("KBKR") Or pIX = PGrid.get_ColOfColName("KB1") Or pIX = PGrid.get_ColOfColName("KB2") Or pIX = PGrid.get_ColOfColName("KB3") Or pIX = PGrid.get_ColOfColName("KB4") Or pIX = PGrid.get_ColOfColName("NO") Then
						PGrid.set_CellEnabled(pROW, pIX, False)
					End If
				Next pIX
				'Exit Function
			Else
				'上の反対
				For pIX = 0 To PGrid.Cols - 1
					If pIX = PGrid.get_ColOfColName("SUKA") Then
						PGrid.set_CellEnabled(pROW, pIX, False)
					End If
					If pIX = PGrid.get_ColOfColName("KBKR") Or pIX = PGrid.get_ColOfColName("KB1") Or pIX = PGrid.get_ColOfColName("KB2") Or pIX = PGrid.get_ColOfColName("KB3") Or pIX = PGrid.get_ColOfColName("KB4") Then
						PGrid.set_CellEnabled(pROW, pIX, True)
					End If
				Next pIX
			End If
			
			'中にチェックあり？？----------------------------------
			If PGrid.get_CellCheckedByName(pROW, "KBCHU") = True Then
				'解答数・空行区分・解答有無入力不可
				For pIX = 0 To PGrid.Cols - 1
					If pIX = PGrid.get_ColOfColName("SUKA") Or pIX = PGrid.get_ColOfColName("KBKR") Or pIX = PGrid.get_ColOfColName("KB1") Or pIX = PGrid.get_ColOfColName("KB2") Or pIX = PGrid.get_ColOfColName("KB3") Or pIX = PGrid.get_ColOfColName("KB4") Or pIX = PGrid.get_ColOfColName("NO") Then
						PGrid.set_CellEnabled(pROW, pIX, False)
					End If
				Next pIX
				'Exit Function
			End If
			
			'2006.01.16 add
			'解答ありにチェックあり？？----------------------------
			If (PGrid.get_CellEnabled(pROW, PGrid.get_ColOfColName("KB1")) = True Or PGrid.get_CellEnabled(pROW, PGrid.get_ColOfColName("KB2")) = True Or PGrid.get_CellEnabled(pROW, PGrid.get_ColOfColName("KB3")) = True Or PGrid.get_CellEnabled(pROW, PGrid.get_ColOfColName("KB4")) = True) And (PGrid.get_CellCheckedByName(pROW, "KB1") = True Or PGrid.get_CellCheckedByName(pROW, "KB2") = True Or PGrid.get_CellCheckedByName(pROW, "KB3") = True Or PGrid.get_CellCheckedByName(pROW, "KB4") = True) Then
				'解答番号入力可
				PGrid.set_CellEnabled(pROW, PGrid.get_ColOfColName("NO"), True)
			Else
				PGrid.set_CellEnabled(pROW, PGrid.get_ColOfColName("NO"), False)
			End If
		End If
		
		'解答有無欄の入力可否を切り替える
		Call ENABLE_Change_SUKA(pROW)
		
		
	End Function
	
	'解答数により、入力可否を切り替える
	Private Function ENABLE_Change_SUKA(ByVal pROW As Integer) As Integer
		
		Dim pCOL As Integer
		Dim pIX As Integer
		Dim pSUKA As Integer
		Dim pIY As Integer
		
		'現在の行の対象となる解答数を探す(大問)
		
		''    If pROW = 0 Then
		''        Exit Function
		''    End If
		
		For pIX = pROW To 0 Step -1
			If PGrid.get_CellCheckedByName(pIX, "KBDAI") = True Then
				pSUKA = PGrid.get_CellValueByName(pIX, "SUKA")
				Exit For
			End If
		Next pIX
		
		'----------------------------------------
		'次の大問の前の行まで
		For pIY = pIX + 1 To PGrid.Items - 1
			''If Grid1SpaceGyoCheck(PGrid, pIX) = False Then
			
			If PGrid.get_CellCheckedByName(pIY, "KBCHU") = False And PGrid.get_CellCheckedByName(pIY, "KBKR") = False Then
				
				Select Case pSUKA
					Case 0
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB1"), False)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB2"), False)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB3"), False)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB4"), False)
					Case 1
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB1"), False)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB2"), False)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB3"), False)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB4"), True)
					Case 2
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB1"), False)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB2"), False)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB3"), True)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB4"), True)
					Case 3
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB1"), False)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB2"), True)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB3"), True)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB4"), True)
					Case 4
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB1"), True)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB2"), True)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB3"), True)
						PGrid.set_CellEnabled(pIY, PGrid.get_ColOfColName("KB4"), True)
				End Select
			End If
			
			'次の行が大問だったらＥＸＩＴ
			If pIY + 1 <= PGrid.Items - 1 Then
				If PGrid.get_CellCheckedByName(pIY + 1, "KBDAI") = True Then
					Exit For
				End If
			End If
			''End If
		Next pIY
		
		
	End Function
	
	
	Private Sub PGrid_CellGotFocus(ByVal eventSender As System.Object, ByVal eventArgs As AxPGRIDLib._DPGridEvents_CellGotFocusEvent) Handles PGrid.CellGotFocus
		
		Dim pIX As Short
		
		''mActive_PGrid = 0
		' 現在入力中の行の背景色を変更する。 ---------------------------------
		PGrid.set_RowBackColor(eventArgs.Row, System.Convert.ToUInt32(System.Drawing.ColorTranslator.FromOle(&HC0E0FF)))
		
		''    Call Grid1Clear
		''    Call Disp_KBSK(Row)
		
	End Sub
	
	Private Sub PGrid2_CellGotFocus(ByVal eventSender As System.Object, ByVal eventArgs As AxPGRIDLib._DPGridEvents_CellGotFocusEvent) Handles PGrid2.CellGotFocus
		
		Dim pIX As Short
		
		''mActive_PGrid = 0
		' 現在入力中の行の背景色を変更する。 ---------------------------------
		PGrid2.set_RowBackColor(eventArgs.Row, System.Convert.ToUInt32(System.Drawing.ColorTranslator.FromOle(&HC0E0FF)))
		
		''    Call Grid1Clear
		''    Call Disp_KBSK(Row)
		
	End Sub
	
	Private Sub PGrid3_CellGotFocus(ByVal eventSender As System.Object, ByVal eventArgs As AxPGRIDLib._DPGridEvents_CellGotFocusEvent) Handles PGrid3.CellGotFocus
		
		Dim pIX As Short
		
		''mActive_PGrid = 0
		' 現在入力中の行の背景色を変更する。 ---------------------------------
		PGrid3.set_RowBackColor(eventArgs.Row, System.Convert.ToUInt32(System.Drawing.ColorTranslator.FromOle(&HC0E0FF)))
		
		''    Call Grid1Clear
		''    Call Disp_KBSK(Row)
		
	End Sub
	
	Private Sub PGrid_CellLostFocus(ByVal eventSender As System.Object, ByVal eventArgs As AxPGRIDLib._DPGridEvents_CellLostFocusEvent) Handles PGrid.CellLostFocus
		
		Dim pRetVal As Short
		Dim pSW_ERROR As Boolean
		Dim pCOL As Integer
		Dim pCellText As String
		Dim pJU As String
		Dim pNMKN As String
		Dim pNMCH As String
		Dim pNMIK As String
		Dim pTIME As String
		
		'
		' PrefectGrid のサイズ設定中はこのイベントを無視する
		'
		If Not mSW_CellFocusEvent Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		'
		' 行を移動した場合
		'
		If PGrid.NextRow > (PGrid.Items - 1) Then
			GoTo PGrid_CellLostFocus_Exit
		End If
		
		If (eventArgs.Row <> PGrid.NextRow) Then
			'
			' 移動先の行の前の行が空白行かチェックする
			'
			If Grid1SpaceGyoCheck(PGrid, PGrid.NextRow - 1) Then
				PGrid.NextRow = eventArgs.Row
				PGrid.NextCol = eventArgs.Col
				GoTo PGrid_CellLostFocus_Exit
			End If
		End If
		'
		' 各セルをチェックする
		'
		pSW_ERROR = False
		
		' 現在入力中の行の背景色を変更する。 ---------------------------------
		PGrid.set_RowBackColor(eventArgs.Row, System.Convert.ToUInt32(System.Drawing.Color.White))
		'PGrid.RowBackColor(PGrid.NextRow) = &HC0E0FF
		
		Select Case eventArgs.Col
			'' ｺｰﾄﾞ -----------------------------------------------------------
			Case PGrid.get_ColOfColName("CDGP")
				If PGrid.get_CellTextByName(eventArgs.Row, "CDGP") <> "" And PGrid.get_CellCheckedByName(eventArgs.Row, "KBKR") = False Then
					If Not CellCheck_Numeric(PGrid, eventArgs.Row, eventArgs.Col) Then
						pSW_ERROR = True
					End If
					'前の行との大小をcheck
					If Not CDGP_Check(PGrid, eventArgs.Row, PGrid.get_CellValue(eventArgs.Row, eventArgs.Col)) Then
						pSW_ERROR = True
					End If
				End If
				'' 解答数 -----------------------------------------------------------
			Case PGrid.get_ColOfColName("SUKA")
				If PGrid.get_CellTextByName(eventArgs.Row, "SUKA") <> "" And PGrid.get_CellCheckedByName(eventArgs.Row, "KBKR") = False Then
					If Not CellCheck_Numeric(PGrid, eventArgs.Row, eventArgs.Col) Then
						pSW_ERROR = True
					End If
					'解答有無欄の入力可否を切り替える
					Call ENABLE_Change_SUKA(eventArgs.Row)
				End If
				
				
				
		End Select
		'
		' フォーカスをもとに戻す
		'
		mCellClickCol = -1
		
		If pSW_ERROR Then
			PGrid.NextRow = eventArgs.Row
			PGrid.NextCol = eventArgs.Col
		End If
		
		GoTo PGrid_CellLostFocus_Exit
		
		'---------------------------------------------
PGrid_CellLostFocus_Exit: 
		
		mSW_CellKeyPress = False
		
	End Sub
	
	' *******************************************************************************
	' 概要    : 会場マスターのＲＥＡＤを行う。会場数の取得
	'         :
	' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=READ成功。
	'         :                    False=READ失敗。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2005.12.02  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function OSC_KAIJOM_Read() As Short
		
		Dim pRetVal As Short
		Dim dbRec As ADODB.Recordset
		Dim pMAXLen As Short
		Dim pIX As Short
		
		OSC_KAIJOM_Read = False
		mSUKJ = 0
		mSUKJM = 0
		
		dbRec = New ADODB.Recordset
		
		'
		' 会場マスターＲＥＡＤ
		'
		'医療面接会場区分毎の会場数
		SqlText = "select KJKBME, COUNT(KJCDKJ) as SUKJ  "
		SqlText = SqlText & " from OSC_KAIJOM "
		SqlText = SqlText & " where KJCDGA  = 1 "
		SqlText = SqlText & "   and KJCDGB  = " & CStr(bCDGB)
		SqlText = SqlText & "   and KJCDGK  = " & CStr(mCDGK)
		SqlText = SqlText & "   and KJNENDO = " & CStr(mNENDO)
		SqlText = SqlText & "   and KJKBSK = " & CStr(mKBSK)
		SqlText = SqlText & "   and KJYEAR = " & CStr(mYEAR)
		SqlText = SqlText & "  group by KJCDGA, KJCDGB, KJCDGK, KJNENDO, KJKBSK, KJYEAR, KJKBME "
		SqlText = SqlText & " order by KJCDGA, KJCDGB, KJCDGK, KJNENDO, KJKBSK, KJYEAR, KJKBME "
		
		On Error Resume Next
		dbRec.Open(SqlText, dbCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Err.Number <> 0 Then
			mMsgText = "会場マスターのＲＥＡＤでエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
			Err.Clear()
			On Error GoTo 0
			GoTo RecordSet_Close
		End If
		On Error GoTo 0
		
		If dbRec.EOF = True Then
			GoTo RecordSet_Close
		End If
		
		
		Do While dbRec.EOF = False
			
			If dbRec.Fields("KJKBME").Value = 0 Then
				'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
				mSUKJ = IIf(IsDbNull(dbRec.Fields("SUKJ").Value), 0, dbRec.Fields("SUKJ").Value)
			ElseIf dbRec.Fields("KJKBME").Value = 1 Then 
				'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
				mSUKJM = IIf(IsDbNull(dbRec.Fields("SUKJ").Value), 0, dbRec.Fields("SUKJ").Value)
			End If
			
			
			On Error Resume Next
			dbRec.MoveNext()
			
			If Err.Number <> 0 Then
				mMsgText = "会場マスターのＲＥＡＤでエラーが発生しました。"
				pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
				Err.Clear()
				On Error GoTo 0
				GoTo RecordSet_Close
			End If
			On Error GoTo 0
			
		Loop 
		
		OSC_KAIJOM_Read = True
		
RecordSet_Close: 
		On Error Resume Next
		dbRec.Close()
		'UPGRADE_NOTE: オブジェクト dbRec をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbRec = Nothing
		On Error GoTo 0
		
	End Function
	
	''' *******************************************************************************
	''' 概要    : ｽﾃｰｼｮﾝマスターのＲＥＡＤを行う。ｽﾃｰｼｮﾝ数の取得
	'''         :
	''' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=READ成功。
	'''         :                    False=READ失敗。
	'''         :
	''' 説明    :
	'''         :
	''' 履歴    : 2005.12.02  REV.0001  吉田  新規作成。
	''' *******************************************************************************
	'''
	''Private Function OSC_STM_READ() As Integer
	''
	''    Dim pRetVal             As Integer
	''    Dim dbRec               As ADODB.Recordset
	''    Dim pMAXLen             As Integer
	''    Dim pIX                 As Integer
	''
	''    OSC_STM_READ = False
	''    mSUST = 0
	''    mSUSTM = 0
	''    mSUBN = 0
	''
	''    Set dbRec = New ADODB.Recordset
	''
	''    '
	''    ' ｽﾃｰｼｮﾝマスターＲＥＡＤ
	''    '
	''    '医療面接区分毎のｽﾃｰｼｮﾝ数
	''    SqlText = "select STKBME, STSUBN  "
	''    SqlText = SqlText + " from OSC_STM "
	''    SqlText = SqlText + " where STCDGA  = 1 "
	''    SqlText = SqlText + "   and STCDGB  = " + CStr(bCDGB)
	''    SqlText = SqlText + "   and STCDGK  = " + CStr(mCDGK)
	''    SqlText = SqlText + "   and STNENDO = " + CStr(mNENDO)
	''    SqlText = SqlText + "   and STKBSK = " + CStr(mKBSK)
	''    SqlText = SqlText + "   and STYEAR = " + CStr(mYEAR)
	''    SqlText = SqlText + " order by STCDGA, STCDGB, STCDGK, STNENDO, STKBSK, STYEAR, STKBME "
	''
	''On Error Resume Next
	''    dbRec.Open SqlText, dbCon, adOpenForwardOnly, adLockReadOnly
	''
	''    If Err.Number <> 0 Then
	''        mMsgText = "ｽﾃｰｼｮﾝマスターのＲＥＡＤでエラーが発生しました。"
	''        pRetVal = ADOErrDisp(mMsgText, vbOKOnly + vbCritical, mMBOXTitle, Err.Description)
	''        Err.Clear
	''On Error GoTo 0
	''        GoTo Recordset_close
	''    End If
	''On Error GoTo 0
	''
	''    If dbRec.EOF = True Then
	''        GoTo Recordset_close
	''    End If
	''
	''
	''    Do While dbRec.EOF = False
	''
	''        If dbRec("STKBME") = 0 Then
	''            mSUST = mSUST + 1
	''        ElseIf dbRec("STKBME") = 1 Then     '医療面接
	''            mSUSTM = mSUSTM + 1
	''            mSUBN = dbRec("STSUBN")     '分割数
	''        End If
	''
	''
	''    On Error Resume Next
	''        dbRec.MoveNext
	''
	''        If Err.Number <> 0 Then
	''            mMsgText = "ｽﾃｰｼｮﾝマスターのＲＥＡＤでエラーが発生しました。"
	''            pRetVal = ADOErrDisp(mMsgText, vbOKOnly + vbCritical, mMBOXTitle, Err.Description)
	''            Err.Clear
	''    On Error GoTo 0
	''            GoTo Recordset_close
	''        End If
	''    On Error GoTo 0
	''    Loop
	''
	''    OSC_STM_READ = True
	''
	''Recordset_close:
	''On Error Resume Next
	''    dbRec.Close
	''    Set dbRec = Nothing
	''On Error GoTo 0
	''
	''End Function
	''
	''
	''
	''
	''
	' *******************************************************************************
	' 概要    : 受験者ファイルのＲＥＡＤを行う。会場数の取得
	'         :
	' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=READ成功。
	'         :                    False=READ失敗。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2005.12.02  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function OSC_JUKENF_Read() As Short
		
		Dim pRetVal As Short
		Dim dbRec As ADODB.Recordset
		Dim pMAXLen As Short
		Dim pIX As Short
		
		OSC_JUKENF_Read = False
		mSUJU = 0
		
		dbRec = New ADODB.Recordset
		
		'
		' 受験者ファイルＲＥＡＤ
		'
		
		SqlText = "select COUNT(JKNOGA) as SUJU "
		SqlText = SqlText & " from OSC_JUKENF "
		SqlText = SqlText & " where JKCDGA  = 1 "
		SqlText = SqlText & "   and JKCDGB  = " & CStr(bCDGB)
		SqlText = SqlText & "   and JKCDGK  = " & CStr(mCDGK)
		SqlText = SqlText & "   and JKNENDO = " & CStr(mNENDO)
		SqlText = SqlText & "   and JKKBSK = " & CStr(mKBSK)
		SqlText = SqlText & "   and JKYEAR = " & CStr(mYEAR)
		SqlText = SqlText & "  group by JKCDGA, JKCDGB, JKCDGK, JKNENDO, JKKBSK, JKYEAR "
		SqlText = SqlText & " order by JKCDGA, JKCDGB, JKCDGK, JKNENDO, JKKBSK, JKYEAR "
		
		On Error Resume Next
		dbRec.Open(SqlText, dbCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Err.Number <> 0 Then
			mMsgText = "受験者ファイルのＲＥＡＤでエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
			Err.Clear()
			On Error GoTo 0
			GoTo RecordSet_Close
		End If
		On Error GoTo 0
		
		If dbRec.EOF = True Then
			GoTo RecordSet_Close
		End If
		
		
		Do While dbRec.EOF = False
			
			'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
			mSUJU = IIf(IsDbNull(dbRec.Fields("SUJU").Value), 0, dbRec.Fields("SUJU").Value)
			
			On Error Resume Next
			dbRec.MoveNext()
			
			If Err.Number <> 0 Then
				mMsgText = "受験者ファイルのＲＥＡＤでエラーが発生しました。"
				pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
				Err.Clear()
				On Error GoTo 0
				GoTo RecordSet_Close
			End If
			On Error GoTo 0
			
		Loop 
		
		OSC_JUKENF_Read = True
		
RecordSet_Close: 
		On Error Resume Next
		dbRec.Close()
		'UPGRADE_NOTE: オブジェクト dbRec をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbRec = Nothing
		On Error GoTo 0
		
	End Function
	
	
	' *******************************************************************************
	' 概要    : ＰｅｒｆｅｃｔＧｒｉｄの行を挿入する。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : pPGrid, I, PerfectGrid, グリッド。
	'         : pRow,   I, Long, 追加する行番号。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2000.12.12  REV.0001  鈴木  新規作成。
	' *******************************************************************************
	'
	Private Sub GridAddItem(ByVal pPGrid As AxPGRIDLib.AxPerfectGrid, ByVal pROW As Integer)
		
		Dim pRetVal As Short
		Dim pII As Integer
		
		If Not Grid1SpaceGyoCheck(PGrid, pPGrid.Items - 1) Then
			mMsgText = CStr(pPGrid.Items) & "行目にデータが入力されています。"
			mMsgText = mMsgText & Chr(10) & "行を挿入しますか？"
			pRetVal = MsgBox(mMsgText, MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, Me.Text)
			
			If pRetVal = MsgBoxResult.No Then
				Exit Sub
			End If
			
		End If
		
		pPGrid.RefreshLater = True
		
		pPGrid.RemoveItems(pPGrid.Items - 1, 1)
		'UPGRADE_NOTE: Text は CtlText にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		pPGrid.CtlText = pPGrid.get_CellText(pPGrid.Row, pPGrid.Col)
		
		pPGrid.TextAtAddItem = ""
		pPGrid.AddItems(pROW, 1)
		
		''Call GridGyoNOSet(pPGrid)
		
		For pII = 0 To (pPGrid.Items - 1)
			pPGrid.set_CellText(pII, -1, CStr(pII + 1))
		Next pII
		
		pPGrid.NextRow = pROW
		pPGrid.NextCol = pPGrid.Col
		
		pPGrid.RefreshLater = False
		'UPGRADE_NOTE: Refresh は CtlRefresh にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		pPGrid.CtlRefresh()
		
	End Sub
	
	' *******************************************************************************
	' 概要  : ステーションマスターをREADする。
	' 　　  :
	' 説明  :
	' 　　  :
	' 履歴  : 2005.12.12  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function OSC_STM_READ(ByVal pCDST As Integer, ByRef pNAME As String, ByRef pSW_MSG As Boolean, Optional ByRef pKBSP As Integer = 0) As Object
		
		Dim dbRec As ADODB.Recordset
		Dim pRetVal As Short
		Dim pMsgText As String
		Dim pIX As Integer
		Dim pTOP As Integer
		
		
		'UPGRADE_WARNING: オブジェクト OSC_STM_READ の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
		OSC_STM_READ = False
		
		pNAME = ""
		pKBSP = 0
		
		dbRec = New ADODB.Recordset
		
		SqlText = "select STCDST,STNMST,STKBSP "
		SqlText = SqlText & " from OSC_STM "
		
		SqlText = SqlText & " where STCDGA  = 1 "
		SqlText = SqlText & "   and STCDGB  = " & bCDGB
		SqlText = SqlText & "   and STNENDO = " & mNENDO
		
		
		SqlText = SqlText & "   and STCDGK = " & cboCDGK.get_ItemData(cboCDGK.ListIndex)
		
		'試験区分
		If optKBSK(0).Value = True Then
			SqlText = SqlText & " and STKBSK = 1 "
		ElseIf optKBSK(1).Value = True Then 
			SqlText = SqlText & " and STKBSK = 2 "
		End If
		
		SqlText = SqlText & " and STYEAR = " & inpYEAR.Value
		SqlText = SqlText & " and STCDST = " & inpCDST.Value
		SqlText = SqlText & "order by STCDGA, STCDGB, STCDGK, STNENDO, STKBSK, STYEAR, STCDST  "
		
		
		On Error Resume Next
		dbRec.Open(SqlText, dbCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Err.Number <> 0 Then
			pMsgText = "ステーションマスターＲＥＡＤでエラーが発生しました。"
			pRetVal = ADOErrDisp(pMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
			GoTo RecordSet_Close
		End If
		On Error GoTo 0
		
		If dbRec.EOF = True Then
			GoTo RecordSet_Close
		End If
		
		'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
		pNAME = IIf(IsDbNull(dbRec.Fields("STNMST").Value), "", dbRec.Fields("STNMST").Value)
		pKBSP = dbRec.Fields("STKBSP").Value '１の時、模擬患者評価あり
		
		'UPGRADE_WARNING: オブジェクト OSC_STM_READ の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
		OSC_STM_READ = True
		'-------------------------------------------------------------------------
RecordSet_Close: 
		
		On Error Resume Next
		dbRec.Close()
		'UPGRADE_NOTE: オブジェクト dbRec をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbRec = Nothing
		Err.Clear()
		On Error GoTo 0
		
	End Function
	
	' *******************************************************************************
	' 概要    : 評価項目マスター（事前登録用）のＲＥＡＤを行う。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=READ成功。
	'         :                    False=READ失敗。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2009.03.04  REV.0001  保延  新規作成。
	' *******************************************************************************
	'
	Private Function OSC_HYOKAKMT_Read(ByVal pNOTR As Integer) As Short
		
		Dim pRetVal As Short
		Dim pROW As Integer
		Dim pNM As String
		Dim dbRec As ADODB.Recordset
		Dim pMAXLen As Short
		Dim pIX As Short
		Dim pIndex As Short
		Dim pFMT As String
		Dim pMVL As String
		Dim pCOL As Integer
		
		OSC_HYOKAKMT_Read = False
		
		dbRec = New ADODB.Recordset
		
		'
		' 評価項目マスターＲＥＡＤ
		'
		SqlText = "select * "
		SqlText = SqlText & " from OSC_HYOKAKM_TEMP "
		SqlText = SqlText & " where HYTCDGA  = 1 "
		SqlText = SqlText & "   and HYTCDGB  = " & CStr(bCDGB)
		SqlText = SqlText & "   and HYTCDGK  = " & CStr(mCDGK)
		SqlText = SqlText & "   and HYTNENDO = " & CStr(mNENDO)
		SqlText = SqlText & "   and HYTNOTR  = " & CStr(pNOTR)
		SqlText = SqlText & " order by HYTCDGA, HYTCDGB, HYTCDGK,HYTNENDO, HYTNOTR, HYTYEAR, HYTCDST, HYTKBHY, HYTNOSQ "
		
		On Error Resume Next
		dbRec.Open(SqlText, dbCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター(事前登録用）のＲＥＡＤでエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
			Err.Clear()
			On Error GoTo 0
			GoTo RecordSet_Close
		End If
		On Error GoTo 0
		
		If dbRec.EOF = True Then
			GoTo RecordSet_Close
		End If
		
		pROW = -1
		Do While dbRec.EOF = False
			
			pROW = pROW + 1
			PGrid.set_CellValueByName(pROW, "CDGP", dbRec.Fields("HYTCDGP").Value)
			
			PGrid.set_CellCheckedByName(pROW, "KBDAI", IIf(dbRec.Fields("HYTKBDAI").Value = 1, True, False))
			PGrid.set_CellCheckedByName(pROW, "KBCHU", IIf(dbRec.Fields("HYTKBCHU").Value = 1, True, False))
			PGrid.set_CellCheckedByName(pROW, "KBKR", IIf(dbRec.Fields("HYTKBKR").Value = 1, True, False))
			'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
			PGrid.set_CellTextByName(pROW, "MONDAI", IIf(IsDbNull(dbRec.Fields("HYTMONDAI").Value), "", dbRec.Fields("HYTMONDAI").Value))
			'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
			PGrid.set_CellTextByName(pROW, "NO", IIf(IsDbNull(dbRec.Fields("HYTNO").Value), "", dbRec.Fields("HYTNO").Value))
			PGrid.set_CellValueByName(pROW, "SUKA", dbRec.Fields("HYTSUKA").Value)
			PGrid.set_CellCheckedByName(pROW, "KB1", IIf(dbRec.Fields("HYTKB1").Value = 1, True, False))
			PGrid.set_CellCheckedByName(pROW, "KB2", IIf(dbRec.Fields("HYTKB2").Value = 1, True, False))
			PGrid.set_CellCheckedByName(pROW, "KB3", IIf(dbRec.Fields("HYTKB3").Value = 1, True, False))
			PGrid.set_CellCheckedByName(pROW, "KB4", IIf(dbRec.Fields("HYTKB4").Value = 1, True, False))
			
			Call ENABLE_Change(pROW)
			
			
			On Error Resume Next
			dbRec.MoveNext()
			
			If Err.Number <> 0 Then
				mMsgText = "評価項目マスター（事前登録用）のＲＥＡＤでエラーが発生しました。"
				pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
				Err.Clear()
				On Error GoTo 0
				GoTo RecordSet_Close
			End If
			On Error GoTo 0
		Loop 
		
		OSC_HYOKAKMT_Read = True
		
RecordSet_Close: 
		On Error Resume Next
		dbRec.Close()
		'UPGRADE_NOTE: オブジェクト dbRec をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbRec = Nothing
		On Error GoTo 0
		
	End Function
	
	' *******************************************************************************
	' 概要    : 評価項目マスターのＲＥＡＤを行う。
	'         :
	' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=READ成功。
	'         :                    False=READ失敗。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2009.03.04  REV.0001  保延  新規作成。
	' *******************************************************************************
	'
	Private Function OSC_HYOKAKMGT_READ(ByVal pNOTR As Integer) As Short
		
		Dim pRetVal As Short
		Dim pROW As Integer
		Dim pNM As String
		Dim dbRec As ADODB.Recordset
		Dim pMAXLen As Short
		Dim pIX As Short
		Dim pIndex As Short
		Dim pFMT As String
		Dim pMVL As String
		Dim pCOL As Integer
		
		OSC_HYOKAKMGT_READ = False
		
		dbRec = New ADODB.Recordset
		
		'
		' 評価項目マスターＲＥＡＤ（概略評価）
		'
		SqlText = "select * "
		SqlText = SqlText & " from OSC_HYOKAKMG_TEMP "
		SqlText = SqlText & " where HGTCDGA  = 1 "
		SqlText = SqlText & "   and HGTCDGB  = " & CStr(bCDGB)
		SqlText = SqlText & "   and HGTCDGK  = " & CStr(mCDGK)
		SqlText = SqlText & "   and HGTNENDO = " & CStr(mNENDO)
		SqlText = SqlText & "   and HGTNOTR  = " & CStr(pNOTR)
		SqlText = SqlText & " order by HGTCDGA, HGTCDGB, HGTCDGK,HGTNENDO, HGTNOTR, HGTYEAR, HGTCDST, HGTKBHY, HGTNOSQ "
		
		On Error Resume Next
		dbRec.Open(SqlText, dbCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター（概略評価）（事前登録用）のＲＥＡＤでエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
			Err.Clear()
			On Error GoTo 0
			GoTo RecordSet_Close
		End If
		On Error GoTo 0
		
		If dbRec.EOF = True Then
			GoTo RecordSet_Close
		End If
		
		pROW = -1
		Do While dbRec.EOF = False
			
			pROW = pROW + 1
			PGrid2.set_CellValueByName(pROW, "CDGP", dbRec.Fields("HGTCDGP").Value)
			
			PGrid2.set_CellCheckedByName(pROW, "KBDAI", IIf(dbRec.Fields("HGTKBDAI").Value = 1, True, False))
			PGrid2.set_CellCheckedByName(pROW, "KBKR", IIf(dbRec.Fields("HGTKBKR").Value = 1, True, False))
			PGrid2.set_CellValueByName(pROW, "HYOKA", dbRec.Fields("HGTHYOKA").Value)
			'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
			PGrid2.set_CellTextByName(pROW, "NAIYO", IIf(IsDbNull(dbRec.Fields("HGTNAIYO").Value), "", dbRec.Fields("HGTNAIYO").Value))
			PGrid2.set_CellCheckedByName(pROW, "KB1", IIf(dbRec.Fields("HGTKB1").Value = 1, True, False))
			
			Call ENABLE_Change2(pROW)
			
			On Error Resume Next
			dbRec.MoveNext()
			
			If Err.Number <> 0 Then
				mMsgText = "評価項目マスター（事前登録用）のＲＥＡＤでエラーが発生しました。"
				pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
				Err.Clear()
				On Error GoTo 0
				GoTo RecordSet_Close
			End If
			On Error GoTo 0
		Loop 
		
		OSC_HYOKAKMGT_READ = True
		
RecordSet_Close: 
		On Error Resume Next
		dbRec.Close()
		'UPGRADE_NOTE: オブジェクト dbRec をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbRec = Nothing
		On Error GoTo 0
		
	End Function
	
	
	' *******************************************************************************
	' 概要    : 評価項目マスターのＲＥＡＤを行う。(問題点)
	'         :
	' ﾊﾟﾗﾒｰﾀ  : 戻り値, O, Integer, True=READ成功。
	'         :                    False=READ失敗。
	'         :
	' 説明    :
	'         :
	' 履歴    : 2005.12.12  REV.0001  吉田  新規作成。
	' *******************************************************************************
	'
	Private Function OSC_HYOKAKMMT_READ(ByVal pNOTR As Integer) As Short
		
		Dim pRetVal As Short
		Dim pROW As Integer
		Dim pNM As String
		Dim dbRec As ADODB.Recordset
		Dim pMAXLen As Short
		Dim pIX As Short
		Dim pIndex As Short
		Dim pFMT As String
		Dim pMVL As String
		Dim pCOL As Integer
		
		OSC_HYOKAKMMT_READ = False
		
		
		
		
		dbRec = New ADODB.Recordset
		
		'
		' 評価項目マスターＲＥＡＤ(問題点)
		'
		SqlText = "select * "
		SqlText = SqlText & " from OSC_HYOKAKMM_TEMP "
		SqlText = SqlText & " where HMTCDGA  = 1 "
		SqlText = SqlText & "   and HMTCDGB  = " & CStr(bCDGB)
		SqlText = SqlText & "   and HMTCDGK  = " & CStr(mCDGK)
		SqlText = SqlText & "   and HMTNENDO = " & CStr(mNENDO)
		SqlText = SqlText & "   and HMTNOTR  = " & CStr(pNOTR)
		SqlText = SqlText & " order by HMTCDGA, HMTCDGB, HMTCDGK, HMTNENDO, HMTNOTR, HMTYEAR, HMTCDST, HMTKBHY, HMTNOSQ "
		
		On Error Resume Next
		dbRec.Open(SqlText, dbCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		If Err.Number <> 0 Then
			mMsgText = "評価項目マスター（概略評価）（事前登録用）のＲＥＡＤでエラーが発生しました。"
			pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
			Err.Clear()
			On Error GoTo 0
			GoTo RecordSet_Close
		End If
		On Error GoTo 0
		
		If dbRec.EOF = True Then
			GoTo RecordSet_Close
		End If
		
		pROW = -1
		Do While dbRec.EOF = False
			
			pROW = pROW + 1
			PGrid3.set_CellValueByName(pROW, "CDGP", dbRec.Fields("HMTCDGP").Value)
			
			PGrid3.set_CellCheckedByName(pROW, "KBDAI", IIf(dbRec.Fields("HMTKBDAI").Value = 1, True, False))
			PGrid3.set_CellCheckedByName(pROW, "KBKR", IIf(dbRec.Fields("HMTKBKR").Value = 1, True, False))
			'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' をクリックしてください。
			PGrid3.set_CellTextByName(pROW, "MONDAI", IIf(IsDbNull(dbRec.Fields("HMTMONDAI").Value), "", dbRec.Fields("HMTMONDAI").Value))
			PGrid3.set_CellCheckedByName(pROW, "KB1", IIf(dbRec.Fields("HMTKB1").Value = 1, True, False))
			
			
			Call ENABLE_Change3(pROW)
			
			On Error Resume Next
			dbRec.MoveNext()
			
			If Err.Number <> 0 Then
				mMsgText = "評価項目マスター（事前登録用）のＲＥＡＤでエラーが発生しました。"
				pRetVal = ADOErrDisp(mMsgText, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, mMBOXTitle, Err.Description)
				Err.Clear()
				On Error GoTo 0
				GoTo RecordSet_Close
			End If
			On Error GoTo 0
		Loop 
		
		OSC_HYOKAKMMT_READ = True
		
RecordSet_Close: 
		On Error Resume Next
		dbRec.Close()
		'UPGRADE_NOTE: オブジェクト dbRec をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbRec = Nothing
		On Error GoTo 0
		
	End Function
End Class