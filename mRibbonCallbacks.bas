Attribute VB_Name = "mRibbonCallbacks"
Option Explicit
Option Private Module

Dim g_Input As String
Dim g_Result As Double
Dim g_Ribbond As IRibbonUI

Public VygruzkaFormy As Boolean
Public PokazatSoobshceniye As Boolean

Public Const PUT_K_SHABLONAM As String = "\\miron4\PRK\33300\_ГРУППА НОРМИРОВАНИЯ\Расчёт норм\Для надстройки\Шаблоны\"

Sub rbnOtkrytPlankaDatron(Control As IRibbonControl)
    Workbooks.Open (PoluchitPutKShablonam & "Планка с надписью Datron.xltm")
End Sub

Sub Ribbon_Onload(ribbon As IRibbonUI)
    Set g_Ribbond = ribbon
End Sub

Private Sub Ribbon_TextChanged(Control As IRibbonControl, text As String)
    g_Input = text
    If IsNumeric(g_Input) Then
        g_Result = Round((1.3 * Ustanov_NaStoleSUporomBezKrepleniya(CDbl(g_Input), False)) + IIf(CDbl(g_Input) > 20, 10, 0), 1)
    Else
        g_Result = 0
    End If

    g_Ribbond.InvalidateControl "etTransport"
    
    
End Sub

Private Sub rbnOtkrytPosleLazerPressa(Control As IRibbonControl)
    Workbooks.Open (PoluchitPutKShablonam & "После пробивочно-вырезной.xltm")
End Sub

Private Sub ShowFrmGalvanika(Control As IRibbonControl)
    Call ShowForm(frmGalvanika)
End Sub

Private Sub ShowFrmZatirka(Control As IRibbonControl)
    Call ShowForm(frmZatirka)
End Sub

Private Sub rbnUstanovka(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = g_Result
End Sub

Private Sub ShowFrmPlanki(Control As IRibbonControl)
    Call ShowForm(frmPlankiSNadpisyu)
End Sub

Private Sub ShowFrmSpeedy(Control As IRibbonControl)
    Call ShowForm(frmSpeedy)
End Sub
Private Sub ShowFrmDatron(Control As IRibbonControl)

    Call ShowForm(frmDatron)
    Call ZapisatKnopku(frmDatron)
    
End Sub

Private Sub ShowForm(frm As Object)
    
    On Error Resume Next
    With frm
        .Left = Application.Left + 10
        .Show
    End With
    
    On Error GoTo 0
    
    If Environ("UserName") <> "hardf" Then Call ZapisatKnopku(frm)
    
End Sub

Private Sub ZapisatKnopku(frm As Object)

    Dim wbUsers As Workbook
    Dim wsUsers As Worksheet
    Dim lastRow As Long
    
    Dim FilePath As String
    FilePath = "\\miron4\PRK\33300\_ГРУППА НОРМИРОВАНИЯ\Расчёт норм\Users.xlsx"
    
    On Error Resume Next
    Application.DisplayAlerts = False
    
    If Dir(FilePath) <> "" Then
    
        Set wbUsers = Workbooks.Open(FilePath)
        wbUsers.Windows(1).Visible = False
        Set wsUsers = wbUsers.Worksheets(2)
        
        With wsUsers
            lastRow = WorksheetFunction.CountA(.Columns(1)) + 1
            .Cells(lastRow, 1) = frm.Name
            .Cells(lastRow, 2) = Application.UserName
            .Cells(lastRow, 3) = Now
        End With
        
        If wbUsers.ReadOnly Then
            wbUsers.Close (False)
        Else
            wbUsers.Close (True)
        End If
        
    End If
    
    On Error GoTo 0
    Application.DisplayAlerts = True

End Sub

Private Sub ShowFrmObrabotkaOtverstiy(Control As IRibbonControl)
    Call ShowForm(frmObrabotkaOtverstiy)
End Sub
Private Sub ShowFrmZametki(Control As IRibbonControl)
    Call ShowForm(frmZametki)
End Sub

Private Sub ShowFrmArea(Control As IRibbonControl)
    Call ShowForm(frmRaschotPloshchadi)
End Sub
Private Sub ShowFrmMass(Control As IRibbonControl)
    Call ShowForm(frmMass)
End Sub

Private Sub ShowFrmSvarka(Control As IRibbonControl)
    Call ShowForm(frmSvarka)
End Sub

Private Sub ShowfrmTok(Control As IRibbonControl)
    Call ShowForm(frmTokFrez)
End Sub

Private Sub ShowFrmGaltovka(Control As IRibbonControl)
    Call ShowForm(frmGaltovka)
End Sub

Private Sub rbnIzSpetsifikatsii(Control As IRibbonControl)
    Call mIzSpetsifikatsii.ZapuskFormyIzSpetsifikatsii(ActiveWorkbook, True)
End Sub

Private Sub ShowFrmTermografiya(Control As IRibbonControl)
    Call ShowForm(frmTermografiya)
End Sub
Private Sub ShowFrmTermografiya_2(Control As IRibbonControl)
    Call ShowForm(frmTermografiya_2)
End Sub
Private Sub ShowFrmElektroeroziya(Control As IRibbonControl)
    Call ShowForm(frmElektroEroziya)
End Sub

Private Sub ShowFrmDopuski(Control As IRibbonControl)
    Call ShowForm(frmDopuski)
End Sub

Private Sub ShowFrmOtrezGilNoz(Control As IRibbonControl)
    Call ShowForm(frmOtrezkaNaGilotinNoznicah)
End Sub
Private Sub ShowFrmTmash(Control As IRibbonControl)
        
    Call ShowForm(frmTmash)
End Sub
Private Sub ShowFrmTpz(Control As IRibbonControl)

    Call ShowForm(frmRaschotTpz)
End Sub
Private Sub ShowFrmCircularPil(Control As IRibbonControl)
    Call ShowForm(frmOtrezkaNaCircularnihPilah)
End Sub

Private Sub ShowFrmLentOtrez(Control As IRibbonControl)
    Call ShowForm(frmLentochnoOtreznieStanki)
End Sub

Private Sub ShowFrmMarkirovanie(Control As IRibbonControl)
    Call ShowForm(frmMarkiravanie)
End Sub

Private Sub ShowFrmGidroAbraziv(Control As IRibbonControl)
    Call ShowForm(frmGidroabraziv)
End Sub

Private Sub ShowFrmTermoObrabotka(Control As IRibbonControl)
    Call ShowForm(frmTermoObrabotka)
End Sub
Private Sub ShowFrmLaserGrav(Control As IRibbonControl)
    frmLazer.Show
End Sub

Private Sub ShowFrmOrkashyvaniye(Control As IRibbonControl)
    Call ShowForm(frmOkrashivaniye)
End Sub

Private Sub ShowFrmGravirovka(Control As IRibbonControl)
    Call ShowForm(frmGravirovka)
End Sub

Private Sub ShowFrmMetallografika(Control As IRibbonControl)
    Call ShowForm(frmMetallografika)
End Sub
Private Sub ShowFrmPolirovaniye(Control As IRibbonControl)
    Call ShowForm(frmPolirovaniye)
End Sub
Private Sub rbnKomplektovaniye33208(Control As IRibbonControl)
    Call RaschotSborkiIzBufera
End Sub
Private Sub ShowFrmOkrashivaniye(Control As IRibbonControl)
    Call ShowForm(frmOkrashivaniye)
End Sub
Private Sub ShowFrmKIM(Control As IRibbonControl)
    Call ShowForm(frmKIM)
End Sub

Private Sub ShowFrmPromyvka(Control As IRibbonControl)
    Call ShowForm(frmPromyvka)
End Sub

Private Sub ShowFrmShlifovaniye(Control As IRibbonControl)

    MsgBox "Временно не работает. Воспользуйтесь ""Расче норм""", vbInformation, "Шлифование"
    Exit Sub
    Call ShowForm(frmShlifonaiye)
End Sub

Private Sub ShowFrmKoordinat(Control As IRibbonControl)
    Call ShowForm(frmKoordinat)
End Sub

Private Sub ShowFrmPeskostruy(Control As IRibbonControl)
    Call ShowForm(frmPeskostruy)
End Sub


Private Sub rbnCloseAllForm(Control As IRibbonControl)
    If UserForms.Count > 0 Then
        Dim i As Integer
        For i = UserForms.Count - 1 To 0 Step -1
            Unload UserForms(i)
        Next
    End If
End Sub

Private Sub rbnOtkrytNormativy(Control As IRibbonControl)
    Workbooks.Open ("\\miron4\PRK\33300\_ГРУППА НОРМИРОВАНИЯ\Расчёт норм\НОРМАТИВЫ.xltm")
End Sub

Private Sub rbnOtkrytTipovyyeDetali(Control As IRibbonControl)
    Shell "explorer.exe \\miron4\PRK\33300\_ГРУППА НОРМИРОВАНИЯ\Расчёт норм\Для надстройки\Типовые", vbMaximizedFocus
End Sub

Private Sub rbnOtkrytRaschetNorm(Control As IRibbonControl)
    Workbooks.Open ("\\miron4\PRK\33300\_ГРУППА НОРМИРОВАНИЯ\Расчёт норм\Расчет норм.xltm")
End Sub

Private Sub rbnOtkrytRaschetNormVremeni2(Control As IRibbonControl)
    Workbooks.Open (PUT_K_SHABLONAM & "Расчет норм времени.xltm")
End Sub

Private Sub rbnOtkrytZIP(Control As IRibbonControl)
    Workbooks.Open (PUT_K_SHABLONAM & "Комплект ЗИП, упаковка.xltm")
End Sub


Private Sub rbnOtkrytPlaty(Control As IRibbonControl)
    Dim wb_spetsifikatsiya As Workbook
    Dim wb_platy As Workbook
    Set wb_spetsifikatsiya = ActiveWorkbook
    Set wb_platy = Workbooks.Open(PUT_K_SHABLONAM & "Платы и узлы.xltx")

End Sub

Private Sub rbnOtkrytStoyki(Control As IRibbonControl)
    Workbooks.Open (PUT_K_SHABLONAM & "Электромонтаж блоков, стоек.xltm")
End Sub

Private Sub rbnOtkrytZhguty(Control As IRibbonControl)
    Workbooks.Open ("\\miron4\PRK\33300\_ГРУППА НОРМИРОВАНИЯ\Расчёт норм\Паспорт ЭМ\Жгуты.xltm")
End Sub

Private Sub rbnOtkrytKalkulyatsiya(Control As IRibbonControl)
    Workbooks.Open ("\\miron4\PRK\33300\_ГРУППА НОРМИРОВАНИЯ\Расчёт норм\Шаблоны калькуляций.xltx")
End Sub


