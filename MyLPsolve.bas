Attribute VB_Name = "MyLPsolve"
Option Explicit

Const MLDATA& = 90 '������������ ���������� �������� ������
Const MLZGT& = 10 '������������ ���������� ���������
Const MLSTACK& = 2000000 '������������ ������ �����
Const MLGENVAR& = 200000 '������������ ���������� ��������������� ���������
Const MAXCOLUMN As Long = 50000 '����������� ��������� ���-�� �������� (���� �������) ��� lpsolve

Private bIsSolve As Boolean
Private bNoShowAlert As Boolean

Private lpsolve As lpsolve55

Private Function ctrlcfunc(ByVal lp As Long, ByVal userhandle As Long) As Long
    'If set to True, then solve is aborted and returncode will indicate this.
    'ctrlcfunc = True
End Function

Sub lpCSP() '������� ������ CSP �������� �����������������
    Dim inpRng As Range, inpRng2 As Range, outRng As Range
    Dim i&, j&, k&, m&, l&, nd&, nz&, ns&, shrez&, krm&, iSpeed&
    Dim shArr(), dtArr&(), dtNewArr&(), zgArr&(), zgNewArr&(), lpArr()
    Dim smDt&, smZg&, maxDt&, minDt&
    Dim nsh&, ncsp&, v, txt$, out, tmp&, dOst
    Dim bGroupe As Boolean, bGraph As Boolean, b2prof As Boolean
    Dim ostMax&, sOst&, ost&, delOst&, bstNCspKrm&, smOst&
    Dim bstNCsp&, bstDelOst&, bstMxOst&, t, bstLpArr
    Dim iTimer!, iter&

    On Error GoTo EndSolution
    Application.EnableCancelKey = xlErrorHandler

    iTimer = Timer

    Set lpsolve = New lpsolve55
    lpsolve.Init Application.ActiveWorkbook.Path

    Set inpRng = [c26] '�������� ������, �������
    Set inpRng2 = [c4] '�������� ������, ���������
    Set outRng = [k4] '���� �������� �������

    shrez = Val([d15])
    krm = Val([d16])
    bGroupe = [e15] = True
    bGraph = [e16] = True
    delOst = Val([d17]) '������� �������
    b2prof = [e17] = True '������� ������� � ��� �������

    iSpeed = Val([h15]) '�������� ����������
    If iSpeed < 1 Then iSpeed = 1
    If iSpeed > 4 Then iSpeed = 5

    Call ClearSolve
    Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
    ActiveSheet.Unprotect

    '-----------------------------------------------------------------------------------------------------------------'

    For i = 1 To MLDATA '���� �������
        m = Val(inpRng.Offset(i - 1, 0)) '������ ������
        l = Val(inpRng.Offset(i - 1, 1)) '���-�� �������
        If m > 0 And l > 0 Then
            smDt = smDt + (m + shrez) * l '��������� ����� ���� ������� � ������ ����
            For j = 1 To nd
                If dtArr(1, j) = m Then Exit For
            Next j
            If j > nd Then '���� ������ ������ ������� ��� �� ����, �� ��������� ��
                nd = nd + 1 '����������� ���������� ��������� �������
                ReDim Preserve dtArr&(1 To 2, 1 To nd) '����������� ������ �������
            End If
            dtArr(1, j) = m '���������� ������
            dtArr(2, j) = dtArr(2, j) + l '���������� ���-��
        End If
    Next i

    '-----------------------------------------------------------------------------------------------------------------'

    If b2prof Then '���������, ������ �� ���-�� �������
        For i = 1 To nd
            If dtArr(2, i) Mod 2 Then b2prof = False: Exit For
        Next i
        If b2prof Then
            For i = 1 To nd
                dtArr(2, i) = dtArr(2, i) \ 2
            Next i
        Else
            [e17] = False
            MsgBox "����� ������� ��������, ������� ����� ���������� � ���� �������"
        End If
    End If

    '-----------------------------------------------------------------------------------------------------------------'

    For i = 1 To MLZGT '���� ���������
        m = Val(inpRng2.Offset(i - 1, 0)) '������ ���������
        '���� ���-�� ��������� �� �������, �� ��������� ������ ���-�� "������" ����������
        If inpRng2.Offset(i - 1, 1) = "" And m > 0 Then l = fGreedyAlgo(dtArr, m, shrez, krm) Else l = Val(inpRng2.Offset(i - 1, 1)) \ IIf(b2prof, 2, 1) '���-�� ���������
        If m > 0 And l > 0 Then
            For j = 1 To nz
                If zgArr(1, j) = m Then Exit For
            Next j
            If j > nz Then '���� ��������� ������ ������� ��� �� ����, �� ��������� ��
                nz = nz + 1 '����������� ���������� ��������� ���������
                ReDim Preserve zgArr&(1 To 2, 1 To nz) '����������� ������ �������
            End If
            zgArr(1, j) = m '���������� ������
            zgArr(2, j) = zgArr(2, j) + l '���������� ���-��
        End If
    Next i
    If nz = 0 Or nd = 0 Then Exit Sub '������ �����������

    '-----------------------------------------------------------------------------------------------------------------'

    ProgressLog.Show False
    ProgressLog.TxtClear
    ProgressLog.TxtAdd = "������� CSP �������� �����������������": DoEvents

    '-----------------------------------------------------------------------------------------------------------------'

    '��������� ����� ������������� ������� (����������� �� ������), � ������������ � ��� ����������� �������
    For k = iSpeed + 2 To 1 Step -1
        nsh = fGenerateSum(dtArr, zgArr, shArr(), shrez, krm, Choose(k, 0, 20, 50, 100, 150, 200, 250))
        Debug.Print nsh
        If nsh > 0 Then '���� ���� ����� �������, �� ���������� ������
            For i = 1 To IIf(k > 0, 1, Choose(iSpeed, 2, 3, 4, 5, 6)) '������ ��������� ������� �������
                iter = iter + 1
                ProgressLog.TxtAdd = "��������: " & iter & vbLf & "���� �������: " & nsh: DoEvents
                ncsp = fCSP_LPsolve(dtArr, zgArr, shArr, lpArr, shrez, krm, Choose(iSpeed, 3, 5, 10, 15, 60)) '��������� LP �������
                If ncsp > 0 Then
                    ostMax = 0 '������������ �������
                    sOst = 0 '����� �������� ��������
                    smOst = 0 '����� ��������
                    For j = 1 To UBound(lpArr, 2)
                        ost = lpArr(3, j) - lpArr(2, j)
                        If ost > ostMax Then ostMax = ost '���������� ������������ �������
                        If ost >= delOst And delOst <> 0 Then sOst = sOst + ost '��������� ����� �������� ��������
                        smOst = smOst + ost '����� ��������
                    Next j
                    If bstNCsp = 0 Or (ncsp < bstNCsp Or (ncsp = bstNCsp And bstDelOst > sOst) Or (ncsp = bstNCsp And bstDelOst = sOst And bstMxOst < ostMax)) Then
                        bstNCsp = ncsp '���������� ��������� �������
                        bstNCspKrm = smOst
                        bstDelOst = sOst
                        bstMxOst = ostMax
                        bstLpArr = lpArr
                    End If
                    ProgressLog.TxtAdd = "����� ���������: " & ncsp & vbLf & "������������ �������: " & bstMxOst & vbLf: DoEvents
                End If
                If ProgressLog.Stoped Then GoTo EndSolution
            Next i
        End If
    Next k

    '-----------------------------------------------------------------------------------------------------------------'

    If bstNCsp > 0 And iSpeed > 2 Then '������������� �������
        ReDim zgNewArr&(1 To 2, 1 To nz) '������� ����� ������ ���������
        For i = 1 To nz
            zgNewArr(1, i) = zgArr(1, i)
            If i = 1 Or maxDt > zgArr(1, i) - krm - 1 Then maxDt = zgArr(1, i) - krm - 1
            For j = 1 To UBound(bstLpArr, 2)
                If bstLpArr(3, j) = zgNewArr(1, i) Then zgNewArr(2, i) = zgNewArr(2, i) + 1
        Next j, i
        minDt = bstMxOst '���������� ��������� �������
        If maxDt > bstNCspKrm Then maxDt = bstNCspKrm '����������� ��������� �������

        dtNewArr = dtArr '������� ����� ������ � ��������
        ReDim Preserve dtNewArr&(1 To 2, 1 To nd + 1) '��������� �������

        For i = 1 To 14 '���������, ���� �� ����� ������� �������
            dtNewArr(1, nd + 1) = (maxDt + minDt) \ 2 - shrez
            dtNewArr(2, nd + 1) = 1
            '������� ����� ����� �������
            nsh = fGenerateSum(dtNewArr, zgNewArr, shArr(), shrez, krm) ', bstNCspKrm - dtNewArr(1, nd + 1) - shrez)

            iter = iter + 1
            ProgressLog.TxtAdd = "��������: " & iter & vbLf & "���� �������: " & nsh: DoEvents
            Debug.Print nsh, i, minDt, maxDt, bstNCspKrm - dtNewArr(1, nd + 1) - shrez

            If nsh > 0 Then
                ncsp = fCSP_LPsolve(dtNewArr, zgNewArr, shArr, lpArr, shrez, krm, 5, nd)
                If ncsp <= 0 Then ncsp = fCSP_LPsolve(dtNewArr, zgNewArr, shArr, lpArr, shrez, krm, 10, nd)
                If ncsp > 0 Then
                    ostMax = 0 '������������ �������
                    sOst = 0 '����� �������� ��������
                    smOst = 0 '����� ��������
                    For j = 1 To UBound(lpArr, 2)
                        ost = lpArr(3, j) - lpArr(2, j)
                        If ost > ostMax Then ostMax = ost '���������� ������������ �������
                        If ost >= delOst And delOst <> 0 Then sOst = sOst + ost '��������� ����� �������� ��������
                        smOst = smOst + ost '����� ��������
                    Next j
                    bstNCsp = ncsp '���������� �������
                    bstNCspKrm = smOst
                    bstDelOst = sOst
                    bstMxOst = ostMax
                    bstLpArr = lpArr
                    minDt = bstMxOst
                Else
                    maxDt = dtNewArr(1, nd + 1) + shrez
                End If
                ProgressLog.TxtAdd = "����� ���������: " & ncsp & vbLf & "������������ �������: " & bstMxOst & vbLf: DoEvents
            End If
            If maxDt - minDt < 2 Then Exit For
            If ProgressLog.Stoped Then GoTo EndSolution
        Next i
    End If

    '-----------------------------------------------------------------------------------------------------------------'

EndSolution:
    Application.EnableCancelKey = xlInterrupt
    ProgressLog.Hide

    '-----------------------------------------------------------------------------------------------------------------'

    If bstNCsp > 0 Then '���� ������� ���� �������
        out = fGenSolution(dtArr, bstLpArr, shrez, krm, bGroupe, b2prof) '��������� ������� ��� ������ �� ����
        ns = UBound(out, 2)
        outRng.Resize(ns, 4) = Application.WorksheetFunction.Transpose(out) '������� ��� �� ����
        outRng.Offset(0, 4).Resize(ns, 1).FormulaR1C1 = "=RC[-4]-RC[-3]"
        outRng.Resize(ns, 5).Borders.LineStyle = 1 '� ������������
        outRng.Resize(ns, 5).Borders(xlInsideHorizontal).Weight = xlHairline
        outRng.Resize(ns).Font.Bold = True
        outRng.Offset(0, 2).Resize(ns, 2).Font.Bold = True
        outRng.Offset(0, 2).Resize(ns).NumberFormat = "0"" ��."""
        outRng.Offset(0, 3).EntireColumn.AutoFit
        '-------------------------------------------------
        k = 0 '��������� ������ ��������
        ReDim dOst(1 To 100, 1 To 2)
        For i = 1 To MLZGT
            If inpRng2.Offset(i - 1, 0) <> "" Then
                If inpRng2.Offset(i - 1, 1) = "" Or inpRng2.Offset(i - 1, 3) > 0 Then
                    For j = 1 To k
                        If dOst(j, 1) = inpRng2.Offset(i - 1, 1) Then Exit For
                    Next j
                    If j > 100 Then Exit For
                    If k < j Then k = j: dOst(j, 1) = inpRng2.Offset(i - 1, 0)
                    If inpRng2.Offset(i - 1, 1) = "" Then dOst(j, 2) = Empty Else dOst(j, 2) = dOst(j, 2) + inpRng2.Offset(i - 1, 3)
                End If
            End If
        Next i
        If delOst > 0 Then
            For i = 1 To ns
                If out(1, i) - out(2, i) >= delOst Then
                    For j = 1 To k
                        If dOst(j, 1) = outRng.Offset(i - 1, 4) Then Exit For
                    Next j
                    If j > 100 Then Exit For
                    If k < j Then k = j: dOst(j, 1) = outRng.Offset(i - 1, 4)
                    dOst(j, 2) = dOst(j, 2) + outRng.Offset(i - 1, 2)
                End If
            Next i
        End If
        '��������� ������� �� ��������
        For i = 2 To k - 1
            For j = i + 1 To k
                If dOst(i, 1) < dOst(j, 1) Then
                    tmp = dOst(j, 1): dOst(j, 1) = dOst(i, 1): dOst(i, 1) = tmp
                    tmp = dOst(j, 2): dOst(j, 2) = dOst(i, 2): dOst(i, 2) = tmp
                End If
        Next j, i
        inpRng2.Offset(0, 5).Resize(MLZGT, 2) = dOst
    '-----------------------------------------------------------------------------------------------------------------'

    Else
        If MsgBox("�� ������� ����� ������� �������� �����������������" & vbLf & _
                "��������� ������� ������������ �����������������?", vbYesNo) = vbYes Then
            Raskroy
            Exit Sub
        End If
    End If

    '-----------------------------------------------------------------------------------------------------------------'

    Debug.Print Timer - iTimer

    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    If bGraph Then Call OutGraph
    'Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Function fCSP_LPsolve(dtArr, zgArr, shArr, out(), Optional shrez& = 0, Optional krm& = 0, Optional mTime& = 1, Optional nnd& = 0) As Long
    Dim lp As Long
    Dim aRow() As Double
    Dim aCol() As Double
    Dim v, txt$, sm&

    Dim i&, j&, k&, nOut&, nd&, nz&, ns&, mx&
    Dim SolveOk As Double, SolveObj As Double, rndArr

    nd = UBound(dtArr, 2) '���-�� ����� �������
    nz = UBound(zgArr, 2) '���-�� ����� ���������
    ns = UBound(shArr, 2) '���-�� ������������ ����
    If nnd = 0 Then nnd = nd '���� �������� nnd ������ nd, �� ��� �������� ���� nnd ����� �������,
    '���������� ��� ������ ������������� �������

    If ns > MAXCOLUMN Then mx = MAXCOLUMN Else mx = ns '����������� ������������ ��������
    rndArr = GenRndArr(ns, mx) '������ ��������� ���������

    With lpsolve
        lp = .make_lp(nd, 0)
        .set_timeout lp, mTime '����������� ���������� �������� � ���.
        .put_abortfunc lp, AddressOf ctrlcfunc, 0

        For i = 1 To nd  '��������� ������� ��� �����
            .set_rh lp, i, dtArr(2, i): .set_constr_type lp, i, GE
        Next i
        For i = 1 To mx '���������� �������
            GenColFromStr aCol, CStr(shArr(1, rndArr(i))), nd, CLng(shArr(3, rndArr(i)))
            .add_column lp, aCol(1): .set_int lp, i, True
        Next i
        '��������� ������� ����������� ���������
        For j = 1 To nz
            If zgArr(2, j) >= 0 Then
                ReDim aRow(0 To mx)
                For i = 1 To mx
                    If shArr(3, rndArr(i)) = zgArr(1, j) Then aRow(i) = 1
                Next i
                .add_constraint lp, aRow(0), LE, zgArr(2, j)
            End If
        Next j
        '.write_lp lp, Application.ActiveWorkbook.Path & "\CSP.lp"
        SolveOk = .solve(lp) '���������� ������
        If SolveOk = 0 Or SolveOk = 1 Then '���� ������� �������
            SolveObj = .get_objective(lp) '���������������� �������
            ReDim aCol(1 To .get_Ncolumns(lp))
            .get_variables lp, aCol(1) '�������� ������ � ��������

            '������� �� ���� � �������� ������ ������
            ReDim ndt&(1 To nd)
            For i = 1 To mx
                If aCol(i) > 0 And aCol(i) < 1000000 Then
                    For j = 1 To aCol(i)
                        v = SplitSchem(CStr(shArr(1, rndArr(i))), nd)
                        txt = ""
                        sm = 0
                        For k = 1 To nnd
                            If v(k) Then If ndt(k) + v(k) > dtArr(2, k) Then v(k) = dtArr(2, k) - ndt(k)
                            ndt(k) = ndt(k) + v(k)
                            If v(k) Then
                                sm = sm + v(k) * (dtArr(1, k) + shrez)
                                txt = txt & "+" & IIf(v(k) > 1, v(k) & "*", "") & "[" & k & "]"
                            End If
                        Next k

                        nOut = nOut + 1
                        ReDim Preserve out(1 To 3, 1 To nOut)
                        out(1, nOut) = Mid$(txt, 2)
                        out(2, nOut) = sm + krm
                        out(3, nOut) = shArr(3, rndArr(i))
                    Next j
                End If
            Next i
        End If
        .delete_lp lp
    End With
    fCSP_LPsolve = SolveObj
End Function

Function fGenSolution(dtArr, lpArr, Optional shrez& = 0, Optional krm& = 0, Optional bGroupe As Boolean = False, Optional b2prof As Boolean = False)
'�������, ����������� ��������� � ��������
    Dim i&, j&, k&, nd&, ns&
    Dim v, txt$, t&

    nd = UBound(dtArr, 2) '���-�� ����� �������
    ns = UBound(lpArr, 2) '���-�� ���� � ��������

    ReDim out(1 To 4, 1 To ns)
    '��������� �������
    For i = 1 To ns
        v = SplitSchem(CStr(lpArr(1, i)), nd)
        txt = ""
        For j = 1 To nd
            If v(j) > 0 Then txt = txt & " + (" & dtArr(1, j) & IIf(shrez, "+" & shrez, "") & IIf(v(j) > 1, " - " & v(j) & " ��.", "") & ")"
        Next j
        out(1, i) = lpArr(3, i)
        out(2, i) = lpArr(2, i)
        out(3, i) = 1
        out(4, i) = "'=" & Mid$(txt, 4) & IIf(krm, " + [" & krm & "]", "")

        For j = 1 To i - 1
            If out(1, j) < out(1, i) Or (out(1, j) = out(1, i) And out(2, j) < out(2, i)) Or _
                    (out(1, j) = out(1, i) And out(2, j) = out(2, i) And out(4, j) < out(4, i)) Then
                t = out(1, j): out(1, j) = out(1, i): out(1, i) = t
                t = out(2, j): out(2, j) = out(2, i): out(2, i) = t
                txt = out(4, j): out(4, j) = out(4, i): out(4, i) = txt
            End If
    Next j, i

    If bGroupe Then '���� ����� ������������ �������
        k = 1
        For i = 2 To ns
            If out(1, i) = out(1, k) And out(4, i) = out(4, k) Then
                out(3, k) = out(3, k) + 1
            Else
                k = k + 1
                out(1, k) = out(1, i)
                out(2, k) = out(2, i)
                out(3, k) = out(3, i)
                out(4, k) = out(4, i)
            End If
        Next i
        ReDim Preserve out(1 To 4, 1 To k)
        If b2prof Then
            For i = 1 To k
                out(3, i) = out(3, i) * 2
            Next i
        End If
    End If
    fGenSolution = out
End Function

Function fGenerateSum(dt, zg, out(), Optional shrez& = 0, Optional krm& = 0, Optional mxKrm& = 0) As Long
'������� ��������� ���� ��������� ���������, ��� ���������� ������ �����
'����� MCH (������ �.), m-ch@mail.ru
'������������ � ������ Cutting stock problem
'dt - ������ �������
'zg - ������ ���������
'out() - ������������ ������ �� �������

    Dim i&, iz&, j&, k&, n&, m&, sl&
    Dim sm&, rws&
    Dim xd, xz

    n = UBound(dt, 2)

    ReDim out(1 To 3, 1 To MLGENVAR)
    For iz = 1 To UBound(zg, 2) '���������� ��� ���������
        sm = zg(1, iz) '������� �����
        If sm > 0 Then
            ReDim smi&(MLSTACK), smt$(MLSTACK), bNoRacio(MLSTACK) As Boolean '������� ��� �����
            sl = 0 '��������� ����� �����
            Do '������ ��������� ���� ��������� ������������
                For i = 1 To n
                    For j = 0 To sl
                        m = (sm - smi(j) - krm) \ (dt(1, i) + shrez)
                        If m > dt(2, i) Then m = dt(2, i)
                        For k = 1 To m
                            If k = 1 Then bNoRacio(j) = True Else bNoRacio(sl) = True '����� j/sl �� �������� "������������ �� ������"
                            sl = sl + 1
                            If sl > MLSTACK Then sl = sl - 1: Debug.Print "���������� ����": fGenerateSum = -2: Exit Do
                            smi(sl) = smi(j) + k * (dt(1, i) + shrez)
                            smt(sl) = smt(j) & "+" & IIf(k > 1, k & "*", "") & "[" & i & "]"
                        Next k
                        For k = i - 1 To 1 Step -1
                            If sm - smi(sl) - krm >= dt(1, k) + shrez Then bNoRacio(sl) = True: Exit For
                    Next k, j, i
            Loop While False

            '��������� ������ ��� ������ �� ����
            For i = 1 To sl
                If Not bNoRacio(i) Then
                    If mxKrm = 0 Or smi(i) >= sm - krm - mxKrm Then
                        rws = rws + 1
                        out(1, rws) = Mid$(smt(i), 2) '�����
                        out(2, rws) = smi(i) + krm '��������� �����
                        out(3, rws) = sm '������ ���������
                        If rws >= MLGENVAR Then fGenerateSum = -1: Exit For
                    End If
                End If
            Next i
            If rws >= MLGENVAR Then fGenerateSum = -1: Exit For
        End If
    Next iz
    If rws Then ReDim Preserve out(1 To 3, 1 To rws) Else fGenerateSum = -1 '�������� ������ �������, ���� ���� ���, �� ���������� -1
    Erase smi, smt, bNoRacio '������� �������
    If fGenerateSum = 0 Then fGenerateSum = rws
End Function

Function fGreedyAlgo&(dtArr, sm, Optional shrez = 0, Optional krm = 0)
'������� ���������� ������������ ���-�� ��������� ������ ����������
    Dim i&, j&, k&, n&, l&, t&, s&, aArr&(), bArr() As Boolean, bFlag As Boolean

    n = UBound(dtArr, 2)
    ReDim dt&(1 To n, 1 To 2)
    For i = 1 To n '����������
        dt(i, 1) = dtArr(1, i) + shrez
        dt(i, 2) = dtArr(2, i)
        For j = 1 To i - 1
            If dt(i, 1) > dt(j, 1) Then
                t = dt(i, 1): dt(i, 1) = dt(j, 1): dt(j, 1) = t
                t = dt(i, 2): dt(i, 2) = dt(j, 2): dt(j, 2) = t
            End If
    Next j, i

    For i = 1 To n '������������� � ���������� ������
        For j = 1 To dt(i, 2)
            k = k + 1
            ReDim Preserve aArr&(1 To k), bArr(1 To k) As Boolean
            aArr(k) = dt(i, 1)
    Next j, i

    Do '���������� ���-�� ������ ����������
        s = sm - krm
        bFlag = False
        l = l + 1
        For i = 1 To k
            If s >= aArr(i) And Not bArr(i) Then
                s = s - aArr(i)
                bArr(i) = True
                bFlag = True
            End If
        Next i
    Loop While bFlag And s < sm - krm
    fGreedyAlgo = l - 1
End Function

Function GenRndArr(ByVal n&, Optional ByVal m& = 0)
'������� ��������� ������� ���������� ��������� � ��������� ������� �� 1 �� n, ������������ m
    Dim i&, j&
    If n < 1 Then n = 1 '�������� ������������ �������� ������
    If m > n Or m < 1 Then m = n

    ReDim a&(1 To n)
    Randomize
    For i = 1 To n
        j = Int(Rnd * i + 1)
        If i <> j Then a(i) = a(j)
        a(j) = i
    Next i
    ReDim Preserve a&(1 To m)
    GenRndArr = a
End Function

Private Sub GenColFromStr(Arry() As Double, txt As String, n As Long, Optional lim As Long = 1)
'��������� ������������ ������� � ��������� �� �����
    Dim i As Long
    Dim j As Long
    Dim v

    v = SplitSchem(txt, n)
    ReDim Arry(0 To (UBound(v) - LBound(v) + 1) + 1)
    j = 1
    Arry(j) = lim
    For i = LBound(v) To UBound(v)
        j = j + 1
        Arry(j) = v(i)
    Next i
End Sub

Function SplitSchem(txt$, n&) '������� �������������� ����� � ������ ���������
    Dim i&, j&, k&, x
    ReDim out&(1 To n)
    For Each x In Split(txt, "+")
        i = InStr(x, "*")
        If i Then j = Val(Left$(x, i - 1)) Else j = 1
        k = Val(Replace(Replace(Mid$(x, i + 1), "[", ""), "]", ""))
        If k <= n Then out(k) = j
    Next x
    SplitSchem = out
End Function
