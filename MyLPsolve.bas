Attribute VB_Name = "MyLPsolve"
Option Explicit

Const MLDATA& = 90 'максимальная количество исходных данных
Const MLZGT& = 10 'максимальная количество заготовок
Const MLSTACK& = 2000000 'максимальный размер стека
Const MLGENVAR& = 200000 'максимальное количество сгенерированных вариантов
Const MAXCOLUMN As Long = 50000 'максимально возможное кол-во столбцов (карт раскроя) для lpsolve

Private bIsSolve As Boolean
Private bNoShowAlert As Boolean

Private lpsolve As lpsolve55

Private Function ctrlcfunc(ByVal lp As Long, ByVal userhandle As Long) As Long
    'If set to True, then solve is aborted and returncode will indicate this.
    'ctrlcfunc = True
End Function

Sub lpCSP() 'решение задачи CSP линейным программированием
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

    Set inpRng = [c26] 'исходные данные, отрезки
    Set inpRng2 = [c4] 'исходные данные, заготовки
    Set outRng = [k4] 'куда выводить решение

    shrez = Val([d15])
    krm = Val([d16])
    bGroupe = [e15] = True
    bGraph = [e16] = True
    delOst = Val([d17]) 'деловой остаток
    b2prof = [e17] = True 'признак раскроя в два профиля

    iSpeed = Val([h15]) 'скорость вычисления
    If iSpeed < 1 Then iSpeed = 1
    If iSpeed > 4 Then iSpeed = 5

    Call ClearSolve
    Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
    ActiveSheet.Unprotect

    '-----------------------------------------------------------------------------------------------------------------'

    For i = 1 To MLDATA 'ввод деталей
        m = Val(inpRng.Offset(i - 1, 0)) 'размер детали
        l = Val(inpRng.Offset(i - 1, 1)) 'кол-во деталей
        If m > 0 And l > 0 Then
            smDt = smDt + (m + shrez) * l 'вычисляем сумму всех деталей с учетом реза
            For j = 1 To nd
                If dtArr(1, j) = m Then Exit For
            Next j
            If j > nd Then 'если детали такого размера еще не было, то добавляем ее
                nd = nd + 1 'увеличиваем количество различных деталей
                ReDim Preserve dtArr&(1 To 2, 1 To nd) 'увеличиваем размер массива
            End If
            dtArr(1, j) = m 'запоминаем размер
            dtArr(2, j) = dtArr(2, j) + l 'запоминаем кол-во
        End If
    Next i

    '-----------------------------------------------------------------------------------------------------------------'

    If b2prof Then 'проверяем, четное ли кол-во деталей
        For i = 1 To nd
            If dtArr(2, i) Mod 2 Then b2prof = False: Exit For
        Next i
        If b2prof Then
            For i = 1 To nd
                dtArr(2, i) = dtArr(2, i) \ 2
            Next i
        Else
            [e17] = False
            MsgBox "Число деталей нечетное, раскрой будет произведен в один профиль"
        End If
    End If

    '-----------------------------------------------------------------------------------------------------------------'

    For i = 1 To MLZGT 'ввод заготовок
        m = Val(inpRng2.Offset(i - 1, 0)) 'размер заготовки
        'если кол-во заготовок не указано, то вычисляем нужное кол-во "жадным" алгоритмом
        If inpRng2.Offset(i - 1, 1) = "" And m > 0 Then l = fGreedyAlgo(dtArr, m, shrez, krm) Else l = Val(inpRng2.Offset(i - 1, 1)) \ IIf(b2prof, 2, 1) 'кол-во заготовок
        If m > 0 And l > 0 Then
            For j = 1 To nz
                If zgArr(1, j) = m Then Exit For
            Next j
            If j > nz Then 'если заготовки такого размера еще не было, то добавляем ее
                nz = nz + 1 'увеличиваем количество различных заготовок
                ReDim Preserve zgArr&(1 To 2, 1 To nz) 'увеличиваем размер массива
            End If
            zgArr(1, j) = m 'запоминаем размер
            zgArr(2, j) = zgArr(2, j) + l 'запоминаем кол-во
        End If
    Next i
    If nz = 0 Or nd = 0 Then Exit Sub 'данные отсутствуют

    '-----------------------------------------------------------------------------------------------------------------'

    ProgressLog.Show False
    ProgressLog.TxtClear
    ProgressLog.TxtAdd = "Решение CSP линейным программированием": DoEvents

    '-----------------------------------------------------------------------------------------------------------------'

    'вычисляем схемы рационального раскроя (оптимальных по Парето), с ограничением и без ограничения остатка
    For k = iSpeed + 2 To 1 Step -1
        nsh = fGenerateSum(dtArr, zgArr, shArr(), shrez, krm, Choose(k, 0, 20, 50, 100, 150, 200, 250))
        Debug.Print nsh
        If nsh > 0 Then 'если есть схемы раскроя, то произвести расчет
            For i = 1 To IIf(k > 0, 1, Choose(iSpeed, 2, 3, 4, 5, 6)) 'делаем несколько попыток решения
                iter = iter + 1
                ProgressLog.TxtAdd = "Итерация: " & iter & vbLf & "Схем раскроя: " & nsh: DoEvents
                ncsp = fCSP_LPsolve(dtArr, zgArr, shArr, lpArr, shrez, krm, Choose(iSpeed, 3, 5, 10, 15, 60)) 'запускаем LP решение
                If ncsp > 0 Then
                    ostMax = 0 'максимальный остаток
                    sOst = 0 'сумма полезных остатков
                    smOst = 0 'сумма остатков
                    For j = 1 To UBound(lpArr, 2)
                        ost = lpArr(3, j) - lpArr(2, j)
                        If ost > ostMax Then ostMax = ost 'определяем максимальный остаток
                        If ost >= delOst And delOst <> 0 Then sOst = sOst + ost 'вычисляем сумму полезных остатков
                        smOst = smOst + ost 'сумма остатков
                    Next j
                    If bstNCsp = 0 Or (ncsp < bstNCsp Or (ncsp = bstNCsp And bstDelOst > sOst) Or (ncsp = bstNCsp And bstDelOst = sOst And bstMxOst < ostMax)) Then
                        bstNCsp = ncsp 'запоминаем наилучшее решение
                        bstNCspKrm = smOst
                        bstDelOst = sOst
                        bstMxOst = ostMax
                        bstLpArr = lpArr
                    End If
                    ProgressLog.TxtAdd = "Длина заготовок: " & ncsp & vbLf & "Максимальный остаток: " & bstMxOst & vbLf: DoEvents
                End If
                If ProgressLog.Stoped Then GoTo EndSolution
            Next i
        End If
    Next k

    '-----------------------------------------------------------------------------------------------------------------'

    If bstNCsp > 0 And iSpeed > 2 Then 'максимизируем остаток
        ReDim zgNewArr&(1 To 2, 1 To nz) 'создаем новый массив заготовок
        For i = 1 To nz
            zgNewArr(1, i) = zgArr(1, i)
            If i = 1 Or maxDt > zgArr(1, i) - krm - 1 Then maxDt = zgArr(1, i) - krm - 1
            For j = 1 To UBound(bstLpArr, 2)
                If bstLpArr(3, j) = zgNewArr(1, i) Then zgNewArr(2, i) = zgNewArr(2, i) + 1
        Next j, i
        minDt = bstMxOst 'минимально возможный остаток
        If maxDt > bstNCspKrm Then maxDt = bstNCspKrm 'максимально возможный остаток

        dtNewArr = dtArr 'создаем новый массив с деталями
        ReDim Preserve dtNewArr&(1 To 2, 1 To nd + 1) 'добавляем элемент

        For i = 1 To 14 'повторять, пока не будет найдено решение
            dtNewArr(1, nd + 1) = (maxDt + minDt) \ 2 - shrez
            dtNewArr(2, nd + 1) = 1
            'создаем новые схемы раскроя
            nsh = fGenerateSum(dtNewArr, zgNewArr, shArr(), shrez, krm) ', bstNCspKrm - dtNewArr(1, nd + 1) - shrez)

            iter = iter + 1
            ProgressLog.TxtAdd = "Итерация: " & iter & vbLf & "Схем раскроя: " & nsh: DoEvents
            Debug.Print nsh, i, minDt, maxDt, bstNCspKrm - dtNewArr(1, nd + 1) - shrez

            If nsh > 0 Then
                ncsp = fCSP_LPsolve(dtNewArr, zgNewArr, shArr, lpArr, shrez, krm, 5, nd)
                If ncsp <= 0 Then ncsp = fCSP_LPsolve(dtNewArr, zgNewArr, shArr, lpArr, shrez, krm, 10, nd)
                If ncsp > 0 Then
                    ostMax = 0 'максимальный остаток
                    sOst = 0 'сумма полезных остатков
                    smOst = 0 'сумма остатков
                    For j = 1 To UBound(lpArr, 2)
                        ost = lpArr(3, j) - lpArr(2, j)
                        If ost > ostMax Then ostMax = ost 'определяем максимальный остаток
                        If ost >= delOst And delOst <> 0 Then sOst = sOst + ost 'вычисляем сумму полезных остатков
                        smOst = smOst + ost 'сумма остатков
                    Next j
                    bstNCsp = ncsp 'запоминаем решение
                    bstNCspKrm = smOst
                    bstDelOst = sOst
                    bstMxOst = ostMax
                    bstLpArr = lpArr
                    minDt = bstMxOst
                Else
                    maxDt = dtNewArr(1, nd + 1) + shrez
                End If
                ProgressLog.TxtAdd = "Длина заготовок: " & ncsp & vbLf & "Максимальный остаток: " & bstMxOst & vbLf: DoEvents
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

    If bstNCsp > 0 Then 'если решение было найдено
        out = fGenSolution(dtArr, bstLpArr, shrez, krm, bGroupe, b2prof) 'формируем решение для вывода на лист
        ns = UBound(out, 2)
        outRng.Resize(ns, 4) = Application.WorksheetFunction.Transpose(out) 'выводим его на лист
        outRng.Offset(0, 4).Resize(ns, 1).FormulaR1C1 = "=RC[-4]-RC[-3]"
        outRng.Resize(ns, 5).Borders.LineStyle = 1 'и раскрашиваем
        outRng.Resize(ns, 5).Borders(xlInsideHorizontal).Weight = xlHairline
        outRng.Resize(ns).Font.Bold = True
        outRng.Offset(0, 2).Resize(ns, 2).Font.Bold = True
        outRng.Offset(0, 2).Resize(ns).NumberFormat = "0"" шт."""
        outRng.Offset(0, 3).EntireColumn.AutoFit
        '-------------------------------------------------
        k = 0 'формируем массив остатков
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
        'сортируем остатки по убыванию
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
        If MsgBox("Не удалось нейти решение линейным программированием" & vbLf & _
                "Запустить решение динамическим программированием?", vbYesNo) = vbYes Then
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

    nd = UBound(dtArr, 2) 'кол-во видов деталей
    nz = UBound(zgArr, 2) 'кол-во видов заготовок
    ns = UBound(shArr, 2) 'кол-во рациональных схем
    If nnd = 0 Then nnd = nd 'если величина nnd меньше nd, то все элементы выше nnd будут отсеяны,
    'необходима при поиске максимального остатка

    If ns > MAXCOLUMN Then mx = MAXCOLUMN Else mx = ns 'ограничение используемых столбцов
    rndArr = GenRndArr(ns, mx) 'массив случайных элементов

    With lpsolve
        lp = .make_lp(nd, 0)
        .set_timeout lp, mTime 'ограничение выполнение расчетов в сек.
        .put_abortfunc lp, AddressOf ctrlcfunc, 0

        For i = 1 To nd  'формируем условия для строк
            .set_rh lp, i, dtArr(2, i): .set_constr_type lp, i, GE
        Next i
        For i = 1 To mx 'генерируем столбцы
            GenColFromStr aCol, CStr(shArr(1, rndArr(i))), nd, CLng(shArr(3, rndArr(i)))
            .add_column lp, aCol(1): .set_int lp, i, True
        Next i
        'добавляем условия ограничения заготовок
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
        SolveOk = .solve(lp) 'производим расчет
        If SolveOk = 0 Or SolveOk = 1 Then 'если решение найдено
            SolveObj = .get_objective(lp) 'оптимизированная функция
            ReDim aCol(1 To .get_Ncolumns(lp))
            .get_variables lp, aCol(1) 'получаем массив с решением

            'удаляем из схем с решением лишние детали
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
'функция, формирующая результат с решением
    Dim i&, j&, k&, nd&, ns&
    Dim v, txt$, t&

    nd = UBound(dtArr, 2) 'кол-во видов деталей
    ns = UBound(lpArr, 2) 'кол-во схем с решением

    ReDim out(1 To 4, 1 To ns)
    'сортируем решение
    For i = 1 To ns
        v = SplitSchem(CStr(lpArr(1, i)), nd)
        txt = ""
        For j = 1 To nd
            If v(j) > 0 Then txt = txt & " + (" & dtArr(1, j) & IIf(shrez, "+" & shrez, "") & IIf(v(j) > 1, " - " & v(j) & " шт.", "") & ")"
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

    If bGroupe Then 'если нужно группировать решение
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
'Функция генерации всех сочетаний слагаемых, для нахождения нужной суммы
'Автор MCH (Михаил Ч.), m-ch@mail.ru
'Использовано в задаче Cutting stock problem
'dt - массив деталей
'zg - массив заготовок
'out() - возвращаемый массив со схемами

    Dim i&, iz&, j&, k&, n&, m&, sl&
    Dim sm&, rws&
    Dim xd, xz

    n = UBound(dt, 2)

    ReDim out(1 To 3, 1 To MLGENVAR)
    For iz = 1 To UBound(zg, 2) 'перебираем все заготовки
        sm = zg(1, iz) 'искомая сумма
        If sm > 0 Then
            ReDim smi&(MLSTACK), smt$(MLSTACK), bNoRacio(MLSTACK) As Boolean 'массивы для стека
            sl = 0 'указатель конца стека
            Do 'запуск генерации всех вариантов суммирования
                For i = 1 To n
                    For j = 0 To sl
                        m = (sm - smi(j) - krm) \ (dt(1, i) + shrez)
                        If m > dt(2, i) Then m = dt(2, i)
                        For k = 1 To m
                            If k = 1 Then bNoRacio(j) = True Else bNoRacio(sl) = True 'схема j/sl не является "рациональной по Парето"
                            sl = sl + 1
                            If sl > MLSTACK Then sl = sl - 1: Debug.Print "Закончился стек": fGenerateSum = -2: Exit Do
                            smi(sl) = smi(j) + k * (dt(1, i) + shrez)
                            smt(sl) = smt(j) & "+" & IIf(k > 1, k & "*", "") & "[" & i & "]"
                        Next k
                        For k = i - 1 To 1 Step -1
                            If sm - smi(sl) - krm >= dt(1, k) + shrez Then bNoRacio(sl) = True: Exit For
                    Next k, j, i
            Loop While False

            'формируем массив для вывода на лист
            For i = 1 To sl
                If Not bNoRacio(i) Then
                    If mxKrm = 0 Or smi(i) >= sm - krm - mxKrm Then
                        rws = rws + 1
                        out(1, rws) = Mid$(smt(i), 2) 'схема
                        out(2, rws) = smi(i) + krm 'набранная сумма
                        out(3, rws) = sm 'размер заготовки
                        If rws >= MLGENVAR Then fGenerateSum = -1: Exit For
                    End If
                End If
            Next i
            If rws >= MLGENVAR Then fGenerateSum = -1: Exit For
        End If
    Next iz
    If rws Then ReDim Preserve out(1 To 3, 1 To rws) Else fGenerateSum = -1 'изменяем размер массива, если схем нет, то возвращаем -1
    Erase smi, smt, bNoRacio 'удаляем массивы
    If fGenerateSum = 0 Then fGenerateSum = rws
End Function

Function fGreedyAlgo&(dtArr, sm, Optional shrez = 0, Optional krm = 0)
'Функция вычисления необходимого кол-ва заготовок жадным алгоритмом
    Dim i&, j&, k&, n&, l&, t&, s&, aArr&(), bArr() As Boolean, bFlag As Boolean

    n = UBound(dtArr, 2)
    ReDim dt&(1 To n, 1 To 2)
    For i = 1 To n 'сортировка
        dt(i, 1) = dtArr(1, i) + shrez
        dt(i, 2) = dtArr(2, i)
        For j = 1 To i - 1
            If dt(i, 1) > dt(j, 1) Then
                t = dt(i, 1): dt(i, 1) = dt(j, 1): dt(j, 1) = t
                t = dt(i, 2): dt(i, 2) = dt(j, 2): dt(j, 2) = t
            End If
    Next j, i

    For i = 1 To n 'перекладываем в одномерный массив
        For j = 1 To dt(i, 2)
            k = k + 1
            ReDim Preserve aArr&(1 To k), bArr(1 To k) As Boolean
            aArr(k) = dt(i, 1)
    Next j, i

    Do 'определяем кол-во жадным алгоритмом
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
'функция генерации массива уникальных элементов в случайном порядке от 1 до n, размерностью m
    Dim i&, j&
    If n < 1 Then n = 1 'проверка корректности исходных данных
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
'процедура формирования столбца с условиями из схемы
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

Function SplitSchem(txt$, n&) 'функция преобразования схемы в массив количеств
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
