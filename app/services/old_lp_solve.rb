class OldLpSolve
  INITIAL_DATA_AMOUNT = 90 # MLDATA&
  BILLET_AMOUNT = 10 # MLZGT&
  STACK_SIZE = 2000000 # MLSTACK&
  GENERATED_VARIANTS_NUMBER = 200000 # MLGENVAR&
  CUTTING_CARDS_NUMBER = 50000 # MAXCOLUMN
  CALCULATION_SPEED = 1

  def lp_csp(input_line_segments, input_billets, output_segment, business_balance, cutting_thickness, krm, bGroupe, bGraph, b2prof, )
  end

  def method_name

  end

  def i_speed
    if CALCULATION_SPEED < 1
      1
    elsif CALCULATION_SPEED > 4
      5
    end
  end

  def clear_solve

  end

  def lpsolve
    @lpsolve ||= LPSolve.new
  end
end


    For i = 1 To MLDATA 'ввод деталей
        m = Val(input_line_segments.Offset(i - 1, 0)) 'размер детали
        l = Val(input_line_segments.Offset(i - 1, 1)) 'кол-во деталей
        If m > 0 And l > 0 Then
            useful_parts_length = useful_parts_length + (m + shrez) * l 'вычисляем сумму всех деталей с учетом реза
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

    For i = 1 To MLZGT 'ввод заготовок
        m = Val(input_billets.Offset(i - 1, 0)) 'размер заготовки
        'если кол-во заготовок не указано, то вычисляем нужное кол-во "жадным" алгоритмом
        If input_billets.Offset(i - 1, 1) = "" And m > 0 Then l = fGreedyAlgo(dtArr, m, shrez, krm) Else l = Val(input_billets.Offset(i - 1, 1)) \ IIf(b2prof, 2, 1) 'кол-во заготовок
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

    ProgressLog.Show False
    ProgressLog.TxtClear
    ProgressLog.TxtAdd = "Решение CSP линейным программированием": DoEvents

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
                        If ost >= business_balance And business_balance <> 0 Then sOst = sOst + ost 'вычисляем сумму полезных остатков
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
                        If ost >= business_balance And business_balance <> 0 Then sOst = sOst + ost 'вычисляем сумму полезных остатков
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

EndSolution:
    Application.EnableCancelKey = xlInterrupt
    ProgressLog.Hide

    If bstNCsp > 0 Then 'если решение было найдено
        out = fGenSolution(dtArr, bstLpArr, shrez, krm, bGroupe, b2prof) 'формируем решение для вывода на лист
        ns = UBound(out, 2)
        output_segment.Resize(ns, 4) = Application.WorksheetFunction.Transpose(out) 'выводим его на лист
        output_segment.Offset(0, 4).Resize(ns, 1).FormulaR1C1 = "=RC[-4]-RC[-3]"
        output_segment.Resize(ns, 5).Borders.LineStyle = 1 'и раскрашиваем
        output_segment.Resize(ns, 5).Borders(xlInsideHorizontal).Weight = xlHairline
        output_segment.Resize(ns).Font.Bold = True
        output_segment.Offset(0, 2).Resize(ns, 2).Font.Bold = True
        output_segment.Offset(0, 2).Resize(ns).NumberFormat = "0"" шт."""
        output_segment.Offset(0, 3).EntireColumn.AutoFit
        '-------------------------------------------------
        k = 0 'формируем массив остатков
        ReDim dOst(1 To 100, 1 To 2)
        For i = 1 To MLZGT
            If input_billets.Offset(i - 1, 0) <> "" Then
                If input_billets.Offset(i - 1, 1) = "" Or input_billets.Offset(i - 1, 3) > 0 Then
                    For j = 1 To k
                        If dOst(j, 1) = input_billets.Offset(i - 1, 1) Then Exit For
                    Next j
                    If j > 100 Then Exit For
                    If k < j Then k = j: dOst(j, 1) = input_billets.Offset(i - 1, 0)
                    If input_billets.Offset(i - 1, 1) = "" Then dOst(j, 2) = Empty Else dOst(j, 2) = dOst(j, 2) + input_billets.Offset(i - 1, 3)
                End If
            End If
        Next i
        If business_balance > 0 Then
            For i = 1 To ns
                If out(1, i) - out(2, i) >= business_balance Then
                    For j = 1 To k
                        If dOst(j, 1) = output_segment.Offset(i - 1, 4) Then Exit For
                    Next j
                    If j > 100 Then Exit For
                    If k < j Then k = j: dOst(j, 1) = output_segment.Offset(i - 1, 4)
                    dOst(j, 2) = dOst(j, 2) + output_segment.Offset(i - 1, 2)
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
        input_billets.Offset(0, 5).Resize(MLZGT, 2) = dOst
        '-------------------------------------------------
    Else
        If MsgBox("Не удалось нейти решение линейным программированием" & vbLf & _
                "Запустить решение динамическим программированием?", vbYesNo) = vbYes Then
            Raskroy
            Exit Sub
        End If
    End If

    Debug.Print Timer - iTimer

    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    If bGraph Then Call OutGraph
    'Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
