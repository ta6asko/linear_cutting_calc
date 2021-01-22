Attribute VB_Name = "MyRibbonMenu"
'Demo-версия, реализация от 03.03.2016

Sub MyRibbonRaskroyDP(control As IRibbonControl)
    Raskroy
End Sub

Sub MyRibbonRaskroyLP(control As IRibbonControl)
    lpCSP
End Sub

Sub MyRibbonNew(control As IRibbonControl)
    ClearData
End Sub

Sub MyRibbonClearSolve(control As IRibbonControl)
    ClearSolve
End Sub

Sub MyRibbonGraph(control As IRibbonControl)
    OutGraph
End Sub

Sub MyRibbonOffGraph(control As IRibbonControl)
    OffGraph
End Sub

Sub MyRibbonExport2PDF(control As IRibbonControl)
    Export2PDF
End Sub

Sub MyRibbonCreateReport(control As IRibbonControl)

End Sub

Sub MyRibbonAbout(control As IRibbonControl)
    frmAbout.Show
End Sub

