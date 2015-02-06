Public Class clsTest2
    Inherits Quasi97.clsQSTTestNET

    Public Const ThisTestID$ = "NETSample 5"
    Private objNativeRes As New NativeRes
    Private mvarAverages As Short = 1

    Private Class NativeRes
        Public bv$ = "Head Voltage (mV)"
        Public res$ = "Resistance (Ohm)"
    End Class

    Public Overrides Sub RunTest()
        Dim Rslt As Quasi97.ResultNet
        Dim MeasuredVal!
        Dim ib!
        Dim ts As Quasi97.TestShell

        Try
            'measuring
            MeasuredVal = QST.QSTHardware.HReadRes(QST.QSTHardware.MRChannel, 10)
            ib = QST.QSTHardware.GetReadBias(QST.QSTHardware.MRChannel)

            'report results
            Rslt = New Quasi97.ResultNet
            colResults.Insert(0, Rslt)

            'example of scanning test collection
            Dim tKey$ = ""
            For Each t In QST.QuasiParameters.TestListLegacy
                If t.Value.TestID = "Transverse" Then tkey = t.Key
            Next
            If tKey <> "" Then ts = QST.QuasiParameters.TestListLegacy(tKey)

            'getting result from another test
            ts = QST.QuasiParameters.TestObj("Transverse", 1)
            If ts.TestPtr.colResults.Count = 0 Then ts.RunTest()

            Dim tobj As Quasi97.clsQSTTestNET = ts.TestPtr
            MeasuredVal = tobj.colResults(0).GetResult("Resistance (Ohm)")
            ib = tobj.colResults(0).GetResult("Bias Current (mA)")


            'adding results
            Rslt.AddResult(objNativeRes.res, MeasuredVal.ToString("F2"), 1)
            Rslt.AddResult(objNativeRes.bv, (ib * MeasuredVal).ToString("F2"), 1)
            Rslt.CalcStats("RESULT")

            'grade results
            QST.GradingParameters.GradeTestNet(Me, 0)

            'add information about run
            Rslt.AddInfo(Me, 0, QST.QuasiParameters.CurInfo)

            'notify the form that new results are available
            MyBase.RaiseNewInfoAvailable()
            MyBase.RaiseNewResultsAvailable(New Integer() {0})

        Catch ex As Exception
            MsgBox("clsTest2:RunTest " & ex.Message)
        End Try
    End Sub

    Public Overrides ReadOnly Property TestID As String
        Get
            Return ThisTestID
        End Get
    End Property

    Public Overrides Function CheckRecords(NewDBase As String) As System.Collections.Generic.List(Of Short)
        Return New System.Collections.Generic.List(Of Short)
    End Function

    Public Overrides Sub ClearResults(Optional doRefreshPlot As Boolean = False)
        colResults.Clear()
        MyBase.RaiseResultsCleared(doRefreshPlot)
    End Sub

    Public Overrides Sub RemoveRecord()

    End Sub

    Public Overrides Sub RestoreParameters()

    End Sub

    Public Overrides Sub SetDBase(ByRef NewDBase As String, Optional ByRef voidParam As Object = Nothing)

    End Sub

    Public Overrides Sub StoreParameters()

    End Sub

    Public Overrides ReadOnly Property ContainsGraph As Short
        Get
            Return False
        End Get
    End Property

    Public Overrides ReadOnly Property ContainsResultPerCycle As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides ReadOnly Property DualChannelCapable As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides ReadOnly Property FeatureVector As UInteger
        Get
            Return 0
        End Get
    End Property

    Public Sub New()
        MyBase.RegisterResults(objNativeRes)
    End Sub
End Class
