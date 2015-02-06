Public Class clsTest1
    Inherits Quasi97.clsQSTTestNET

    Public Const ThisTestID$ = "Head Voltage Test"
    Private objNativeRes As New NativeRes
    Private mvarAverages As Short = 1

    Private Class NativeRes
        Public bv$ = "Bias Voltage"
        Public res$ = "Resistance (Ohm)"
    End Class

    Public Property Average() As Short
        Get
            Return mvarAverages
        End Get
        Set(ByVal Value As Short)
            If Value < 1 Or Value > 100 Then
                Throw New System.ArgumentOutOfRangeException("Average", Value, "Out of range 1..100 [" & Value & "]")
            Else
                mvarAverages = Value
            End If
        End Set
    End Property

    Public Overrides Sub RunTest()
        Dim Rslt As Quasi97.ResultNet
        Dim MeasuredVal!
        Dim ib!

        Try

            QST.Normalization.InitTest(TestID, Setup)
            QST.Normalization.GetAdaptParam(TestID, Setup, "", 0)
            'measuring
            MeasuredVal = QST.QSTHardware.HReadRes(QST.QSTHardware.MRChannel, 10)
            ib = QST.QSTHardware.GetReadBias(QST.QSTHardware.MRChannel)

            'report results
            Rslt = New Quasi97.ResultNet
            Rslt.AddParameters(Me.colParameters)
            colResults.Insert(0, Rslt)
            Rslt.AddResult("Resistance (Ohm)", MeasuredVal.ToString("F2"), 1)
            Rslt.AddResult(objNativeRes.bv, (ib * MeasuredVal).ToString("F2"), 1)
            Rslt.CalcStats("RESULT")

            Call QST.Normalization.AddResultforRecordNET(TestID, Setup, "", Rslt) 'get all parameters, because we are not sweeping anything
            'grade results
            QST.GradingParameters.GradeTestNet(Me, 0)

            'add information about run
            Rslt.AddInfo(Me, 0, QST.QuasiParameters.CurInfo)

            'notify the form that new results are available
            MyBase.RaiseNewInfoAvailable()
            MyBase.RaiseNewResultsAvailable(New Integer() {0})

        Catch ex As Exception
            MsgBox("clsTest1:RunTest " & ex.Message)
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
        MyBase.colParameters.Add( _
            New Quasi97.clsTestParam("Average", "Average", Me, mvarAverages.GetType, True, True))

    End Sub
End Class
