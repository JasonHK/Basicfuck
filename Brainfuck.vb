Module Brainfuck

    Friend Cell(29999) As SByte
    Friend Delay As Integer = 50
    Friend IsRunning As Boolean = False
    Friend ForceStop As Boolean = False

    Friend Sub Initialization()

        For index = 0 To 29999
            Cell(index) = 0
        Next

    End Sub

    Friend Sub Interpreter(ByVal Code As String)

        Dim CodeLength As Integer = Code.Length
        Dim CellPointer As UInteger = 0
        Dim LoopDepth As Integer = 0
        Dim LoopJump(1000) As Integer
        Dim InitialComment As Boolean = False
        Dim SkipDelay As Boolean = False

        For index = 0 To CodeLength - 1
            If Not ForceStop Then
                Select Case Code.Chars(index)
                    Case ">"
                        If Not InitialComment Then
                            Try
                                CellPointer += 1
                            Catch ex As OverflowException
                                CellPointer = 29999
                            End Try
                        End If
                    Case "<"
                        If Not InitialComment Then
                            Try
                                CellPointer -= 1
                            Catch ex As OverflowException
                                CellPointer = 0
                            End Try
                        End If
                    Case "+"
                        If Not InitialComment Then
                            Try
                                Cell(CellPointer) += 1
                            Catch ex As OverflowException
                                Cell(CellPointer) = -128
                            End Try
                        End If
                    Case "-"
                        If Not InitialComment Then
                            Try
                                Cell(CellPointer) -= 1
                            Catch ex As OverflowException
                                Cell(CellPointer) = 127
                            End Try
                        End If
                    Case "."
                        If Not InitialComment Then MainForm.BrainfuckOutput(Cell(CellPointer))
                    Case ","
                        If Not InitialComment Then Cell(CellPointer) = MainForm.BrainfuckInput()
                    Case "["
                        If index = 0 Then InitialComment = True
                        LoopDepth += 1
                        LoopJump(LoopDepth) = index
                    Case "]"
                        If Cell(CellPointer) = 0 Then
                            LoopDepth -= 1
                        Else
                            index = LoopJump(LoopDepth)
                        End If
                        If LoopDepth = 0 Then InitialComment = False
                    Case Else
                        SkipDelay = True
                End Select

                If Not SkipDelay And Not InitialComment Then
                    MainForm.BrainfuckPointer(CellPointer, Cell(CellPointer))
                    Debug.WriteLine("Cell: " & CStr(CellPointer))
                    Debug.WriteLine("Value: " & CStr(Cell(CellPointer)))
                    Sleep(Delay)
                End If
                SkipDelay = False
            Else
                ForceStop = False
                Exit Sub
            End If
        Next

    End Sub

    Private Sub Sleep(ByVal Delay As Integer)

        For index = 1 To Delay
            Application.DoEvents() : Threading.Thread.Sleep(1) : Application.DoEvents()
        Next

    End Sub
End Module
