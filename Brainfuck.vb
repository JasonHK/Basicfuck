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
        Dim CommentLoop As Boolean = False
        Dim CommentDepth As Integer = 0
        Dim SkipDelay As Boolean = False

        For index = 0 To CodeLength - 1
            If Not ForceStop Then
                Select Case Code.Chars(index)
                    Case ">"
                        If Not CommentLoop Then
                            Try
                                CellPointer += 1
                            Catch ex As OverflowException
                                CellPointer = 29999
                            End Try
                        End If
                    Case "<"
                        If Not CommentLoop Then
                            Try
                                CellPointer -= 1
                            Catch ex As OverflowException
                                CellPointer = 0
                            End Try
                        End If
                    Case "+"
                        If Not CommentLoop Then
                            Try
                                Cell(CellPointer) += 1
                            Catch ex As OverflowException
                                Cell(CellPointer) = -128
                            End Try
                        End If
                    Case "-"
                        If Not CommentLoop Then
                            Try
                                Cell(CellPointer) -= 1
                            Catch ex As OverflowException
                                Cell(CellPointer) = 127
                            End Try
                        End If
                    Case "."
                        If Not CommentLoop Then MainForm.BrainfuckOutput(Cell(CellPointer))
                    Case ","
                        If Not CommentLoop Then Cell(CellPointer) = MainForm.BrainfuckInput()
                    Case "["
                        If Cell(CellPointer) = 0 And Not CommentLoop Then
                            CommentLoop = True
                            CommentDepth += 1
                        ElseIf CommentLoop Then
                            CommentDepth += 1
                        Else
                            LoopDepth += 1
                            LoopJump(LoopDepth) = index
                        End If
                    Case "]"
                        If Cell(CellPointer) = 0 And Not CommentLoop Then
                            LoopDepth -= 1
                        ElseIf Not CommentLoop Then
                            index = LoopJump(LoopDepth)
                        Else
                            CommentDepth -= 1
                            If CommentDepth = 0 Then CommentLoop = False
                        End If
                    Case Else
                        SkipDelay = True
                End Select

                If Not SkipDelay And Not CommentLoop Then
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
