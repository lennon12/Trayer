Class MainWindow
    Private colCDROMs As Object

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        Dim oWMP = CreateObject("WMPlayer.OCX.7")

        colCDROMs = oWMP.cdromCollection

        If colCDROMs.Count >= 1 Then

            For i = 0 To colCDROMs.Count - 1

                colCDROMs.Item(i).Eject

            Next ' cdrom

        End If

    End Sub

    Private Function oWMP() As Object
        Throw New NotImplementedException()
    End Function

    Private Sub Button_Click_1(sender As Object, e As RoutedEventArgs)
        For d = 0 To colCDROMs.Count - 1

            colCDROMs.Item(d).Eject
        Next 'null
    End Sub
End Class
