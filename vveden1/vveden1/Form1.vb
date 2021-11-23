Public Class Form1
    Const k As Integer = 4
    Const n As Integer = 100
    Private a(n) As Double
    Const L As Long = 10 ^ (2 * k)
    Private ch As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim s As Integer
        Dim vpr(9, 9) As Integer
        Dim vch(9) As Integer
        Const xin = 4.168
        Const xiv = 14.683
        Const xin_pr = 81.449
        Const xiv_pr = 117.407

        While s < ch.Length
            For i = 0 To 9
                For j = 0 To 9
                    If ch(s) = Format(i) And ch(s + 1) = Format(j) Then
                        vpr(i, j) += 1
                    End If
                Next
            Next
            s += 2
        End While

        For i = 0 To 9
            For j = 0 To 9
                vch(i) += vpr(i, j) + vpr(j, i)
            Next
        Next

        Dim xi_pr, xi_ch As Double
        Const рch As Double = 0.1
        Const рpr As Double = 0.01

        For i = 0 To 9
            xi_ch += (vch(i) - 2 * k * n * рch) ^ 2
        Next
        xi_ch /= (рch * 2 * k * n)

        For i = 0 To 9
            For j = 0 To 9
                xi_pr += (vpr(i, j) - k * n * рpr) ^ 2
            Next
        Next
        xi_pr /= (рpr * k * n)

        MsgBox("Критерий согласия частот = " + Format(xi_ch) + Chr(13) + Chr(10) +
               "Критерий согласия пар = " + Format(xi_pr) + Chr(13) + Chr(10) +
               "Доверительный интервал для критерия согласия частот: [" + Format(xin) + "; " + Format(xiv) + "]" + Chr(13) + Chr(10) +
               "Доверительный интервал для критерия согласия пар: [" + Format(xin_pr) + "; " + Format(xiv_pr) + "]", , "Проверка частот и пар")
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim i As Integer
        Dim d1, d As Double
        a(0) = 0.12345678

        For i = 0 To n - 1
            d = Int(a(i) ^ 2 * 10 ^ (3 * k))
            d1 = d * 10 ^ (-2 * k)
            a(i + 1) = d1 - Int(d1)

            ListBox1.Items.Add(Format(a(i), "0.00000000"))
            ch += Format(a(i) * L, "00000000")
        Next i

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim S1, S2, S3, p1, p2, p, pmax As Double
        Const zp = 1.65

        For i = 0 To n - 1

            S1 += a(i) * i
            S2 += a(i)
            S3 += a(i) ^ 2

        Next i

        p1 = (S1 - S2 * (n + 1) / 2) / n
        p2 = ((S3 / n - (S2 / n) ^ 2) * (n ^ 2 - 1) / 12) ^ (1 / 2)
        p = p1 / p2
        pmax = zp * (1 - p ^ 2) / n ^ (1 / 2)

        MsgBox("Коэффициент корреляции = " + Format(p, "0.0000") + Chr(13) + Chr(10) +
               "Верхняя граница доверительного интервала = " + Format(pmax, "0.0000"), , "Проверка независимости")
    End Sub
End Class
