Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim pdf As New FPDFnet
        pdf.AddPage()
        pdf.SetFont("Arial", "B", 16)
        pdf.Cell(40, 10, "¡Hola, Mundo!")
        pdf.Output("tutorial1.pdf")
        System.Diagnostics.Process.Start _
       (Application.StartupPath + "\tutorial1.pdf")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim fpdf As New FPDFnet


        AddHandler fpdf.header, Sub(ByRef pdf As FPDFnet)
                                    pdf.image("logo_pb.png", 10, 8, 33)
                                    pdf.SetFont("Arial", "B", 15)
                                    pdf.Cell(80)
                                    pdf.Cell(30, 10, "Title", 1, 0, "C")
                                    pdf.Ln(20)
                                End Sub

        AddHandler fpdf.footer, Sub(ByRef pdf As FPDFnet)
                                    pdf.SetY(-15)
                                    pdf.SetFont("Arial", "I", 8)
                                    pdf.Cell(0, 10, "Page " + CStr(pdf.PageNo()) + "/{nb}", 0, 0, "C")
                                End Sub



        fpdf.AliasNbPage()
        fpdf.AddPage()
        fpdf.SetFont("Times", "", 12)
        For i As Integer = 1 To 40
            fpdf.Cell(0, 10, "Imprimiendo línea número " + CStr(i), 0, 1)
        Next
        fpdf.Output("tutorial2.pdf")
        System.Diagnostics.Process.Start _
       (Application.StartupPath + "\tutorial2.pdf")
    End Sub
    End Class
