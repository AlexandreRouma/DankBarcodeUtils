Imports System.Drawing

Namespace Utils
    Public Class Ean
        Public Shared ReadOnly EanAElements As String() = {"___XX_X", "__XX__X", "__X__XX", "_XXXX_X", "_X___XX", "_XX___X", "_X_XXXX", "_XXX_XX", "_XX_XXX", "___X_XX"}
        Public Shared ReadOnly EanBElements As String() = {"_X__XXX", "_XX__XX", "__XX_XX", "_X____X", "__XXX_X", "_XXX__X", "____X_X", "__X___X", "___X__X", "__X_XXX"}
        Public Shared ReadOnly EanCElements As String() = {"XXX__X_", "XX__XX_", "XX_XX__", "X____X_", "X_XXX__", "X__XXX_", "X_X____", "X___X__", "X__X___", "XXX_X__"}
        Public Shared ReadOnly Ean13FirstNum As String() = {"AAAAAA", "AABABB", "AABBAB", "AABBBA", "ABAABB", "ABBAAB", "ABBBAA", "ABABAB", "ABABBA", "ABBABA"}
        Public Shared Function GetChecksum(ByVal str As String, ByVal eanmask As String) As Integer
            Dim rtn As Integer
            Dim i As Integer = 0
            For Each car In str

                If eanmask(i) = "1" Then
                    rtn = rtn + (Integer.Parse(car) * 1)
                Else
                    rtn = rtn + (Integer.Parse(car) * 3)
                End If

                i = i + 1
            Next
            rtn = rtn Mod 10
            rtn = 10 - rtn
            If rtn = 10 Then
                rtn = 0
            End If
            Return rtn
        End Function
    End Class

    Public Class Code128
        Public Shared ReadOnly Code128Caracters As String() = {"11011001100", "11001101100", "11001100110", "10010011000", "10010001100", "10001001100", "10011001000", "10011000100", "10001100100", "11001001000", "11001000100", "11000100100", "10110011100", "10011011100", "10011001110", "10111001100", "10011101100", "10011100110", "11001110010", "11001011100", "11001001110", "11011100100", "11001110100", "11101101110", "11101001100", "11100101100", "11100100110", "11101100100", "11100110100", "11100110010", "11011011000", "11011000110", "11000110110", "10100011000", "10001011000", "10001000110", "10110001000", "10001101000", "10001100010", "11010001000", "11000101000", "11000100010", "10110111000", "10110001110", "10001101110", "10111011000", "10111000110", "10001110110", "11101110110", "11010001110", "11000101110", "11011101000", "11011100010", "11011101110", "11101011000", "11101000110", "11100010110", "11101101000", "11101100010", "11100011010", "11101111010", "11001000010", "11110001010", "10100110000", "10100001100", "10010110000", "10010000110", "10000101100", "10000100110", "10110010000", "10110000100", "10011010000", "10011000010", "10000110100", "10000110010", "11000010010", "11001010000", "11110111010", "11000010100", "10001111010", "10100111100", "10010111100", "10010011110", "10111100100", "10011110100", "10011110010", "11110100100", "11110010100", "11110010010", "11011011110", "11011110110", "11110110110", "10101111000", "10100011110", "10001011110", "10111101000", "10111100010", "11110101000", "11110100010", "10111011110", "10111101110", "11101011110", "11110101110"}
        Public Shared ReadOnly Code128StartA As String = "11010000100"
        Public Shared ReadOnly Code128StartB As String = "11010010000"
        Public Shared ReadOnly Code128StartC As String = "11010011100"
        Public Shared ReadOnly Code128Stop As String = "1100011101011"

        Public Shared Function GetChecksum(ByVal str As String, ByVal StartCode As Integer) As Integer
            Dim rtn As Long = 0
            For i As Integer = 0 To str.Length - 1
                rtn = rtn + ((Asc(str(i)) - 32) * (i + 1))
            Next
            rtn = rtn + StartCode
            Return rtn Mod 103
        End Function
    End Class

    Public Class Common
        Public Shared Function Resize(img As Bitmap, ByVal newWidth As Integer, ByVal newHeight As Integer) As Bitmap
            Dim out As New Bitmap(newWidth, newHeight)
            Dim gfx As Graphics = Graphics.FromImage(out)
            gfx.InterpolationMode = Drawing2D.InterpolationMode.NearestNeighbor
            gfx.PixelOffsetMode = Drawing2D.PixelOffsetMode.Half
            gfx.DrawImage(img, 0, 0, newWidth, newHeight)
            gfx.Save()
            Return out
        End Function

        Public Shared Function Invert(img As Bitmap) As Bitmap
            Dim out As Bitmap = img
            For y As Integer = 0 To img.Height - 1
                For x As Integer = 0 To img.Width - 1
                    out.SetPixel(x, y, Color.FromArgb(255 - img.GetPixel(x, y).R, 255 - img.GetPixel(x, y).G, 255 - img.GetPixel(x, y).B))
                Next
            Next
            Return out
        End Function

        Public Shared Sub DrawString(ByRef gfx As Graphics, ByVal str As String, ByVal font As Font, ByVal brush As Brush, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer)
            Dim tgfx As Graphics = Graphics.FromImage(New Bitmap(w, w * 3))
            Dim tlng As Integer = tgfx.MeasureString(str, font).Width
            Dim cl As Integer = tgfx.MeasureString("0", font).Width
            Dim nbcar As Integer = str.Length - 1
            Dim spacing As Integer = (w - tlng) / nbcar

            For c As Integer = 0 To nbcar + 1
                Try
                    gfx.DrawString(str(c), font, brush, x + (c * cl) + (c * spacing), y)
                Catch ex As Exception

                End Try
            Next
        End Sub

        Public Shared Function GetCaracterWidth(ByVal font As Font)
            Dim tgfx As Graphics = Graphics.FromImage(New Bitmap(font.Height, font.Height))
            Return tgfx.MeasureString("0", font).Width
        End Function
    End Class
End Namespace
