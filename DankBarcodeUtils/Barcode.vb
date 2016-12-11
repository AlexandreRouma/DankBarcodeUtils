Imports System.Drawing
Imports System.Text.RegularExpressions

Namespace Barcode
    Public Class Ean8

        Public ReadOnly Data As String
        Public ReadOnly Text As Boolean = False
        Public ReadOnly Flat As Boolean = True
        Public ReadOnly Checksum As Integer

        Public Sub New(ByVal _data As String, ByVal _flat As Boolean, ByVal _text As String)
            Text = _text
            Flat = _flat
            _data = Regex.Replace(_data, "[^0-9]", "")
            Data = _data.PadRight(7, "0")
            Checksum = Utils.Ean.GetChecksum(Data, "31313131")
        End Sub

        Public Sub New(ByVal _data As String, ByVal _flat As Boolean)
            Flat = _flat
            _data = Regex.Replace(_data, "[^0-9]", "")
            Data = _data.PadRight(7, "0")
            Checksum = Utils.Ean.GetChecksum(Data, "31313131")
        End Sub

        Public Sub New(ByVal _data As String, ByVal _text As String)
            Text = _text
            _data = Regex.Replace(_data, "[^0-9]", "")
            Data = _data.PadRight(7, "0")
            Checksum = Utils.Ean.GetChecksum(Data, "31313131")
        End Sub

        Public Sub New(ByVal _data As String)
            _data = Regex.Replace(_data, "[^0-9]", "")
            Data = _data.PadRight(7, "0")
            Checksum = Utils.Ean.GetChecksum(Data, "31313131")
        End Sub

        Public Function GetImage(ByVal size As Size, ByVal inverted As Boolean) As Image
            Dim codeline As String = getCodeLine(Data)
            Dim tmp_bmp As Bitmap = New Bitmap(67, Size.Height)
            Dim cheight As Integer = Size.Height
            If Flat = False Or Text = True Then
                cheight = size.Height - (size.Height / 7)
            End If
            For y As Integer = 0 To cheight - 1
                For x As Integer = 0 To 66
                    If codeline(x) = "1" Then
                        tmp_bmp.SetPixel(x, y, Color.Black)
                    Else
                        tmp_bmp.SetPixel(x, y, Color.White)
                    End If
                Next
            Next
            If Flat = False Then
                For y As Integer = cheight To Size.Height - 1
                    tmp_bmp.SetPixel(0, y, Color.Black)
                    tmp_bmp.SetPixel(1, y, Color.White)
                    tmp_bmp.SetPixel(2, y, Color.Black)
                    For x As Integer = 3 To 30
                        tmp_bmp.SetPixel(x, y, Color.White)
                    Next
                    tmp_bmp.SetPixel(31, y, Color.White)
                    tmp_bmp.SetPixel(32, y, Color.Black)
                    tmp_bmp.SetPixel(33, y, Color.White)
                    tmp_bmp.SetPixel(34, y, Color.Black)
                    tmp_bmp.SetPixel(35, y, Color.White)
                    For x As Integer = 36 To 63
                        tmp_bmp.SetPixel(x, y, Color.White)
                    Next
                    tmp_bmp.SetPixel(64, y, Color.Black)
                    tmp_bmp.SetPixel(65, y, Color.White)
                    tmp_bmp.SetPixel(66, y, Color.Black)
                Next
            ElseIf Text = True Then
                For y As Integer = cheight To size.Height - 1
                    For x As Integer = 0 To 66
                        tmp_bmp.SetPixel(x, y, Color.White)
                    Next
                Next
            End If
            Dim image_src As Image = Utils.Common.Resize(tmp_bmp, size.Width, size.Height)
            Dim image As New Bitmap(size.Width, size.Height)
            Dim gfx As Graphics = Graphics.FromImage(image)
            gfx.DrawImage(image_src, 0, 0)
            If Text = True Then
                Utils.Common.drawString(gfx, Data.Substring(0, 4), New Font("Arial", size.Height / 9), Brushes.Black, (size.Width / 7.5), cheight, (size.Width / 2) - (size.Width / 3.3))
                Utils.Common.drawString(gfx, Data.Substring(4, 3) & Checksum, New Font("Arial", size.Height / 9), Brushes.Black, (size.Width / 1.65), cheight, (size.Width / 2) - (size.Width / 3.3))
            End If
            If inverted Then
                image = Utils.Common.Invert(image)
            End If
            Return image
        End Function

        Public Function GetImage(ByVal width As Integer, ByVal height As Integer, ByVal inverted As Boolean) As Image
            Return GetImage(New Size(width, height), inverted)
        End Function

        Public Function GetBitmap(ByVal size As Size, ByVal inverted As Boolean) As Bitmap
            Return GetImage(size, inverted)
        End Function

        Public Function GetBitmap(ByVal width As Integer, ByVal height As Integer, ByVal inverted As Boolean) As Bitmap
            Return GetImage(New Size(width, height), inverted)
        End Function

        Public Sub GetFile(ByVal size As Size, ByVal inverted As Boolean, ByVal filename As String)
            Try
                GetImage(size, inverted).Save(filename)
            Catch ex As Exception

            End Try
        End Sub

        Public Sub GetFile(ByVal width As Integer, ByVal height As Integer, ByVal inverted As Boolean, ByVal filename As String)
            Try
                GetImage(New Size(width, height), inverted).Save(filename)
            Catch ex As Exception

            End Try
        End Sub

        Private Function getCodeLine(ByVal _data As String) As String
            Dim rtn As String = "101"
            _data = _data.Substring(0, 7)
            _data = _data & Checksum

            For i As Integer = 0 To 3 Step 1
                For b As Integer = 0 To 6 Step 1
                    If Utils.Ean.EanAElements(Integer.Parse(_data.ElementAt(i))).ElementAt(b) = "X" Then
                        rtn = rtn & "1"
                    Else
                        rtn = rtn & "0"
                    End If
                Next
            Next

            rtn = rtn & "01010"

            For i As Integer = 0 To 3 Step 1
                For b As Integer = 0 To 6 Step 1
                    If Utils.Ean.EanCElements(Integer.Parse(_data.ElementAt(i + 4))).ElementAt(b) = "X" Then
                        rtn = rtn & "1"
                    Else
                        rtn = rtn & "0"
                    End If
                Next
            Next

            rtn = rtn & "101"
            Return rtn
        End Function

    End Class

    Public Class Ean13
        Public ReadOnly Data As String
        Public ReadOnly Text As Boolean = False
        Public ReadOnly Flat As Boolean = True
        Public ReadOnly Checksum As Integer

        Public Sub New(ByVal _data As String, ByVal _flat As Boolean, ByVal _text As Boolean)
            Text = _text
            Flat = _flat
            _data = Regex.Replace(_data, "[^0-9]", "")
            Data = _data.PadRight(12, "0")
            Checksum = Utils.Ean.GetChecksum(Data, "131313131313")
        End Sub

        Public Sub New(ByVal _data As String, ByVal _flat As Boolean)
            Flat = _flat
            _data = Regex.Replace(_data, "[^0-9]", "")
            Data = _data.PadRight(12, "0")
            Checksum = Utils.Ean.GetChecksum(Data, "131313131313")
        End Sub

        Public Sub New(ByVal _data As String)
            _data = Regex.Replace(_data, "[^0-9]", "")
            Data = _data.PadRight(12, "0")
            Checksum = Utils.Ean.GetChecksum(Data, "131313131313")
        End Sub

        Public Function GetImage(ByVal size As Size, ByVal inverted As Boolean) As Image
            Dim codeline As String = getCodeLine(Data)
            Dim tmp_bmp As Bitmap = New Bitmap(95, size.Height)
            Dim cheight As Integer = size.Height
            If Flat = False Or Text = True Then
                cheight = size.Height - (size.Height / 7)
            End If
            For y As Integer = 0 To cheight - 1
                For x As Integer = 0 To 94
                    If codeline(x) = "1" Then
                        tmp_bmp.SetPixel(x, y, Color.Black)
                    Else
                        tmp_bmp.SetPixel(x, y, Color.White)
                    End If
                Next
            Next
            If Flat = False Then
                For y As Integer = cheight To size.Height - 1
                    tmp_bmp.SetPixel(0, y, Color.Black)
                    tmp_bmp.SetPixel(1, y, Color.White)
                    tmp_bmp.SetPixel(2, y, Color.Black)
                    For x As Integer = 3 To 44
                        tmp_bmp.SetPixel(x, y, Color.White)
                    Next
                    tmp_bmp.SetPixel(45, y, Color.White)
                    tmp_bmp.SetPixel(46, y, Color.Black)
                    tmp_bmp.SetPixel(47, y, Color.White)
                    tmp_bmp.SetPixel(48, y, Color.Black)
                    tmp_bmp.SetPixel(49, y, Color.White)
                    For x As Integer = 50 To 91
                        tmp_bmp.SetPixel(x, y, Color.White)
                    Next
                    tmp_bmp.SetPixel(92, y, Color.Black)
                    tmp_bmp.SetPixel(93, y, Color.White)
                    tmp_bmp.SetPixel(94, y, Color.Black)
                Next
            ElseIf Text = True Then
                For y As Integer = cheight To size.Height - 1
                    For x As Integer = 0 To 94
                        tmp_bmp.SetPixel(x, y, Color.White)
                    Next
                Next
            End If
            Dim font As New Font("Arial", size.Height / 9)
            Dim cw As Integer = Utils.Common.GetCaracterWidth(font)
            Dim image_src As Image
            If Text = True Then
                image_src = Utils.Common.Resize(tmp_bmp, size.Width - (cw + 1), size.Height)
            Else
                image_src = Utils.Common.Resize(tmp_bmp, size.Width, size.Height)
            End If
            Dim image As New Bitmap(size.Width, size.Height)
            Dim gfx As Graphics = Graphics.FromImage(image)
            If Text = True Then
                gfx.DrawImage(image_src, cw + 1, 0)
                gfx.FillRectangle(Brushes.White, 0, 0, cw + 1, size.Height)
                gfx.DrawString(Data(0), font, Brushes.Black, 0, cheight)
                Utils.Common.DrawString(gfx, Data.Substring(1, 6), font, Brushes.Black, (size.Width / 11) + cw + 1, cheight, (size.Width / 2) - (size.Width / 3.3))
                Utils.Common.DrawString(gfx, Data.Substring(6, 5) & Checksum, font, Brushes.Black, (size.Width / 1.85) + cw + 1, cheight, (size.Width / 2) - (size.Width / 3.3))
            Else
                gfx.DrawImage(image_src, 0, 0)
            End If
            If inverted Then
                image = Utils.Common.Invert(image)
            End If
            Return image
        End Function

        Public Function GetImage(ByVal width As Integer, ByVal height As Integer, ByVal inverted As Boolean) As Image
            Return GetImage(New Size(width, height), inverted)
        End Function

        Public Function GetBitmap(ByVal size As Size, ByVal inverted As Boolean) As Bitmap
            Return GetImage(size, inverted)
        End Function

        Public Function GetBitmap(ByVal width As Integer, ByVal height As Integer, ByVal inverted As Boolean) As Bitmap
            Return GetImage(New Size(width, height), inverted)
        End Function

        Public Sub GetFile(ByVal size As Size, ByVal inverted As Boolean, ByVal filename As String)
            Try
                GetImage(size, inverted).Save(filename)
            Catch ex As Exception

            End Try
        End Sub

        Public Sub GetFile(ByVal width As Integer, ByVal height As Integer, ByVal inverted As Boolean, ByVal filename As String)
            Try
                GetImage(New Size(width, height), inverted).Save(filename)
            Catch ex As Exception

            End Try
        End Sub

        Public Function getCodeLine(ByVal _data As String) As String
            Dim rtn As String = "101"
            Dim ET As Char = ""
            _data = _data.Substring(0, 12)
            _data = _data & Checksum

            For i As Integer = 0 To 5 Step 1
                ET = Utils.Ean.Ean13FirstNum(Integer.Parse(_data(0))).ElementAt(i)
                For b As Integer = 0 To 6 Step 1
                    If ET = "A" Then
                        If Utils.Ean.EanAElements(Integer.Parse(_data.ElementAt(i + 1))).ElementAt(b) = "X" Then
                            rtn = rtn & "1"
                        Else
                            rtn = rtn & "0"
                        End If
                    Else
                        If Utils.Ean.EanBElements(Integer.Parse(_data.ElementAt(i + 1))).ElementAt(b) = "X" Then
                            rtn = rtn & "1"
                        Else
                            rtn = rtn & "0"
                        End If
                    End If
                Next
            Next

            rtn = rtn & "01010"

            For i As Integer = 0 To 5 Step 1
                For b As Integer = 0 To 6 Step 1
                    If Utils.Ean.EanCElements(Integer.Parse(_data.ElementAt(i + 7))).ElementAt(b) = "X" Then
                        rtn = rtn & "1"
                    Else
                        rtn = rtn & "0"
                    End If
                Next
            Next

            rtn = rtn & "101"
            Return rtn
        End Function
    End Class

    Public Class Code128
        Public ReadOnly Data As String

        Public Sub New(ByVal _data As String)
            Data = _data
        End Sub

        Public Function GetImage(ByVal size As Size, ByVal inverted As Boolean)
            Dim codeline As String = GetCodeLine(Data)
            Dim tmp_bmp As New Bitmap(codeline.Length, 1)
            For x As Integer = 0 To codeline.Length - 1
                If codeline(x) = "1" Then
                    tmp_bmp.SetPixel(x, 0, Color.Black)
                Else
                    tmp_bmp.SetPixel(x, 0, Color.White)
                End If
            Next
            tmp_bmp = Utils.Common.Resize(tmp_bmp, size.Width, size.Height)
            If inverted Then
                tmp_bmp = Utils.Common.Invert(tmp_bmp)
            End If
            Return tmp_bmp
        End Function

        Public Function GetImage(ByVal width As Integer, ByVal height As Integer, ByVal inverted As Boolean) As Image
            Return GetImage(New Size(width, height), inverted)
        End Function

        Public Function GetBitmap(ByVal size As Size, ByVal inverted As Boolean) As Bitmap
            Return GetImage(size, inverted)
        End Function

        Public Function GetBitmap(ByVal width As Integer, ByVal height As Integer, ByVal inverted As Boolean) As Bitmap
            Return GetImage(New Size(width, height), inverted)
        End Function

        Public Sub GetFile(ByVal size As Size, ByVal inverted As Boolean, ByVal filename As String)
            Try
                GetImage(size, inverted).Save(filename)
            Catch ex As Exception

            End Try
        End Sub

        Public Sub GetFile(ByVal width As Integer, ByVal height As Integer, ByVal inverted As Boolean, ByVal filename As String)
            Try
                GetImage(New Size(width, height), inverted).Save(filename)
            Catch ex As Exception

            End Try
        End Sub

        Public Function GetCodeLine(ByVal _data As String) As String
            Dim rtn As String = "11010010000"
            For i As Integer = 0 To _data.Length - 1
                rtn = rtn & Utils.Code128.Code128Caracters(Asc(_data(i)) - 32)
            Next
            rtn = rtn & Utils.Code128.Code128Caracters(Utils.Code128.GetChecksum(_data, 104))
            rtn = rtn & "1100011101011"
            Return rtn
        End Function
    End Class
End Namespace
