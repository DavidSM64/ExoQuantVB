'ExoQuantVB(ExoQuant v0.7)
'
'Copyright(c) 2019 David Benepe
'Copyright(c) 2004 Dennis Ranke
'
'Permission Is hereby granted, free Of charge, to any person obtaining a copy of
'this software And associated documentation files (the "Software"), To deal In
'the Software without restriction, including without limitation the rights To
'use, copy, modify, merge, publish, distribute, sublicense, And/Or sell copies
'of the Software, And to permit persons to whom the Software Is furnished to do
'so, subject to the following conditions:
'
'The above copyright notice And this permission notice shall be included In all
'copies Or substantial portions of the Software.
'
'THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY Of ANY KIND, EXPRESS Or
'IMPLIED, INCLUDING BUT Not LIMITED To THE WARRANTIES Of MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE And NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS Or COPYRIGHT HOLDERS BE LIABLE For ANY CLAIM, DAMAGES Or OTHER
'LIABILITY, WHETHER In AN ACTION Of CONTRACT, TORT Or OTHERWISE, ARISING FROM,
'OUT OF Or IN CONNECTION WITH THE SOFTWARE Or THE USE Or OTHER DEALINGS IN THE
'SOFTWARE.

'/******************************************************************************
'* Usage:
'* ------
'*
'* Dim exq As ExoQuantVB.ExoQuant = New ExoQuantVB.ExoQuant() // init quantizer (per image)
'* exq.Feed(<byte array of rgba32 data>) // feed pixel data (32bpp)
'* exq.Quantize(<num of colors>) // find palette
'* exq.GetPalette(<output rgba32 palette>, <num of colors>) // get palette
'* exq.MapImage(<num of pixels>, <byte array of rgba32 data>, <output index data>)
'* Or:
'* exq.MapImageOrdered(<width>, <height>, <byte array of rgba32 data>, <output index data>)
'* // map image to palette
'*
'* Notes:
'* ------
'*
'* All 32bpp data (input data And palette data) Is considered a byte stream
'* of the format:
'* R0 G0 B0 A0 R1 G1 B1 A1 ...
'* If you want to use a different order, the easiest way to do this Is to
'* change the SCALE_x constants in expquant.h, as those are the only differences
'* between the channels.
'*
'******************************************************************************/

Namespace ExoQuantVB
    Class ExoQuant
        Shared ReadOnly EXQ_HASH_BITS As Integer = 16
        Shared ReadOnly EXQ_HASH_SIZE As Integer = (1 << EXQ_HASH_BITS)

        Shared ReadOnly SCALE_R As Single = 1.0F
        Shared ReadOnly SCALE_G As Single = 1.2F
        Shared ReadOnly SCALE_B As Single = 0.8F
        Shared ReadOnly SCALE_A As Single = 1.0F

        Private pExq As ExqData
        Private sortDir As ExqColor = New ExqColor()

        Public Class ExqColor
            Public r, g, b, a As Double
        End Class

        Public Class ExqHistogramEntry
            Public color As ExqColor = New ExqColor()
            Public ored, ogreen, oblue, oalpha As Byte
            Public palIndex As Integer
            Public ditherScale As ExqColor = New ExqColor()
            Public ditherIndex As Integer() = New Integer(3) {}
            Public num As Integer = 0
            Public pNext As ExqHistogramEntry = Nothing
            Public pNextInHash As ExqHistogramEntry = Nothing
        End Class

        Public Class ExqNode
            Public dir As ExqColor = New ExqColor(), avg As ExqColor = New ExqColor()
            Public vdif As Double
            Public err As Double
            Public num As Integer
            Public pHistogram As ExqHistogramEntry = Nothing
            Public pSplit As ExqHistogramEntry = Nothing
        End Class

        Public Class ExqData
            Public pHash As ExqHistogramEntry() = New ExqHistogramEntry(EXQ_HASH_SIZE - 1) {}
            Public node As ExqNode() = New ExqNode(255) {}
            Public numColors As Integer
            Public numBitsPerChannel As Integer
            Public optimized As Boolean
            Public transparency As Boolean
        End Class

        Public Sub New()
            pExq = New ExqData()

            For i As Integer = 0 To 256 - 1
                pExq.node(i) = New ExqNode()
            Next

            For i As Integer = 0 To EXQ_HASH_SIZE - 1
                pExq.pHash(i) = Nothing
            Next

            pExq.numColors = 0
            pExq.optimized = False
            pExq.transparency = True
            pExq.numBitsPerChannel = 8
        End Sub

        Public Sub NoTransparency()
            pExq.transparency = False
        End Sub

        Private Function SafeLShift(ByVal value As UInteger, ByVal amount As UInteger) As UInteger
            Return BitConverter.ToUInt32(BitConverter.GetBytes(value << amount), 0)
        End Function
        Private Function SafeRShift(ByVal value As UInteger, ByVal amount As UInteger) As UInteger
            Return BitConverter.ToUInt32(BitConverter.GetBytes(value >> amount), 0)
        End Function

        Private Function SafeMinusForUInteger(ByVal value As UInteger, ByVal amount As UInteger) As UInteger
            If value >= amount Then
                Return value - amount
            Else
                ' amount is larger than value, so to prevent arithmetic underflow I need to
                ' simulate looping around by subtracting by the max UInteger value.
                Return UInteger.MaxValue - (amount - value)
            End If
        End Function

        Private Function MakeHash(ByVal rgba As UInteger) As UInteger
            rgba = SafeMinusForUInteger(rgba, SafeRShift(rgba, 13) Or SafeLShift(rgba, 19))
            rgba = SafeMinusForUInteger(rgba, SafeRShift(rgba, 13) Or SafeLShift(rgba, 19))
            rgba = SafeMinusForUInteger(rgba, SafeRShift(rgba, 13) Or SafeLShift(rgba, 19))
            rgba = SafeMinusForUInteger(rgba, SafeRShift(rgba, 13) Or SafeLShift(rgba, 19))
            rgba = SafeMinusForUInteger(rgba, SafeRShift(rgba, 13) Or SafeLShift(rgba, 19))
            rgba = rgba And (EXQ_HASH_SIZE - 1)
            Return rgba
        End Function

        Private Function ToRGBA(ByVal r As UInteger, ByVal g As UInteger, ByVal b As UInteger, ByVal a As UInteger) As UInteger
            Return r Or (g << 8) Or (b << 16) Or (a << 24)
        End Function

        Public Sub Feed(ByVal pData As Byte())
            Dim channelMask As Byte = &HFF00 >> pExq.numBitsPerChannel
            Dim nPixels As Integer = Int(pData.Length / 4)

            For i As Integer = 0 To nPixels - 1
                Dim r As Byte = pData(i * 4 + 0),
                    g As Byte = pData(i * 4 + 1),
                    b As Byte = pData(i * 4 + 2),
                    a As Byte = pData(i * 4 + 3)
                Dim hash As UInteger = MakeHash(ToRGBA(r, g, b, a))
                Dim pCur As ExqHistogramEntry = pExq.pHash(hash)

                While pCur IsNot Nothing AndAlso (pCur.ored <> r OrElse pCur.ogreen <> g OrElse pCur.oblue <> b OrElse pCur.oalpha <> a)
                    pCur = pCur.pNextInHash
                End While

                If pCur IsNot Nothing Then
                    pCur.num += 1
                Else
                    pCur = New ExqHistogramEntry()
                    pCur.pNextInHash = pExq.pHash(hash)
                    pExq.pHash(hash) = pCur
                    pCur.ored = r
                    pCur.ogreen = g
                    pCur.oblue = b
                    pCur.oalpha = a
                    r = r And channelMask
                    g = g And channelMask
                    b = b And channelMask
                    pCur.color.r = r / 255.0F * SCALE_R
                    pCur.color.g = g / 255.0F * SCALE_G
                    pCur.color.b = b / 255.0F * SCALE_B
                    pCur.color.a = a / 255.0F * SCALE_A

                    If pExq.transparency Then
                        pCur.color.r *= pCur.color.a
                        pCur.color.g *= pCur.color.a
                        pCur.color.b *= pCur.color.a
                    End If

                    pCur.num = 1
                    pCur.palIndex = -1
                    pCur.ditherScale.r = -1
                    pCur.ditherScale.g = -1
                    pCur.ditherScale.b = -1
                    pCur.ditherScale.a = -1
                    pCur.ditherIndex(0) = -1
                    pCur.ditherIndex(1) = -1
                    pCur.ditherIndex(2) = -1
                    pCur.ditherIndex(3) = -1
                End If
            Next
        End Sub

        Public Sub Quantize(ByVal nColors As Integer)
            QuantizeEx(nColors, False)
        End Sub

        Public Sub QuantizeHq(ByVal nColors As Integer)
            QuantizeEx(nColors, True)
        End Sub

        Public Sub QuantizeEx(ByVal nColors As Integer, ByVal hq As Boolean)
            Dim besti As Integer
            Dim beste As Double
            Dim pCur, pNext As ExqHistogramEntry
            If nColors > 256 Then nColors = 256

            If pExq.numColors = 0 Then
                pExq.node(0).pHistogram = Nothing

                For i = 0 To EXQ_HASH_SIZE - 1
                    pCur = pExq.pHash(i)

                    While pCur IsNot Nothing
                        pCur.pNext = pExq.node(0).pHistogram
                        pExq.node(0).pHistogram = pCur
                        pCur = pCur.pNextInHash
                    End While
                Next

                SumNode(pExq.node(0))
                pExq.numColors = 1
            End If

            For i = pExq.numColors To nColors - 1
                beste = 0
                besti = 0

                For j = 0 To i - 1
                    If pExq.node(j).vdif >= beste Then
                        beste = pExq.node(j).vdif
                        besti = j
                    End If
                Next

                pCur = pExq.node(besti).pHistogram
                pExq.node(besti).pHistogram = Nothing
                pExq.node(i).pHistogram = Nothing

                While pCur IsNot Nothing AndAlso pCur IsNot pExq.node(besti).pSplit
                    pNext = pCur.pNext
                    pCur.pNext = pExq.node(i).pHistogram
                    pExq.node(i).pHistogram = pCur
                    pCur = pNext
                End While

                While pCur IsNot Nothing
                    pNext = pCur.pNext
                    pCur.pNext = pExq.node(besti).pHistogram
                    pExq.node(besti).pHistogram = pCur
                    pCur = pNext
                End While

                SumNode(pExq.node(besti))
                SumNode(pExq.node(i))
                pExq.numColors = i + 1
                If hq Then OptimizePalette(1)
            Next

            pExq.optimized = False
        End Sub

        Public Function GetMeanError() As Double
            Dim n As Integer = 0
            Dim err As Double = 0

            For i As Integer = 0 To pExq.numColors - 1
                n += pExq.node(i).num
                err += pExq.node(i).err
            Next

            Return Math.Sqrt(err / n) * 256
        End Function

        Public Sub GetPalette(ByRef pPal As Byte(), ByVal nColors As Integer)
            Dim r, g, b, a As Double
            Dim channelMask As Byte = &HFF00 >> pExq.numBitsPerChannel

            pPal = New Byte(nColors * 4 - 1) {}
            If nColors > pExq.numColors Then nColors = pExq.numColors
            If Not pExq.optimized Then OptimizePalette(4)

            For i As Integer = 0 To nColors - 1
                r = pExq.node(i).avg.r
                g = pExq.node(i).avg.g
                b = pExq.node(i).avg.b
                a = pExq.node(i).avg.a

                If pExq.transparency AndAlso a <> 0 Then
                    r /= a
                    g /= a
                    b /= a
                End If

                Dim pPalIndex As Integer = i * 4
                pPal(pPalIndex + 0) = CByte(Int(r / SCALE_R * 255.9F))
                pPal(pPalIndex + 1) = CByte(Int(g / SCALE_G * 255.9F))
                pPal(pPalIndex + 2) = CByte(Int(b / SCALE_B * 255.9F))
                pPal(pPalIndex + 3) = CByte(Int(a / SCALE_A * 255.9F))

                For j As Integer = 0 To 2
                    pPal(pPalIndex + j) = (pPal(pPalIndex + j) + Int((1 << (8 - pExq.numBitsPerChannel)) / 2)) And channelMask
                Next
            Next
        End Sub

        Public Sub SetPalette(ByVal pPal As Byte(), ByVal nColors As Integer)
            pExq.numColors = nColors

            For i As Integer = 0 To nColors - 1
                pExq.node(i).avg.r = pPal(i * 4 + 0) * SCALE_R / 255.9F
                pExq.node(i).avg.g = pPal(i * 4 + 1) * SCALE_G / 255.9F
                pExq.node(i).avg.b = pPal(i * 4 + 2) * SCALE_B / 255.9F
                pExq.node(i).avg.a = pPal(i * 4 + 3) * SCALE_A / 255.9F
            Next

            pExq.optimized = True
        End Sub

        Public Sub MapImage(ByVal nPixels As Integer, ByVal pIn As Byte(), ByRef pOut As Byte())
            Dim c As ExqColor = New ExqColor()
            Dim pHist As ExqHistogramEntry
            pOut = New Byte(nPixels - 1) {}
            If Not pExq.optimized Then OptimizePalette(4)

            For i = 0 To nPixels - 1
                pHist = FindHistogram(pIn, i)

                If pHist IsNot Nothing AndAlso pHist.palIndex <> -1 Then
                    pOut(i) = CByte(pHist.palIndex)
                Else
                    c.r = pIn(i * 4 + 0) / 255.0F * SCALE_R
                    c.g = pIn(i * 4 + 1) / 255.0F * SCALE_G
                    c.b = pIn(i * 4 + 2) / 255.0F * SCALE_B
                    c.a = pIn(i * 4 + 3) / 255.0F * SCALE_A

                    If pExq.transparency Then
                        c.r *= c.a
                        c.g *= c.a
                        c.b *= c.a
                    End If

                    pOut(i) = FindNearestColor(c)
                    If pHist IsNot Nothing Then pHist.palIndex = i
                End If
            Next
        End Sub

        Public Sub MapImageOrdered(ByVal width As Integer, ByVal height As Integer, ByVal pIn As Byte(), ByRef pOut As Byte())
            MapImageDither(width, height, pIn, pOut, True)
        End Sub

        Public Sub MapImageRandom(ByVal nPixels As Integer, ByVal pIn As Byte(), ByRef pOut As Byte())
            MapImageDither(nPixels, 1, pIn, pOut, False)
        End Sub

        Private ReadOnly random As Random = New Random()

        Private Sub MapImageDither(ByVal width As Integer, ByVal height As Integer, ByVal pIn As Byte(), ByRef pOut As Byte(), ByVal ordered As Boolean)
            Dim ditherMatrix As Double() = {-0.375, 0.125, 0.375, -0.125}
            Dim i, j, d As Integer
            Dim p As ExqColor = New ExqColor(), scale As ExqColor = New ExqColor(), tmp As ExqColor = New ExqColor()
            Dim pHist As ExqHistogramEntry
            pOut = New Byte(width * height - 1) {}
            If Not pExq.optimized Then OptimizePalette(4)

            For y As Integer = 0 To height - 1

                For x As Integer = 0 To width - 1
                    Dim index As Integer = y * width + x

                    If ordered Then
                        d = (x And 1) + (y And 1) * 2
                    Else
                        d = random.[Next]() And 3
                    End If

                    pHist = FindHistogram(pIn, index)
                    p.r = pIn(index * 4 + 0) / 255.0F * SCALE_R
                    p.g = pIn(index * 4 + 1) / 255.0F * SCALE_G
                    p.b = pIn(index * 4 + 2) / 255.0F * SCALE_B
                    p.a = pIn(index * 4 + 3) / 255.0F * SCALE_A

                    If pExq.transparency Then
                        p.r *= p.a
                        p.g *= p.a
                        p.b *= p.a
                    End If

                    If pHist Is Nothing OrElse pHist.ditherScale.r < 0 Then
                        i = FindNearestColor(p)
                        scale.r = pExq.node(i).avg.r - p.r
                        scale.g = pExq.node(i).avg.g - p.g
                        scale.b = pExq.node(i).avg.b - p.b
                        scale.a = pExq.node(i).avg.a - p.a
                        tmp.r = p.r - scale.r / 3
                        tmp.g = p.g - scale.g / 3
                        tmp.b = p.b - scale.b / 3
                        tmp.a = p.a - scale.a / 3
                        j = FindNearestColor(tmp)

                        If i = j Then
                            tmp.r = p.r - scale.r * 3
                            tmp.g = p.g - scale.g * 3
                            tmp.b = p.b - scale.b * 3
                            tmp.a = p.a - scale.a * 3
                            j = FindNearestColor(tmp)
                        End If

                        If i <> j Then
                            scale.r = (pExq.node(j).avg.r - pExq.node(i).avg.r) * 0.8F
                            scale.g = (pExq.node(j).avg.g - pExq.node(i).avg.g) * 0.8F
                            scale.b = (pExq.node(j).avg.b - pExq.node(i).avg.b) * 0.8F
                            scale.a = (pExq.node(j).avg.a - pExq.node(i).avg.a) * 0.8F
                            If scale.r < 0 Then scale.r = -scale.r
                            If scale.g < 0 Then scale.g = -scale.g
                            If scale.b < 0 Then scale.b = -scale.b
                            If scale.a < 0 Then scale.a = -scale.a
                        Else
                            scale.r = 0
                            scale.g = 0
                            scale.b = 0
                            scale.a = 0
                        End If

                        If pHist IsNot Nothing Then
                            pHist.ditherScale.r = scale.r
                            pHist.ditherScale.g = scale.g
                            pHist.ditherScale.b = scale.b
                            pHist.ditherScale.a = scale.a
                        End If
                    Else
                        scale.r = pHist.ditherScale.r
                        scale.g = pHist.ditherScale.g
                        scale.b = pHist.ditherScale.b
                        scale.a = pHist.ditherScale.a
                    End If

                    If pHist IsNot Nothing AndAlso pHist.ditherIndex(d) >= 0 Then
                        pOut(index) = CByte(pHist.ditherIndex(d))
                    Else
                        tmp.r = p.r + scale.r * ditherMatrix(d)
                        tmp.g = p.g + scale.g * ditherMatrix(d)
                        tmp.b = p.b + scale.b * ditherMatrix(d)
                        tmp.a = p.a + scale.a * ditherMatrix(d)
                        pOut(index) = FindNearestColor(tmp)

                        If pHist IsNot Nothing Then
                            pHist.ditherIndex(d) = pOut(index)
                        End If
                    End If
                Next
            Next
        End Sub

        Private Sub SumNode(ByVal pNode As ExqNode)
            Dim n, n2 As Integer
            Dim fsum As ExqColor = New ExqColor(), fsum2 As ExqColor = New ExqColor(), vc As ExqColor = New ExqColor(), tmp As ExqColor = New ExqColor(), tmp2 As ExqColor = New ExqColor(), sum As ExqColor = New ExqColor(), sum2 As ExqColor = New ExqColor()
            Dim pCur As ExqHistogramEntry
            Dim isqrt, nv, v As Double
            n = 0
            fsum.r = 0
            fsum.g = 0
            fsum.b = 0
            fsum.a = 0
            fsum2.r = 0
            fsum2.g = 0
            fsum2.b = 0
            fsum2.a = 0
            pCur = pNode.pHistogram

            While pCur IsNot Nothing
                n += pCur.num
                fsum.r += pCur.color.r * pCur.num
                fsum.g += pCur.color.g * pCur.num
                fsum.b += pCur.color.b * pCur.num
                fsum.a += pCur.color.a * pCur.num
                fsum2.r += pCur.color.r * pCur.color.r * pCur.num
                fsum2.g += pCur.color.g * pCur.color.g * pCur.num
                fsum2.b += pCur.color.b * pCur.color.b * pCur.num
                fsum2.a += pCur.color.a * pCur.color.a * pCur.num
                pCur = pCur.pNext
            End While

            pNode.num = n

            If n = 0 Then
                pNode.vdif = 0
                pNode.err = 0
                Return
            End If

            pNode.avg.r = fsum.r / n
            pNode.avg.g = fsum.g / n
            pNode.avg.b = fsum.b / n
            pNode.avg.a = fsum.a / n
            vc.r = fsum2.r - fsum.r * pNode.avg.r
            vc.g = fsum2.g - fsum.g * pNode.avg.g
            vc.b = fsum2.b - fsum.b * pNode.avg.b
            vc.a = fsum2.a - fsum.a * pNode.avg.a
            v = vc.r + vc.g + vc.b + vc.a
            pNode.err = v
            pNode.vdif = -v

            If vc.r > vc.g AndAlso vc.r > vc.b AndAlso vc.r > vc.a Then
                Sort(pNode.pHistogram, New SortFunction(AddressOf SortByRed))
            ElseIf vc.g > vc.b AndAlso vc.g > vc.a Then
                Sort(pNode.pHistogram, New SortFunction(AddressOf SortByGreen))
            ElseIf vc.b > vc.a Then
                Sort(pNode.pHistogram, New SortFunction(AddressOf SortByBlue))
            Else
                Sort(pNode.pHistogram, New SortFunction(AddressOf SortByAlpha))
            End If

            pNode.dir.r = 0
            pNode.dir.g = 0
            pNode.dir.b = 0
            pNode.dir.a = 0
            pCur = pNode.pHistogram

            While pCur IsNot Nothing
                tmp.r = (pCur.color.r - pNode.avg.r) * pCur.num
                tmp.g = (pCur.color.g - pNode.avg.g) * pCur.num
                tmp.b = (pCur.color.b - pNode.avg.b) * pCur.num
                tmp.a = (pCur.color.a - pNode.avg.a) * pCur.num

                If tmp.r * pNode.dir.r + tmp.g * pNode.dir.g + tmp.b * pNode.dir.b + tmp.a * pNode.dir.a < 0 Then
                    tmp.r = -tmp.r
                    tmp.g = -tmp.g
                    tmp.b = -tmp.b
                    tmp.a = -tmp.a
                End If

                pNode.dir.r += tmp.r
                pNode.dir.g += tmp.g
                pNode.dir.b += tmp.b
                pNode.dir.a += tmp.a
                pCur = pCur.pNext
            End While

            isqrt = 1 / Math.Sqrt(pNode.dir.r * pNode.dir.r + pNode.dir.g * pNode.dir.g + pNode.dir.b * pNode.dir.b + pNode.dir.a * pNode.dir.a)
            pNode.dir.r *= isqrt
            pNode.dir.g *= isqrt
            pNode.dir.b *= isqrt
            pNode.dir.a *= isqrt

            sortDir = pNode.dir
            Sort(pNode.pHistogram, New SortFunction(AddressOf SortByDir))
            sum.r = 0
            sum.g = 0
            sum.b = 0
            sum.a = 0
            sum2.r = 0
            sum2.g = 0
            sum2.b = 0
            sum2.a = 0
            n2 = 0
            pNode.pSplit = pNode.pHistogram
            pCur = pNode.pHistogram

            While pCur IsNot Nothing
                If pNode.pSplit Is Nothing Then pNode.pSplit = pCur
                n2 += pCur.num
                sum.r += pCur.color.r * pCur.num
                sum.g += pCur.color.g * pCur.num
                sum.b += pCur.color.b * pCur.num
                sum.a += pCur.color.a * pCur.num
                sum2.r += pCur.color.r * pCur.color.r * pCur.num
                sum2.g += pCur.color.g * pCur.color.g * pCur.num
                sum2.b += pCur.color.b * pCur.color.b * pCur.num
                sum2.a += pCur.color.a * pCur.color.a * pCur.num
                If n = n2 Then Exit While
                tmp.r = sum2.r - sum.r * sum.r / n2
                tmp.g = sum2.g - sum.g * sum.g / n2
                tmp.b = sum2.b - sum.b * sum.b / n2
                tmp.a = sum2.a - sum.a * sum.a / n2
                tmp2.r = (fsum2.r - sum2.r) - (fsum.r - sum.r) * (fsum.r - sum.r) / (n - n2)
                tmp2.g = (fsum2.g - sum2.g) - (fsum.g - sum.g) * (fsum.g - sum.g) / (n - n2)
                tmp2.b = (fsum2.b - sum2.b) - (fsum.b - sum.b) * (fsum.b - sum.b) / (n - n2)
                tmp2.a = (fsum2.a - sum2.a) - (fsum.a - sum.a) * (fsum.a - sum.a) / (n - n2)
                nv = tmp.r + tmp.g + tmp.b + tmp.a + tmp2.r + tmp2.g + tmp2.b + tmp2.a

                If -nv > pNode.vdif Then
                    pNode.vdif = -nv
                    pNode.pSplit = Nothing
                End If

                pCur = pCur.pNext
            End While

            If pNode.pSplit Is pNode.pHistogram Then pNode.pSplit = pNode.pSplit.pNext
            pNode.vdif += v
        End Sub

        Private Sub OptimizePalette(ByVal iter As Integer)
            Dim pCur As ExqHistogramEntry
            pExq.optimized = True

            For n As Integer = 0 To iter - 1

                For i As Integer = 0 To pExq.numColors - 1
                    pExq.node(i).pHistogram = Nothing
                Next

                For i As Integer = 0 To EXQ_HASH_SIZE - 1
                    pCur = pExq.pHash(i)

                    While pCur IsNot Nothing
                        Dim j As Byte = FindNearestColor(pCur.color)
                        pCur.pNext = pExq.node(j).pHistogram
                        pExq.node(j).pHistogram = pCur
                        pCur = pCur.pNextInHash
                    End While
                Next

                For i As Integer = 0 To pExq.numColors - 1
                    SumNode(pExq.node(i))
                Next
            Next
        End Sub

        Private Function FindNearestColor(ByVal pColor As ExqColor) As Byte
            Dim dif As ExqColor = New ExqColor()
            Dim bestv As Double = 16
            Dim besti As Integer = 0

            For i As Integer = 0 To pExq.numColors - 1
                dif.r = pColor.r - pExq.node(i).avg.r
                dif.g = pColor.g - pExq.node(i).avg.g
                dif.b = pColor.b - pExq.node(i).avg.b
                dif.a = pColor.a - pExq.node(i).avg.a

                If dif.r * dif.r + dif.g * dif.g + dif.b * dif.b + dif.a * dif.a < bestv Then
                    bestv = dif.r * dif.r + dif.g * dif.g + dif.b * dif.b + dif.a * dif.a
                    besti = i
                End If
            Next

            Return besti
        End Function

        Private Function FindHistogram(ByVal pCol As Byte(), ByVal index As Integer) As ExqHistogramEntry
            Dim hash As UInteger
            Dim pCur As ExqHistogramEntry
            Dim r As UInteger = pCol(index * 4 + 0),
                g As UInteger = pCol(index * 4 + 1),
                b As UInteger = pCol(index * 4 + 2),
                a As UInteger = pCol(index * 4 + 3)

            hash = MakeHash(ToRGBA(r, g, b, a))
            pCur = pExq.pHash(hash)

            While pCur IsNot Nothing AndAlso (pCur.ored <> r OrElse pCur.ogreen <> g OrElse pCur.oblue <> b OrElse pCur.oalpha <> a)
                pCur = pCur.pNextInHash
            End While

            Return pCur
        End Function

        Public Delegate Function SortFunction(ByVal pHist As ExqHistogramEntry) As Double

        Private Sub Sort(ByRef ppHist As ExqHistogramEntry, ByVal sortfunc As SortFunction)
            Dim pLow, pHigh, pCur, pNext As ExqHistogramEntry
            Dim n As Integer = 0
            Dim sum As Double = 0
            pCur = ppHist

            While pCur IsNot Nothing
                n += 1
                sum += sortfunc(pCur)
                pCur = pCur.pNext
            End While

            If n < 2 Then Return

            sum /= n

            pLow = Nothing
            pHigh = Nothing
            pCur = ppHist

            While pCur IsNot Nothing
                pNext = pCur.pNext

                If sortfunc(pCur) < sum Then
                    pCur.pNext = pLow
                    pLow = pCur
                Else
                    pCur.pNext = pHigh
                    pHigh = pCur
                End If

                pCur = pNext
            End While

            If pLow Is Nothing Then
                ppHist = pHigh
                Return
            End If

            If pHigh Is Nothing Then
                ppHist = pLow
                Return
            End If

            Sort(pLow, sortfunc)
            Sort(pHigh, sortfunc)
            ppHist = pLow

            While pLow.pNext IsNot Nothing
                pLow = pLow.pNext
            End While

            pLow.pNext = pHigh
        End Sub

        Private Function SortByRed(ByVal pHist As ExqHistogramEntry) As Double
            Return pHist.color.r
        End Function

        Private Function SortByGreen(ByVal pHist As ExqHistogramEntry) As Double
            Return pHist.color.g
        End Function

        Private Function SortByBlue(ByVal pHist As ExqHistogramEntry) As Double
            Return pHist.color.b
        End Function

        Private Function SortByAlpha(ByVal pHist As ExqHistogramEntry) As Double
            Return pHist.color.a
        End Function

        Private Function SortByDir(ByVal pHist As ExqHistogramEntry) As Double
            Return pHist.color.r * sortDir.r +
                pHist.color.g * sortDir.g +
                pHist.color.b * sortDir.b +
                pHist.color.a * sortDir.a
        End Function
    End Class
End Namespace
