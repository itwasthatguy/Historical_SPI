Imports System.Math, System.IO
Imports System.ComponentModel

'       Standardized Precipitation Index.
'
'      Usage: spi length [length ...] [<infile] [>outfile]
'
'      Where:
'         run_length - integer > 0; running sum length for SPI
'        infile - optional input file; if omitted, spi reads
'                  from standard input.
'         outfile - optional output file; if omitted, spi writes
'                  to standard output.
'
'      Notes on the ForTran version:
'
'      1) Lengths are hard coded to be 1, 2, 3, 6, 9, 12, 24, 36, 48 and 60 months.  This is
'      necessasry since ForTran has no standard way of passing command
'      line arguments.
'
'      2) System dependent code is bracketed with 'c*****'s

'      Following class is VB Version

Public Class modCalculateSPI

    Private Const iMaxYears As Integer = 66
    Private Const iBegYear As Integer = 1950
    Private Const iTotalPeriods As Integer = iMaxYears * 12


    Private iCountPeriods As Integer = 0
    Private iTimeLen() As Integer = New Integer() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 15, 18, 21, 24}

    Private sPrecip(iTotalPeriods - 1) As Single
    Private sSpi(,) As Single
    Private strObsDate(iTotalPeriods - 1) As String
    Private strHeader As String
    Private strID As String

    Private Sub Init()
        For iIndex As Integer = 0 To iTotalPeriods - 1
            sPrecip(iIndex) = MISSING_P
            strObsDate(iIndex) = CStr(MISSING_P)
        Next
    End Sub

    ' Call this function to start the process of calculating SPI for each station
    Public Sub ProcessSPI()
        Dim sStationsList() As String = Directory.GetFiles(SPI_IN_DIR, "*.csv")
        Dim iCountStn As Integer = 0

        ' Process each station 
        For Each sStationInputFile As String In sStationsList

            'Init arrays
            Init()          'innit

            ' Read station's input file,and to get preciptation and count how many periods 
            ReadInput(sStationInputFile)

            ' To get station ID
            Dim strStnID As String = Path.GetFileNameWithoutExtension(sStationInputFile)

            ' The array to store SPI for just one station 
            ReDim sSpi(iCountPeriods - 1, iTimeLen.Length - 1)

            strID = strStnID

            ' Compute SPI based on time length 
            For i As Integer = 0 To iTimeLen.Length - 1
                SPIGam(iTimeLen(i), sSpi, strStnID)
            Next

            ' write output
            WriteOutputFile(strStnID, sSpi)

            ' report progress
            iCountStn += 1
            iCountPeriods = 0
        Next

    End Sub

    ' Read station preciptation and count periods
    Private Sub ReadInput(ByVal strStationInputFile As String)
        Const INDEX_YEAR As Integer = 0
        Const INDEX_MONTH As Integer = 1
        Const INDEX_PRECIP As Integer = 2

        'Dim iCountMonth As Integer = 0
        Dim sr As StreamReader
        Dim sLine As String = ""

        ' Read input file
        sr = New StreamReader(New FileStream(strStationInputFile, FileMode.Open, FileAccess.Read, FileShare.Read))

        ' Record head line
        'strHeader = sr.ReadLine()

        'read rest of lines 
        While sr.Peek > -1
            sLine = sr.ReadLine()
            Dim sParts() As String = sLine.Split(New String() {}, StringSplitOptions.RemoveEmptyEntries)

            If CInt(sParts(INDEX_YEAR)) >= iBegYear Then
                sPrecip(iCountPeriods) = CSng(sParts(INDEX_PRECIP))
                strObsDate(iCountPeriods) = sParts(INDEX_YEAR) & "  " & sParts(INDEX_MONTH)
                iCountPeriods += 1
            End If

        End While
        sr.Close()
    End Sub

    ''' <summary>
    ''' Calculate indices assuming incomplete gamma distribution.
    ''' </summary>
    ''' <param name="nrun">pass in iTimeLen(i) </param>
    ''' <remarks></remarks>
    Private Sub SPIGam(ByVal nrun As Integer, ByRef sSpi(,) As Single, ByVal strStnID As String)
        'Dim sIndex(iTotalPeriods - 1) As Single
        Dim tmpArry(iMaxYears - 1) As Single
        Dim sBeta(12 - 1), sGamm(12 - 1), sPzero(12 - 1) As Single
        Dim im As Integer
        Dim iIndexOfTimeLen As Integer = Array.IndexOf(iTimeLen, nrun)

        ' init sSpi(i, iIndexOfTimeLen) The first nrun-1 index values will be missing.
        If Not iIndexOfTimeLen = 0 Then
            For i As Integer = 0 To nrun - 1 - 1
                sSpi(i, iIndexOfTimeLen) = MISSING_SPI
            Next
        End If

        'Sum nrun precip. values; 
        'store them in the appropriate index location.
        'If any value is missing; set the sum to missing.         
        For j As Integer = nrun To iCountPeriods
            sSpi(j - 1, iIndexOfTimeLen) = 0
            For i As Integer = 0 To nrun - 1
                If Not sPrecip(j - 1 - i) = MISSING_P Then
                    sSpi(j - 1, iIndexOfTimeLen) += sPrecip(j - 1 - i)
                Else
                    sSpi(j - 1, iIndexOfTimeLen) = MISSING_SPI
                    Exit For
                End If
            Next
        Next
        'End If

        'index(month) = sum(pp(i),i=month..month-period) where period = nrun
        'therefore index(i) = sum of nrun months of precipitation prior to month i (inclusive)
        '
        ' For nrun<12, the monthly distributions will be substantially
        'different.  So we need to compute gamma parameters for
        'each month starting with the (nrun-1)th.
        For i As Integer = 0 To 11
            Dim n As Integer = 0
            'nrun = number of months for this spi calculation (1,2,3,6, ...)
            For j As Integer = nrun + i - 1 To iCountPeriods - 1 Step 12
                'maxyrs*12 - nrun) / 12 times
                'go through each month i
                If Not sSpi(j, iIndexOfTimeLen) = MISSING_SPI Then
                    tmpArry(n) = sSpi(j, iIndexOfTimeLen)
                    n = n + 1
                End If
            Next
            im = (nrun + i - 1) Mod 12 + 1
            GamFit(tmpArry, n, sBeta(im - 1), sGamm(im - 1), sPzero(im - 1))
        Next

        Dim listofin As New List(Of Single)
        listofin.AddRange(sBeta)
        '    Replace precip. sums stored in index with SPI's

        'Save sBeta, sGamm, and sPzero here!
        WriteCalibrationFile(strStnID, nrun, sBeta, sGamm, sPzero)

        For j As Integer = nrun - 1 To iCountPeriods - 1
            im = j Mod 12 + 1
            If Not sSpi(j, iIndexOfTimeLen) = MISSING_SPI Then

                '    Get the probability
                sSpi(j, iIndexOfTimeLen) = gamcdf(sBeta(im - 1), sGamm(im - 1), sPzero(im - 1), sSpi(j, iIndexOfTimeLen))

                '     Convert prob. to z value. 
                sSpi(j, iIndexOfTimeLen) = anvnrm(sSpi(j, iIndexOfTimeLen))
            End If
        Next

    End Sub

    Private Function YesWrite() As Boolean
        If strID.Equals("1163842") Then Return True
        If strID.Equals("1161662") Then Return True
        If strID.Equals("4016651") Then Return True
        If strID.Equals("4020286") Then Return True
        If strID.Equals("4031844") Then Return True
        If strID.Equals("4032322") Then Return True
        Return False
    End Function

    ''' <summary>
    ''' See Abromowitz and Stegun _Handbook of Mathematical Functions_, p. 933
    ''' </summary>
    ''' <param name="sProb"> input prob;</param>
    ''' <returns> return z.</returns>
    ''' <remarks></remarks>
    Private Function anvnrm(ByVal sProb As Single) As Single
        Dim sC0 As Single = 2.515517
        Dim sC1 As Single = 0.802853
        Dim sC2 As Single = 0.010328
        Dim sD1 As Single = 1.432788
        Dim sD2 As Single = 0.189269
        Dim sD3 As Single = 0.001308

        Dim sSign As Single

        If (sProb > 0.5) Then
            sSign = 1.0
            sProb = 1.0 - sProb
        Else
            sSign = -1.0
        End If

        If (sProb < 0.0) Then
            'write(0, *) 'Error in anvnrm(). Prob. not in [0,1.0]'
            Return 0.0
        End If

        If (sProb = 0.0) Then
            ' we shouldnt get a probability zero(or 100); I think this value was here as a flag which would point to an issue
            'Return 1.0E+37 * sSign
            ' it is still an issue but we will now just set the return value very high but within the range of comfort for printing
            Return 6
        End If

        Dim sT As Single = Sqrt(Log(1.0 / (sProb * sProb)))
        anvnrm = (sSign * (sT - ((((sC2 * sT) + sC1) * sT) + sC0) / ((((((sD3 * sT) + sD2) * sT) + sD1) * sT) + 1.0)))

    End Function

    ''' <summary>
    ''' Estimate incomplete gamma parameters.
    ''' </summary>
    ''' <param name="dataArr">data array</param>
    ''' <param name="n">size of data array</param>
    ''' <output> beta, gamma - gamma parametersprosb
    ''' pzero - probability of zero.</output>
    ''' <remarks></remarks>
    Private Sub GamFit(ByVal dataArr() As Single, ByVal n As Integer, ByRef sBeta As Single, ByRef sGamm As Single, ByRef sPzero As Single)
        Dim sAlpha As Single
        Dim sSum As Single = 0.0
        Dim sSumlog As Single = 0.0
        Dim sAve As Single = 0
        Dim iNact As Integer = 0
        sPzero = 0
        If dataArr.Length = 0 Then
            Exit Sub
        End If

        '   compute sums
        For i As Integer = 0 To n - 1
            If (dataArr(i) > 0.0) Then
                sSum += dataArr(i)
                sSumlog += Log(dataArr(i))
                iNact += 1
            Else
                sPzero += 1
            End If
        Next
        sPzero /= n
        If Not iNact = 0.0 Then sAve = sSum / iNact

        '     Bogus data array but do something reasonable
        If (iNact = 1) Then
            sAlpha = 0.0
            sGamm = 1.0
            sBeta = sAve
            Return
        End If

        '     They were all zeroes. 
        If (sPzero = 1.0) Then
            sAlpha = 0.0
            sGamm = 1.0
            sBeta = sAve
            Return
        End If

        'Use(MLE)
        sAlpha = Log(sAve) - (sSumlog / iNact)
        sGamm = (1.0 + Sqrt(1.0 + (4.0 * sAlpha / 3.0))) / (4.0 * sAlpha)
        sBeta = sAve / sGamm
    End Sub

    ''' <summary>
    ''' Compute probability of a less than x using incomplete gamma parameters.
    ''' </summary>
    ''' <param name="sBeta">beta, - gamma parameters</param>
    ''' <param name="sGamm"> gamma - gamma parameters</param>
    ''' <param name="sPzero">pzero - probability of zero.</param>
    ''' <param name="sValue"> svalue - sSpi(,)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function gamcdf(ByVal sBeta As Single, ByVal sGamm As Single, ByVal sPzero As Single, ByVal sValue As Single) As Single
        If (sValue <= 0.0) Then
            gamcdf = sPzero
        Else
            'If sValue > 500 Then sValue = sValue
            Dim sTmpValue As Single = sValue / sBeta
            gamcdf = sPzero + (1.0 - sPzero) * gammap(sGamm, sTmpValue)
        End If

    End Function

    '     Evaluate the incomplete gamma function P(a,x), choosing the most 
    '     appropriate representation.
    Private Function gammap(ByVal sGamm As Single, ByVal sValue As Single) As Single
        If (sValue < sGamm + 1.0) Then
            gammap = gamser(sGamm, sValue)
        Else
            gammap = 1.0 - gammcf(sGamm, sValue)
        End If
    End Function


    ' Evaluate P(sGamm,x) by its series representation.  
    Private Function gamser(ByVal sGamm As Single, ByVal x As Single) As Single
        '     Maximum number of iterations, and bound on error.
        Dim iMaxitr As Integer = 1000
        Dim sEps As Single = 0.0000003
        Dim iWarn As Integer = 0
        Dim sGln As Single = gammln(sGamm)

        If (x = 0.0) Then
            Return 0
        End If
        Dim sAp As Single = sGamm
        Dim sSum As Single = 1.0 / sGamm
        Dim sDel As Single = sSum

        For i As Integer = 0 To iMaxitr - 1
            sAp += 1
            sDel *= x / sAp
            sSum += sDel

        Next

        gamser = sSum * Exp(-x + sGamm * Log(x) - sGln)

    End Function

    '     For those who don't have a ln(gamma) function.
    Private Function gammln(ByVal sGamm As Single) As Single
        Dim sCof() As Single = New Single() {76.18009173, -86.50532033, 24.01409822, -1.231739516,
                                                0.00120858003, -0.00000536382}
        Dim sX As Single = sGamm - 1.0
        Dim sTmp As Single = sX + 5.5
        Dim sSer As Single = 1

        sTmp -= (sX + 0.5) * Log(sTmp)
        For i As Integer = 0 To 5 - 1
            sX += 1
            sSer += sCof(i) / sX
        Next

        gammln = -sTmp + Log(2.50662827465 * sSer)

    End Function

    '     Evaluate P(a,x) in its continued fraction representation.
    Private Function gammcf(ByVal sGamm As Single, ByVal sValue As Single)
        Dim iMaxitr As Integer = 1000
        Dim sEps As Single = 0.0000003
        Dim sG As Single = 0.0

        Dim sGln As Single = gammln(sGamm)
        Dim sGold As Single = 0.0
        Dim sA0 As Single = 1.0
        Dim sA1 As Single = sValue
        Dim sB0 As Single = 0
        Dim sB1 As Single = 1.0
        Dim sFac As Single = 1.0

        For i As Integer = 1 To iMaxitr
            Dim iAn As Integer = i
            Dim sAna As Single = iAn - sGamm
            sA0 = (sA1 + sA0 * sAna) * sFac
            sB0 = (sB1 + sB0 * sAna) * sFac
            Dim sAnf As Single = iAn * sFac
            sA1 = sValue * sA0 + sAnf * sA1
            sB1 = sValue * sB0 + sAnf * sB1
            If Not sA1 = 0 Then
                sFac = 1 / sA1
                sG = sB1 * sFac
                If Abs((sG - sGold) / sG) >= sEps Then
                    sGold = sG
                End If
            End If
        Next

        gammcf = sG * Exp(-sValue + sGamm * Log(sValue) - sGln)

    End Function
    Private Sub WriteCalibrationFile(ByVal strStnID As String, Run As Integer, ByVal Beta() As Single, ByVal Gamm() As Single, ByVal Pzero() As Single)
        'Write the calibrated fitting values so they can be reused
        Dim RunString As String = Str(Run)
        RunString = "_" & RunString & "_"

        Dim strOutFile As String = Path.GetDirectoryName(CAL_OUT_DIR) & "\" & strStnID & RunString & SPI_OUT_EXT
        Dim sw As StreamWriter

        If File.Exists(strOutFile) Then
            File.Delete(strOutFile)
        End If

        sw = New StreamWriter(New FileStream(strOutFile, FileMode.Create))

        Dim sOutputLine As String = "Beta "
        ' get SPIs for the date based on time length
        For j As Integer = 0 To 11
            sOutputLine &= " " & Beta(j).ToString()
        Next

        sw.WriteLine(sOutputLine)

        sOutputLine = "Gamma "
        ' get SPIs for the date based on time length
        For j As Integer = 0 To 11
            sOutputLine &= " " & Gamm(j).ToString()
        Next

        sw.WriteLine(sOutputLine)

        sOutputLine = "Pzero "
        ' get SPIs for the date based on time length
        For j As Integer = 0 To 11
            sOutputLine &= " " & Pzero(j).ToString()
        Next

        sw.WriteLine(sOutputLine)

        sw.Close()

    End Sub
    Private Sub WriteOutputFile(ByVal strStnID As String, ByVal sSPI(,) As Single)
        ' write the output file
        Dim strOutFile As String = Path.GetDirectoryName(SPI_OUT_DIR) & "\" & strStnID & SPI_OUT_EXT
        Dim sw As StreamWriter

        If File.Exists(strOutFile) Then
            File.Delete(strOutFile)
        End If

        sw = New StreamWriter(New FileStream(strOutFile, FileMode.Create))

        'write header
        sw.WriteLine(strHeader)
        ' write the rest of parts (sSPI(,))
        For i As Integer = 0 To iCountPeriods - 1
            ' get the date
            Dim sOutputLine As String = " " & strObsDate(i)
            ' get SPIs for the date based on time length
            For j As Integer = 0 To iTimeLen.Length - 1
                sOutputLine &= " " & sSPI(i, j).ToString("0.00")
            Next
            ' format the output line, and write it 
            sw.WriteLine(Format(sOutputLine))
        Next
        sw.Close()

    End Sub

    ' format line structure for each line of the output file 
    Private Function Format(ByVal strOutputLine As String) As String
        Const INDEX_YEAR As Integer = 0
        Const INDEX_MONTH As Integer = 1

        Dim sParts() As String = strOutputLine.Split(New String() {}, StringSplitOptions.RemoveEmptyEntries)

        Format = ""
        For i As Integer = 0 To sParts.Length - 1

            If i = INDEX_YEAR Then
                Format &= sParts(i).ToString.PadLeft(5)
            ElseIf i = INDEX_MONTH Then
                Format &= sParts(i).ToString.PadLeft(3)
            Else
                Format &= (CDbl(sParts(i))).ToString("0.00").PadLeft(7)
            End If
        Next
    End Function

End Class
