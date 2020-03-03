Module modMain

    Public Const MISSING_P As Single = -99.99
    Public Const MISSING_SPI As Single = -99.99
    Public Const SPI_In_DIR As String = "E:\GriddedWeather\MonthlyPrecip_Selected\"
    Public Const SPI_OUT_EXT As String = ".csv"
    Public Const SPI_Out_DIR As String = "E:\GriddedWeather\SPI\"
    Public Const CAL_OUT_DIR As String = "E:\GriddedWeather\SPICal\"


    Sub Main()
        RunSPI()
    End Sub
    Private Sub RunSPI()

        Dim SPI As modCalculateSPI = New modCalculateSPI()
        SPI.ProcessSPI()
    End Sub

End Module
