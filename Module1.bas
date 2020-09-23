Attribute VB_Name = "Module1"
Option Explicit

Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Const CCHDEVICENAME& = 32
Private Const CCHFORMNAME& = 32
Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
End Type

Type JOB_INFO_1
    JobId As Long
    pPrinterName As Long
    pMachineName As Long
    pUserName As Long
    pDocument As Long
    pDatatype As Long
    pStatus As Long
    status As Long
    Priority As Long
    Position As Long
    TotalPages As Long
    PagesPrinted As Long
    Submitted As SYSTEMTIME
End Type


Type PRINTER_DEFAULTS
        pDatatype As Long
        pDevMode As Long
        DesiredAccess As Long
End Type

Declare Function OpenPrinter& Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS)
Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, pJob As JOB_INFO_1, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Declare Function SetJob Lib "winspool.drv" Alias "SetJobA" (ByVal hPrinter As Long, ByVal JobId As Long, ByVal Level As Long, pJob As JOB_INFO_1, ByVal Command As Long) As Long
Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long

Public Const JOB_STATUS_DELETING& = &H4
Public Const JOB_STATUS_ERROR& = &H2
Public Const JOB_STATUS_OFFLINE& = &H20
Public Const JOB_STATUS_PAPEROUT& = &H40
Public Const JOB_STATUS_PAUSED& = &H1
Public Const JOB_STATUS_PRINTED& = &H80
Public Const JOB_STATUS_PRINTING& = &H10
Public Const JOB_STATUS_SPOOLING& = &H8

Function InizioStampa() As String
    Dim PD          As PRINTER_DEFAULTS
    Dim Ji          As JOB_INFO_1
    Dim vetJobs()   As JOB_INFO_1
    Dim Esito       As Long
    Dim hPrinter    As Long
    Dim FJob        As Long
    Dim NJob        As Long
    Dim richiesto   As Long
    Dim ritorno     As Long
    Dim Stampante   As Printer
    Dim dt          As Date
    Dim dtJob       As Date
    Dim Ix          As Integer
    Dim Secondi     As Integer
    
    '
    ' Questa funzione inizializza una nuova procedura di stampa, quindi
    ' cerca nello spool di stampa l'identificativo del JOB di questa nuova
    ' stampa
    '
    
    ' Avvio la stampa
    Printer.Print ""
    Call GetSystemTime(Ji.Submitted)
    dt = DateSerial(Ji.Submitted.wYear, Ji.Submitted.wMonth, Ji.Submitted.wDay) + TimeSerial(Ji.Submitted.wHour, Ji.Submitted.wMinute, Ji.Submitted.wSecond)

    
    hPrinter = 0
    Esito = OpenPrinter(Printer.DeviceName, hPrinter, PD)
    If Esito = 0 Then Exit Function
    
    richiesto = 0
    ritorno = 0
    Esito = EnumJobs(hPrinter, 0, 1000, 1, Ji, 0, richiesto, ritorno)
    NJob = richiesto / Len(Ji)
    ReDim vetJobs(NJob)
    Esito = EnumJobs(hPrinter, 0, NJob, 1, vetJobs(0), Len(Ji) * NJob, richiesto, ritorno)
    
    If Esito Then
        For Secondi = 0 To 60
            For Ix = 0 To ritorno - 1
                dtJob = DateSerial(vetJobs(Ix).Submitted.wYear, vetJobs(Ix).Submitted.wMonth, vetJobs(Ix).Submitted.wDay) + TimeSerial(vetJobs(Ix).Submitted.wHour, vetJobs(Ix).Submitted.wMinute, vetJobs(Ix).Submitted.wSecond)
                If (vetJobs(Ix).status And JOB_STATUS_SPOOLING) <> 0 Then
                    ' controllo se questo JOB e' stato creato quando l'ho creato io
                    dtJob = DateSerial(vetJobs(Ix).Submitted.wYear, vetJobs(Ix).Submitted.wMonth, vetJobs(Ix).Submitted.wDay) + TimeSerial(vetJobs(Ix).Submitted.wHour, vetJobs(Ix).Submitted.wMinute, vetJobs(Ix).Submitted.wSecond)
                    If Abs(DateDiff("s", dtJob, dt)) = Secondi Then
                        InizioStampa = Printer.DeviceName & Chr$(0) & vetJobs(Ix).JobId
                        Exit For
                    End If
                End If
            Next Ix
            If Len(InizioStampa) Then Exit For
        Next Secondi
    End If
    Call ClosePrinter(hPrinter)

End Function

Function EliminaStampa(IdStampa$) As String
    Const JOB_CONTROL_CANCEL = 3
    Const JOB_CONTROL_PAUSE = 1
    Const JOB_CONTROL_RESTART = 4
    Const JOB_CONTROL_RESUME = 2
    
    Dim PD          As PRINTER_DEFAULTS
    Dim Ji          As JOB_INFO_1
    Dim vetJobs()   As JOB_INFO_1
    Dim Esito       As Long
    Dim hPrinter    As Long
    Dim NJob        As Long
    Dim richiesto   As Long
    Dim ritorno     As Long
    Dim Stampante   As Printer
    Dim Ix          As Integer
    Dim IDJob       As Long
    
    '
    ' Questa funzione elimina una stampa precedentemente creata
    '
    
    hPrinter = 0
    Esito = OpenPrinter(Left$(IdStampa$, InStr(IdStampa$, Chr$(0)) - 1), hPrinter, PD)
    If Esito = 0 Then Exit Function
    
    richiesto = 0
    ritorno = 0
    Esito = EnumJobs(hPrinter, 0, 1000, 1, Ji, richiesto, richiesto, ritorno)
    NJob = richiesto / Len(Ji)
    ReDim vetJobs(NJob)
    Esito = EnumJobs(hPrinter, 0, NJob, 1, vetJobs(0), Len(Ji) * NJob, richiesto, ritorno)
    
    If Esito Then
        IDJob = Val(Mid$(IdStampa$, InStr(IdStampa$, Chr$(0)) + 1))
        For Ix = 0 To ritorno - 1
            If vetJobs(Ix).JobId = IDJob Then
                ' elimino il documento
                Esito = SetJob(hPrinter, vetJobs(Ix).JobId, 1, vetJobs(Ix), JOB_CONTROL_CANCEL)
            End If
        Next Ix
    End If
    Call ClosePrinter(hPrinter)

End Function

Function GetStringa(Puntatore As Long)
    Dim Stringa As String
    Stringa = Space$(32767)
    Call lstrcpyn(Stringa, Puntatore, 32766)
    GetStringa = Left$(Stringa, InStr(Stringa, Chr$(0)) - 1)
End Function

