Attribute VB_Name = "EGBEMU"
Option Explicit
'COMPLETAMENTE FUNCIONAL

Private Declare Function getFrequency Lib "kernel32" _
      Alias "QueryPerformanceFrequency" ( _
      ByRef Frequency As Currency) _
      As Long
 
Private Declare Function getTime Lib "kernel32" _
      Alias "QueryPerformanceCounter" ( _
      ByRef Counter As Currency) _
      As Long

Public Declare Function GetAsyncKeyState Lib "user32" ( _
        ByVal vKey As Long) As Integer
        

Dim regA As Byte
Dim regF As Byte
Dim regB As Byte
Dim regC As Byte
Dim regD As Byte
Dim regE As Byte
Dim regH As Byte
Dim regL As Byte

Dim flagZ As Byte
Dim flagH As Byte
Dim flagN As Byte
Dim flagC As Byte

Dim flagIME As Boolean

Dim regPC As Long
Dim regSP As Long
Dim tick As Long
Dim tickNextInstruction As Long

Dim TimerControlClockSelectedValue As Integer
Dim TimerEnabled As Boolean


Global Cycles(256) As Byte
Global OpCodeBytes(256) As Byte
Global OpCodeNemonic(256) As String
Global IOPortsNames(76) As String

Dim RAM(65536) As Byte
Dim BIOS(256) As Byte

Dim screen(160, 144) As Byte
Dim spritesDirtyBuffer(263) As Long '255 + 8
Dim windowBuffer(300, 300) As Byte
'Dim tilesBuffer(0 To 383, 0 To 7, 0 To 7) As Byte
Dim tilesBuffer(0 To 1023, 0 To 7, 0 To 7) As Byte '32*32 tiles

Dim LCDEnabled As Boolean
Dim BGEnabled As Boolean
Dim windowEnabled As Boolean
Dim spriteEnabled As Boolean
Dim spritesSize As Boolean

Dim colorArrayrangeStr(3, 50) As String
Dim contadorPixelsDeCadaColor(4) As Long

Dim rgbColors(3) As Long
Dim BGP(3) As Byte
Dim OBP0(3) As Byte
Dim OBP1(3) As Byte

Dim BGDirty As Boolean
Dim windowDirty As Boolean
Dim spritesDirty As Boolean
Dim screenDirty As Boolean

Const regLY As Long = 65348 'FF44
Const regLYC As Long = 65349 'FF45
Const regIF As Long = 65295 'FF0F
Const regIE As Long = 65535 'FFFF
Const regLCDC As Long = 65344 'FF40
Const regSTAT As Long = 65345 'FF41
Const regDMA As Long = 65350 'FF46
Const regJOYPAD As Long = 65280 'FF00
Const regNR51 As Long = 65317 'FF25
Const regSB As Long = 65281 'FF01
Const regSC As Long = 65282 'FF02
Const regSCX As Long = 65347 'FF43
Const regSCY As Long = 65346 'FF42
Const regBGP As Long = 65351 'FF47
Const regOBP0 As Long = 65352 'FF48
Const regOBP1 As Long = 65353 'FF49
Const regBIOS As Long = 65360 'FF50
Const regDIV As Long = 65284 'FF04
Const regTIMA As Long = 65285 'FF05
Const regTMA As Long = 65286 'FF06
Const regTAC As Long = 65287 'FF07

Const dirROM As Long = 0 '0x0000
Const dirVRAM As Long = 32768 '0x8000
'Const dirTileMap As Long = 38912 '0x9800
Dim dirBGTileMap As Long '0x9800 o 0x9C00
Dim dirBGTileData As Long '0x8800 o 0x8000
Const dirRAMBank As Long = 40960 '0xA000
Const dirIntRAM As Long = 49152 '0xC000
Const dirEchoRAM As Long = 57344 '0xE000
Const dirOAM As Long = 65024 '0xFE00
Const dirIOPorts As Long = 65280 '0xFF00
Const dirHighRAM As Long = 65408 '0xFF80
Const dirIntEnReg As Long = 65535 '0xFFFF

Dim teclaPulsadaEscape As Boolean
Dim teclaPulsadaA As Boolean
Dim teclaPulsadaB As Boolean
Dim teclaPulsadaEnter As Boolean
Dim teclaPulsadaSelect As Boolean
Dim teclaPulsadaFlechaIzda As Boolean
Dim teclaPulsadaFlechaDcha As Boolean
Dim teclaPulsadaFlechaArriba As Boolean
Dim teclaPulsadaFlechaAbajo As Boolean

Dim FPS As Byte
Dim pixelsPaintedPerFPS As Long


Dim debug2Excel As Boolean
Dim debug2Screen As Boolean
Dim filaDebug As Long
Dim contadorInstrucciones(255) As Long


Sub EGBEMU()
    Dim i As Long
    
    debug2Excel = False
    debug2Screen = False
    
    'Funciones auxiliares
    Call createCyclesArray
    Call createOpCodeBytesArray
    Call createOpCodeNemonics
    Call createIOPortsNames

    Call DeclareKeys
    
    'Cargo la ROM en memoria
    Call loadROM
    'Cargo la BIOS
    'Call loadBIOS
    'For i = 0 To 255
    '    RAM(i) = BIOS(i)
    'Next i
    
    'Inicializo los registros y flags
    Call initializeRegisters
    
    rgbColors(0) = RGB(155, 188, 15)
    rgbColors(1) = RGB(139, 172, 15)
    rgbColors(2) = RGB(48, 98, 48)
    rgbColors(3) = RGB(15, 56, 15)
    BGP(0) = 0
    BGP(1) = 1
    BGP(2) = 2
    BGP(3) = 3
    OBP0(0) = 0
    OBP0(1) = 1
    OBP0(2) = 2
    OBP0(3) = 3
    OBP1(0) = 0
    OBP1(1) = 1
    OBP1(2) = 0
    OBP1(3) = 3
    
    'Pongo toda la pantalla en blanco (en verde)
    Call initScreen
    
    'Cabecera
    If debug2Excel = True Then
        filaDebug = 0
        Range("A1").Offset(0, 0).Value = "OPCODE;NEMONIC;NEM_HEX;A;B;C;DE;HL;SP"
'        Call printInitialDebugInfo
    End If
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayStatusBar = False
        .Calculation = xlCalculationManual
    End With
    
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
 
   Dim FPSstartTime As Currency
   Dim FPSendTime As Currency
   Dim FPSperSecond As Currency
   Dim FPStimeElapsed As Double
   Dim startNewFPSCounter As Boolean
   Dim instructionsPerSecond As Long
    
   getFrequency perSecond
   getTime startTime

    startNewFPSCounter = True
    getFrequency FPSperSecond

    'Comienzo la CPU
    Call CPU 'Primera instrucción
    
    'For tick = 1 To 40000000
'    Do While tick < 40000000 And teclaPulsadaEscape <> True
    Do While teclaPulsadaEscape <> True
        tick = tick + 1
        If startNewFPSCounter = True Then
            getTime FPSstartTime
            startNewFPSCounter = False
        End If
        
        If tick >= tickNextInstruction Then
            Call CPU
            instructionsPerSecond = instructionsPerSecond + 1
        End If
        Call PPU
        Call Timers
        
        getTime FPSendTime
        FPStimeElapsed = (FPSendTime - FPSstartTime) / FPSperSecond
        If FPStimeElapsed > 1 Then
            startNewFPSCounter = True
            Debug.Print "FPS: " & FPS & "*** Instructions: " & instructionsPerSecond & "*** Pixels: " & pixelsPaintedPerFPS
            FPS = 0
            pixelsPaintedPerFPS = 0
            instructionsPerSecond = 0
        End If
    Loop
    'Next tick
    
   getTime endTime
   timeElapsed = (endTime - startTime) / perSecond
   Debug.Print "Code took " & timeElapsed & " seconds to run"
 
    Call initScreen 'Limpio la pantalla antes de salir
 
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayStatusBar = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Private Sub CPU()
    Dim opCode As Integer
    
    Dim temp As Long
    Dim temp2 As Long
    Dim temp3 As Long
    
    Dim setInterruptsInNextCycle As Boolean
    setInterruptsInNextCycle = False
    
    opCode = RAM(regPC)
    
    'Para depuracion
    Dim opCodeHex As String
    Dim regPCHex As String
    opCodeHex = Hex(opCode)
    regPCHex = Hex(regPC)
    Dim stringTemp As String
    Dim regHLstr As String
    Dim regDEstr As String
    Static oldValue As Byte
    
    If debug2Excel Or debug2Screen Then
        If regD < 16 Then
            regDEstr = "0"
        End If
        regDEstr = regDEstr & Hex(regD)
        If regE < 16 Then
            regDEstr = regDEstr & "0"
        End If
        regDEstr = regDEstr & Hex(regE)
        If regH < 16 Then
            regHLstr = "0"
        End If
        regHLstr = regHLstr & Hex(regH)
        If regL < 16 Then
            regHLstr = regHLstr & "0"
        End If
        regHLstr = regHLstr & Hex(regL)
            
        stringTemp = regPCHex & ";" & OpCodeNemonic(opCode) & ";" & Hex(RAM(regPC)) & " " & Hex(RAM(regPC + 1)) & " " & Hex(RAM(regPC + 2)) & ";" & regA & ";" & regB & ";" & regC & ";" & regDEstr & ";" & regHLstr & ";" & regSP
        If debug2Screen Then
            Debug.Print (stringTemp)
        End If
        If debug2Excel Then 'And (tick > 15200000 And tick < 15300000) Then
            filaDebug = filaDebug + 1
            With Sheets("Hoja2")
                .Range("A1").Offset(filaDebug, 0).Value = stringTemp
            End With
        End If
    End If
    
    contadorInstrucciones(opCode) = contadorInstrucciones(opCode) + 1
    
    'Empieza el código
    Select Case opCode
        Case 0 'NOP
        Case 1 'LD BC, d16
            regB = RAM(regPC + 2)
            regC = RAM(regPC + 1)
        Case 2 'LD (BC), A
            RAM(regB * 256& + regC) = regA
            Call writeRAM(regB * 256& + regC, regA)
        Case 3 'INC BC
            temp = regB * 256& + regC
            temp = temp + 1
            'regB = Application.WorksheetFunction.Bitrshift(temp, 8)
            If temp > 65535 Then
                temp = 0
            End If
            regB = temp \ 256
            regC = temp And 255
        Case 4 'INC B
            flagH = 0
            flagZ = 0
            If (regB And 15) + 1 > 15 Then
                flagH = 1
            End If
            If regB = 255 Then
                regB = 0
                flagZ = 1
            Else
                regB = regB + 1
            End If
        Case 5 'DEC B
            flagZ = 0
            flagH = 0
            If (regB And 15) - 1 < 0 Then
                flagH = 1
            End If
            If regB = 0 Then
                regB = 255
            Else
                regB = regB - 1
                If regB = 0 Then
                    flagZ = 1
                End If
            End If
        Case 6 'LD B, d8
            regB = RAM(regPC + 1)
        Case 7 'RLCA
            flagZ = 0
            flagH = 0
            flagC = regA And 128 '128 = 80x -> Para coger el MSB
            'flagC = Application.WorksheetFunction.Bitrshift(flagC, 7)
            flagC = flagC \ 128
            'temp = Application.WorksheetFunction.Bitlshift(regA, 1)
            temp = regA * 2
            regA = temp And 255
            regA = regA + flagC
        Case 8 'LD (a16), SP
            RAM(regPC + 1) = regSP And 255
'            RAM(regPC + 2) = Application.WorksheetFunction.Bitrshift(regSP, 8) And 255
            RAM(regPC + 2) = (regSP \ 256) And 255
        Case 9 'ADD HL, BC
            flagC = 0
            flagH = 0
            temp = regH * 256& + regL
            temp2 = regB * 256& + regC
            temp3 = temp + temp2
            If (((temp And 4095) + (temp2 And 4095)) And 4096) > 0 Then
                flagH = 1
            End If
            If temp3 > 65535 Then
                flagC = 1
            End If
            regH = (temp3 \ 256) And 255
            regL = temp3 And 255
        Case 10 '0x0A LD A, (BC)
            regA = RAM(regB * 256& + regC)
        Case 11 '0x0B DEC BC
            temp = regB * 256& + regC
            temp = temp - 1
            regB = temp \ 256
            regC = temp And 255
        Case 12 '0x0C INC C
            flagH = 0
            flagZ = 0
            If (regC And 15) + 1 > 15 Then
                flagH = 1
            End If
            If regC = 255 Then
                regC = 0
                flagZ = 1
            Else
                regC = regC + 1
            End If
        Case 13 '0x0D DEC C
            flagZ = 0
            flagH = 0
            If (regC And 15) - 1 < 0 Then
                flagH = 1
            End If
            If regC = 0 Then
                regC = 255
            Else
                regC = regC - 1
                If regC = 0 Then
                    flagZ = 1
                End If
            End If
        Case 14 '0x0E LD C, d8
            regC = RAM(regPC + 1)
        Case 15 '0x0F RRCA
            flagZ = 0
            flagH = 0
            flagC = regA And 1 '128 = 80x -> Para coger el MSB
            regA = regA \ 2
            regA = regA + (flagC * 128&)
        Case 16 '0x10 STOP
            Debug.Print "STOP"
        Case 17 '0x11 LD DE, d16
            regD = RAM(regPC + 2)
            regE = RAM(regPC + 1)
        Case 18 '0x12 LD (DE), A
            RAM(regD * 256& + regE) = regA
            Call writeRAM(regD * 256& + regE, regA)
        Case 19 '0x13 INC DE
            temp = regD * 256& + regE
            temp = temp + 1
'            regD = Application.WorksheetFunction.Bitrshift(temp, 8) And 255
            regD = (temp \ 256) And 255
            regE = temp And 255
        Case 20 '0x14 INC D
            flagH = 0
            flagZ = 0
            If (regD And 15) + 1 > 15 Then
                flagH = 1
            End If
            If regD = 255 Then
                regD = 0
                flagZ = 1
            Else
                regD = regD + 1
            End If
        Case 21 '0x15 DEC D
            flagZ = 0
            flagH = 0
            If (regD And 15) - 1 < 0 Then
                flagH = 1
            End If
            If regD = 0 Then
                regD = 255
            Else
                regD = regD - 1
                If regD = 0 Then
                    flagZ = 1
                End If
            End If
        Case 22 '0x16 LD D, d8
            regD = RAM(regPC + 1)
        Case 23 '0x17 RLA
            flagZ = 0
            flagH = 0
            temp = regA * 2
            temp = temp + flagC
            regA = temp And 255
            flagC = (temp And 256) \ 256 'No estoy seguro si no se tiene que hacer antes de la suma anterior
        Case 24 '0x18 JR s8
            temp = RAM(regPC + 1)
            If temp > 127 Then 'Es un número negativo, hay que restar
                temp = 256 - temp
                regPC = regPC - temp
            Else
                regPC = regPC + temp
            End If
        Case 25 '0x19 ADD HL, DE
            flagC = 0
            flagH = 0
            temp = regH * 256& + regL
            temp2 = regD * 256& + regE
            temp3 = temp + temp2
            If (((temp And 4095) + (temp2 And 4095)) And 4096) > 0 Then
                flagH = 1
            End If
            If temp3 > 65535 Then
                flagC = 1
            End If
            regH = (temp3 \ 256) And 255
            regL = temp3 And 255
        Case 26 '0x1A LD A, (DE)
            regA = RAM(regD * 256& + regE)
        Case 27 '0x1B DEC DE
            temp = regD * 256& + regE
            temp = temp - 1
            regD = temp \ 256
            regE = temp And 255
        Case 28 '0x1C INC E
            flagH = 0
            flagZ = 0
            If (regE And 15) + 1 > 15 Then
                flagH = 1
            End If
            If regE = 255 Then
                regE = 0
                flagZ = 1
            Else
                regE = regE + 1
            End If
        Case 29 '0x1D DEC E
            flagZ = 0
            flagH = 0
            If (regE And 15) - 1 < 0 Then
                flagH = 1
            End If
            If regE = 0 Then
                regE = 255
            Else
                regE = regE - 1
                If regE = 0 Then
                    flagZ = 1
                End If
            End If
        Case 30 '0x1E LD E, d8
            regE = RAM(regPC + 1)
        Case 31 '0x1F RRA
            flagZ = 0
            flagH = 0
            temp = regA And 1
            regA = regA \ 2
            regA = regA Or (flagC * 128)
            flagC = temp
        Case 32 '0x20 JR NZ, s8
            If flagZ = 0 Then
                temp = RAM(regPC + 1)
                If temp > 127 Then 'Es un número negativo, hay que restar
                    temp = 256 - temp
                    regPC = regPC - temp
                Else
                    regPC = regPC + temp
                End If
            Else
                flagZ = 1 'Lo pongo para poder poner un breakpoint aqui
            End If
        Case 33 '0x21 LD HL, d16
            regH = RAM(regPC + 2)
            regL = RAM(regPC + 1)
        Case 34 '0x22 LD (HL+), A
            temp = regH * 256& + regL
            RAM(temp) = regA
            Call writeRAM(temp, regA)
            temp = temp + 1
            regH = temp \ 256
            regL = temp And 255
        Case 35 '0x23 INC HL
            temp = regH * 256& + regL
            temp = temp + 1
            regH = temp \ 256
            regL = temp And 255
        Case 36 '0x24 INC H
            flagH = 0
            flagZ = 0
            If (regH And 15) + 1 > 15 Then
                flagH = 1
            End If
            If regH = 255 Then
                regH = 0
                flagZ = 1
            Else
                regH = regH + 1
            End If
        Case 37 '0x25 DEC H
            flagZ = 0
            flagH = 0
            If (regH And 15) - 1 < 0 Then
                flagH = 1
            End If
            If regH = 0 Then
                regH = 255
            Else
                regH = regH - 1
                If regH = 0 Then
                    flagZ = 1
                End If
            End If
        Case 38 '0X26 LD H, d8
            regH = RAM(regPC + 1)
        Case 39 '0x27 DAA
            'Debug.Print "OPERACION DAA"
            Dim correction As Byte
            
            If flagH = 1 Or ((flagN = 0) And ((regA And 15) > 9)) Then
                correction = correction Or 6
            End If
            
            If flagC = 1 Or ((flagN = 0) And (regA > 153)) Then
                correction = correction Or 96
                flagC = 1
            End If
            
            If flagN = 1 Then
                regA = regA - correction
            Else
                temp = CLng(regA) + correction
                If temp > 255 Then
                    regA = temp - 255
                Else
                    regA = regA + correction
                End If
            End If
            regA = regA And 255
            
            If regA = 0 Then
                flagZ = 1
            Else
                flagZ = 0
            End If
            flagH = 0
        Case 40 '0x28 JR Z, s8
            If flagZ = 1 Then
                temp = RAM(regPC + 1)
                If temp > 127 Then 'Es un número negativo, hay que restar
                    temp = 256 - temp
                    regPC = regPC - temp
                Else
                    regPC = regPC + temp
                End If
'            Else
'                flagZ = 0 'Lo pongo para poder poner un breakpoint aqui
            End If
        Case 41 '0x29 ADD HL, HL
            flagC = 0
            flagH = 0
            temp = regH * 256& + regL
            temp2 = regH * 256& + regL
            temp3 = temp + temp2
            If (((temp And 4095) + (temp2 And 4095)) And 4096) > 0 Then
                flagH = 1
            End If
            If temp3 > 65535 Then
                flagC = 1
            End If
            regH = (temp3 \ 256) And 255
            regL = temp3 And 255
        Case 42 '0x2A LD A, (HL+)
            temp = regH * 256& + regL
            regA = RAM(temp)
            temp = temp + 1
'            regH = Application.WorksheetFunction.Bitrshift(temp, 8)
            regH = temp \ 256
            regL = temp And 255
        Case 43 '0x2B DEC HL
            temp = regH * 256& + regL
            temp = temp - 1
            regH = temp \ 256
            regL = temp And 255
        Case 44 '0x2C INC L
            flagH = 0
            flagZ = 0
            If (regL And 15) + 1 > 15 Then
                flagH = 1
            End If
            If regL = 255 Then
                regL = 0
                flagZ = 1
            Else
                regL = regL + 1
            End If
        Case 45 '0x2D DEC L
            flagZ = 0
            flagH = 0
            If (regL And 15) - 1 < 0 Then
                flagH = 1
            End If
            If regL = 0 Then
                regL = 255
            Else
                regL = regL - 1
                If regL = 0 Then
                    flagZ = 1
                End If
            End If
        Case 46 '0x2E LD L, d8
            regL = RAM(regPC + 1)
        Case 47 '0x2F CPL
            regA = regA Xor 255
            flagH = 1
        Case 48 '0x30 JR NC, s8
            If flagC = 0 Then
                temp = RAM(regPC + 1)
                If temp > 127 Then 'Es un número negativo, hay que restar
                    temp = 256 - temp
                    regPC = regPC - temp
                Else
                    regPC = regPC + temp
                End If
'            Else
'                flagC = 1 'Lo pongo para poder poner un breakpoint aqui
            End If
        Case 49 '0x31 LD SP, d16
            regSP = (RAM(regPC + 2) * 256&) + RAM(regPC + 1)
        Case 50 '0x32 LD (HL-),A
            temp = regH * 256& + regL
            RAM(temp) = regA
            Call writeRAM(temp, regA)
            'Ahora decremento HL
            temp = temp - 1
'            regH = Application.WorksheetFunction.Bitrshift(temp, 8)
            regH = temp \ 256
            regL = temp And 255
        Case 51 '0x33 INC SP
            regSP = regSP + 1
        Case 52 '0x34 INC (HL)
            flagZ = 0
            flagH = 0
            temp = RAM(regH * 256& + regL)
            If temp + 1 > 15 Then
                flagH = 1
            End If
            temp = temp + 1
            If temp > 255 Then
                temp = 0
                flagZ = 1
            End If
            RAM(regH * 256& + regL) = temp
            Call writeRAM(regH * 256& + regL, CByte(temp))
        Case 53 '0x35 DEC (HL)
            flagZ = 0
            flagH = 0
            temp = RAM(regH * 256& + regL)
            If (temp And 15) - 1 < 0 Then
                flagH = 1
            End If
            If temp = 0 Then
                temp = 255
            Else
                temp = temp - 1
            End If
            RAM(regH * 256& + regL) = temp
            Call writeRAM(regH * 256& + regL, CByte(temp))
            If temp = 0 Then
                flagZ = 1
            End If
        Case 54 '0x36 LD (HL), d8
            RAM(regH * 256& + regL) = RAM(regPC + 1)
            Call writeRAM(regH * 256& + regL, RAM(regPC + 1))
        Case 55 '0x37 SCF
            flagH = 0
            flagC = 1
        Case 56 '0x38 JR C, s8
            If flagC = 1 Then
                temp = RAM(regPC + 1)
                If temp > 127 Then 'Es un número negativo, hay que restar
                    temp = 256 - temp
                    regPC = regPC - temp
                Else
                    regPC = regPC + temp
                End If
'            Else
'                flagC = 0 'Lo pongo para poder poner un breakpoint aqui
            End If
        Case 57 '0x39 ADD HL, SP
            flagC = 0
            flagH = 0
            temp = regH * 256& + regL
            temp2 = regSP
            temp3 = temp + temp2
            If (((temp And 4095) + (temp2 And 4095)) And 4096) > 0 Then
                flagH = 1
            End If
            If temp3 > 65535 Then
                flagC = 1
            End If
            regH = (temp3 \ 256) And 255
            regL = temp3 And 255
        Case 58 '0x3A LD A, (HL-)
            temp = regH * 256& + regL
            regA = RAM(temp)
            'Ahora decremento HL
            If temp = 0 Then
                temp = 65535
            Else
                temp = temp - 1
            End If
'            regH = Application.WorksheetFunction.Bitrshift(temp, 8)
            regH = temp \ 256
            regL = temp And 255
        Case 59 '0x3B DEC SP
            If regSP = 0 Then
                regSP = 65535
            Else
                regSP = regSP - 1
            End If
        Case 60 '0x3C INC A
            flagH = 0
            flagZ = 0
            If (regA And 15) + 1 > 15 Then
                flagH = 1
            End If
            If regA = 255 Then
                regA = 0
                flagZ = 1
            Else
                regA = regA + 1
            End If
        Case 61 '0x3D DEC A
            flagZ = 0
            flagH = 0
            If (regA And 15) - 1 < 0 Then
                flagH = 1
            End If
            If regA = 0 Then
                regA = 255
            Else
                regA = regA - 1
                If regA = 0 Then
                    flagZ = 1
                End If
            End If
        Case 62 '0x3E LD A, d8
            regA = RAM(regPC + 1)
        Case 63 '0x3F CCF
            flagH = 0
            If flagC = 0 Then
                flagC = 1
            Else
                flagC = 0
            End If
        Case 64 '0x40 LD B, B
            'regB = regB 'Lo quito hasta que no descubra para que sirve
        Case 65 '0x41 LD B, C
            regB = regC
        Case 66 '0x42 LD B, D
            regB = regD
        Case 67 '0x43 LD B, E
            regB = regE
        Case 68 '0x44 LD B, H
            regB = regH
        Case 69 '0x45 LD B, L
            regB = regL
        Case 70 '0x46 LD B, (HL)
            regB = RAM((regH * 256&) + regL)
        Case 71 '0x47 LD B, A
            regB = regA
        Case 72 '0x48 LD C, B
            regC = regB
        Case 73 '0x49 LD C, C
'            regC = regC
        Case 74 '0x4A LD C, D
            regC = regD
        Case 75 '0x4B LD C, E
            regC = regE
        Case 76 '0x4C LD C, H
            regC = regH
        Case 77 '0x4D LD C, L
            regC = regL
        Case 78 '0x4E LD C, (HL)
            regC = RAM((regH * 256&) + regL)
        Case 79 '0x4F LD C, A
            regC = regA
        Case 80 '0x50 LD D, B
            regD = regB
        Case 81 '0x51 LD D, C
            regD = regC
        Case 82 '0x52 LD D, D
'            regD = regD
        Case 83 '0x53 LD D, E
            regD = regE
        Case 84 '0x54 LD D, H
            regD = regH
        Case 85 '0x55 LD D, L
            regD = regL
        Case 86 '0x56 LD D, (HL)
            regD = RAM((regH * 256&) + regL)
        Case 87 '0x57 LD D, A
            regD = regA
        Case 88 '0x58 LD E, B
            regE = regB
        Case 89 '0x59 LD E, C
            regE = regC
        Case 90 '0x5A LD E, D
            regE = regD
        Case 91 '0x5B LD E, E
 '           regE = regE
        Case 92 '0x5C LD E, H
            regE = regH
        Case 93 '0x5D LD E, L
            regE = regL
        Case 94 '0x5E LD E, (HL)
            regE = RAM((regH * 256&) + regL)
        Case 95 '0x5F LD E, A
            regE = regA
        Case 96 '0x60 LD H, B
            regH = regB
        Case 97 '0x61 LD H, C
            regH = regC
        Case 98 '0x62 LD H, D
            regH = regD
        Case 99 '0x63 LD H, E
            regH = regE
        Case 100 '0x64 LD H, H
'            regH = regH
        Case 101 '0x65 LD H, L
            regH = regL
        Case 102 '0x66 LD H, (HL)
            regH = RAM((regH * 256&) + regL)
        Case 103 '0x67 LD H, A
            regH = regA
        Case 104 '0x68 LD L, B
            regL = regB
        Case 105 '0x69 LD L, C
            regL = regC
        Case 106 '0x6A LD L, D
            regL = regD
        Case 107 '0x6B LD L, E
            regL = regE
        Case 108 '0x6C LD L, H
            regL = regH
        Case 109 '0x6D LD L, L
'            regL = regL
        Case 110 '0x6E LD L, (HL)
            regL = RAM((regH * 256&) + regL)
        Case 111 '0x6F LD L, A
            regL = regA
        Case 112 '0x70 LD (HL), B
            RAM(regH * 256& + regL) = regB
            Call writeRAM(regH * 256& + regL, regB)
        Case 113 '0x71 LD (HL), C
            RAM(regH * 256& + regL) = regC
            Call writeRAM(regH * 256& + regL, regC)
        Case 114 '0x72 LD (HL), D
            RAM(regH * 256& + regL) = regD
            Call writeRAM(regH * 256& + regL, regD)
        Case 115 '0x73 LD (HL), E
            RAM(regH * 256& + regL) = regE
            Call writeRAM(regH * 256& + regL, regE)
        Case 116 '0x74 LD (HL), H
            RAM(regH * 256& + regL) = regH
            Call writeRAM(regH * 256& + regL, regH)
        Case 117 '0x75 LD (HL), L
            RAM(regH * 256& + regL) = regL
            Call writeRAM(regH * 256& + regL, regL)
        Case 118 '0x76 HALT
'            Debug.Print "HALT"
'            Debug.Assert (1)
'            regPC = regPC - OpCodeBytes(opCode) 'Para quedarme en el halt -> no funciona bien, revisar
        Case 119 '0x77 LD (HL), A
            RAM(regH * 256& + regL) = regA
            Call writeRAM(regH * 256& + regL, regA)
        Case 120 '0x78 LD A, B
            regA = regB
        Case 121 '0x79 LD A, C
            regA = regC
        Case 122 '0x7A LD A, D
            regA = regD
        Case 123 '0x7B LD A, E
            regA = regE
        Case 124 '0x7C LD A, H
            regA = regH
        Case 125 '0x7D LD A, L
            regA = regL
        Case 126 '0x7E LD A, (HL)
            regA = RAM(regH * 256& + regL)
        Case 127 '0x7F LD A, A
'            regA = regA
        Case 128 '0x80 ADD A, B
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) + CInt(regB)
            regA = temp And 255
            If temp > 255 Then
                flagC = 1
            End If
            If regA = 0 Then
                flagZ = 1
            End If
            If ((regA And 15) + (regB And 15)) > 15 Then
                flagH = 1
            End If
        Case 129 '0x81 ADD A, C
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) + CInt(regC)
            regA = temp And 255
            If temp > 255 Then
                flagC = 1
            End If
            If regA = 0 Then
                flagZ = 1
            End If
            If ((regA And 15) + (regC And 15)) > 15 Then
                flagH = 1
            End If
        Case 130 '0x82 ADD A, D
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) + CInt(regD)
            regA = temp And 255
            If temp > 255 Then
                flagC = 1
            End If
            If regA = 0 Then
                flagZ = 1
            End If
            If ((regA And 15) + (regD And 15)) > 15 Then
                flagH = 1
            End If
        Case 131 '0x83 ADD A, E
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) + CInt(regE)
            regA = temp And 255
            If temp > 255 Then
                flagC = 1
            End If
            If regA = 0 Then
                flagZ = 1
            End If
            If ((regA And 15) + (regE And 15)) > 15 Then
                flagH = 1
            End If
        Case 132 '0x84 ADD A, H
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) + CInt(regH)
            regA = temp And 255
            If temp > 255 Then
                flagC = 1
            End If
            If regA = 0 Then
                flagZ = 1
            End If
            If ((regA And 15) + (regH And 15)) > 15 Then
                flagH = 1
            End If
        Case 133 '0x85 ADD A, L
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) + CInt(regL)
            regA = temp And 255
            If temp > 255 Then
                flagC = 1
            End If
            If regA = 0 Then
                flagZ = 1
            End If
            If ((regA And 15) + (regL And 15)) > 15 Then
                flagH = 1
            End If
        Case 134 '0X86 ADD A, (HL)
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) + CInt(RAM(regH * 256& + regL))
            regA = temp And 255
            If temp > 255 Then
                flagC = 1
            End If
            If regA = 0 Then
                flagZ = 1
            End If
            If ((regA And 15) + (RAM(regH * 256& + regL) And 15)) > 15 Then
                flagH = 1
            End If
        Case 135 '0x87 ADD A, A
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) + CInt(regA)
            regA = temp And 255
            If temp > 255 Then
                flagC = 1
            End If
            If regA = 0 Then
                flagZ = 1
            End If
            If ((regA And 15) + (regA And 15)) > 15 Then
                flagH = 1
            End If
        Case 136 '0x88 ADC A, B
            flagZ = 0
            flagH = 0
            temp = CInt(regA) + CInt(regB) + flagC
            flagC = 0
            If temp > 255 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regB Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 137 '0x89 ADC A, C
            flagZ = 0
            flagH = 0
            temp = CInt(regA) + CInt(regC) + flagC
            flagC = 0
            If temp > 255 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regC Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 138 '0x8A ADC A, D
            flagZ = 0
            flagH = 0
            temp = CInt(regA) + CInt(regD) + flagC
            flagC = 0
            If temp > 255 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regD Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 139 '0x8B ADC A, E
            flagZ = 0
            flagH = 0
            temp = CInt(regA) + CInt(regE) + flagC
            flagC = 0
            If temp > 255 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regE Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 140 '0x8C ADC A, H
            flagZ = 0
            flagH = 0
            temp = CInt(regA) + CInt(regH) + flagC
            flagC = 0
            If temp > 255 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regH Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 141 '0x8D ADC A, L
            flagZ = 0
            flagH = 0
            temp = CInt(regA) + CInt(regL) + flagC
            flagC = 0
            If temp > 255 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regL Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 142 '0x8E ADC A, (HL)
            flagZ = 0
            flagH = 0
            temp = CInt(regA) + CInt(RAM(regH * 256& + regL)) + flagC
            flagC = 0
            If temp > 255 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor RAM(regH * 256& + regL) Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 143 '0x8F ADC A, A
            flagZ = 0
            flagH = 0
            temp = CInt(regA) + CInt(regA) + flagC
            flagC = 0
            If temp > 255 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regA Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 144 '0x90 SUB B
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regB)
            If temp < 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regB Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 145 '0x91 SUB C
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regC)
            If temp < 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regC Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 146 '0x92 SUB D
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regD)
            If temp < 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regD Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 147 '0x93 SUB E
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regE)
            If temp < 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regE Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 148 '0x94 SUB H
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regH)
            If temp < 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regH Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 149 '0x95 SUB L
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regL)
            If temp < 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regL Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 150 '0x96 SUB (HL)
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(RAM(regH * 256& + regL))
            If temp < 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor RAM(regH * 256& + regL) Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 151 '0x97 SUB A
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regA)
            If temp < 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regA Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 152 '0x98 SBC A, B
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regB) - flagC
            If temp And 256 > 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regB Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 153 '0x99 SBC A, C
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regC) - flagC
            If temp And 256 > 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regC Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 154 '0x9A SBC A, D
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regD) - flagC
            If temp And 256 > 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regD Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 155 '0x9B SBC A, E
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regE) - flagC
            If temp And 256 > 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regE Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 156 '0x9C SBC A, H
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regH) - flagC
            If temp And 256 > 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regH Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 157 '0x9D SBC A, L
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regL) - flagC
            If temp And 256 > 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regL Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 158 '0x9E SBC A, HL
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(RAM(regH * 256& + regL)) - flagC
            If temp And 256 > 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor RAM(regH * 256& + regL) Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 159 '0x9F SBC A, A
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - CInt(regA) - flagC
            If temp And 256 > 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor regA Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 160 '0xA0 AND B
            flagZ = 0
            flagC = 0
            flagH = 1
            regA = regA And regB
            If regA = 0 Then
                flagZ = 1
            End If
        Case 161 '0xA1 AND C
            flagZ = 0
            flagC = 0
            flagH = 1
            regA = regA And regC
            If regA = 0 Then
                flagZ = 1
            End If
        Case 162 '0xA2 AND D
            flagZ = 0
            flagC = 0
            flagH = 1
            regA = regA And regD
            If regA = 0 Then
                flagZ = 1
            End If
        Case 163 '0xA3 AND E
            flagZ = 0
            flagC = 0
            flagH = 1
            regA = regA And regE
            If regA = 0 Then
                flagZ = 1
            End If
        Case 164 '0xA4 AND H
            flagZ = 0
            flagC = 0
            flagH = 1
            regA = regA And regH
            If regA = 0 Then
                flagZ = 1
            End If
        Case 165 '0xA5 AND L
            flagZ = 0
            flagC = 0
            flagH = 1
            regA = regA And regL
            If regA = 0 Then
                flagZ = 1
            End If
        Case 166 '0xA6 AND (HL)
            flagZ = 0
            flagC = 0
            flagH = 1
            regA = regA And RAM(regH * 256& + regL)
            If regA = 0 Then
                flagZ = 1
            End If
        Case 167 '0xA7 AND A
            flagZ = 0
            flagC = 0
            flagH = 1
            'regA = regA And regA 'Lo quito hasta que no descubra para que hay que hacer esto
            If regA = 0 Then
                flagZ = 1
            End If
        Case 168 '0xA8 XOR B
            flagH = 0
            flagC = 0
            flagZ = 0
            regA = regA Xor regB
            If regA = 0 Then
                flagZ = 1
            End If
        Case 169 '0xA9 XOR C
            flagH = 0
            flagC = 0
            flagZ = 0
            regA = regA Xor regC
            If regA = 0 Then
                flagZ = 1
            End If
        Case 170 '0xAA XOR D
            flagH = 0
            flagC = 0
            flagZ = 0
            regA = regA Xor regD
            If regA = 0 Then
                flagZ = 1
            End If
        Case 171 '0xAB XOR E
            flagH = 0
            flagC = 0
            flagZ = 0
            regA = regA Xor regE
            If regA = 0 Then
                flagZ = 1
            End If
        Case 172 '0xAC XOR H
            flagH = 0
            flagC = 0
            flagZ = 0
            regA = regA Xor regH
            If regA = 0 Then
                flagZ = 1
            End If
        Case 173 '0xAD XOR L
            flagZ = 0
            flagH = 0
            flagC = 0
            regA = regA Xor regL
            If regA = 0 Then
                flagZ = 1
            End If
        Case 174 '0xAE XOR (HL)
            flagZ = 0
            flagC = 0
            flagH = 0
            regA = regA Xor RAM(regH * 256& + regL)
        Case 175 '0xAF XOR A
            flagZ = 1
            flagC = 0
            flagH = 0
            regA = 0
        Case 176 '0xB0 OR B
            flagC = 0
            flagH = 0
            flagZ = 0
            regA = regB Or regA
            If regA = 0 Then
                flagZ = 1
            End If
        Case 177 '0xB1 OR C
            flagC = 0
            flagH = 0
            flagZ = 0
            regA = regC Or regA
            If regA = 0 Then
                flagZ = 1
            End If
        Case 178 '0xB2 OR D
            flagC = 0
            flagH = 0
            flagZ = 0
            regA = regD Or regA
            If regA = 0 Then
                flagZ = 1
            End If
        Case 179 '0xB3 OR E
            flagC = 0
            flagH = 0
            flagZ = 0
            regA = regE Or regA
            If regA = 0 Then
                flagZ = 1
            End If
        Case 180 '0xB4 OR H
            flagC = 0
            flagH = 0
            flagZ = 0
            regA = regH Or regA
            If regA = 0 Then
                flagZ = 1
            End If
        Case 181 '0xB5 OR L
            flagC = 0
            flagH = 0
            flagZ = 0
            regA = regL Or regA
            If regA = 0 Then
                flagZ = 1
            End If
        Case 182 '0xB6 OR (HL)
            flagC = 0
            flagH = 0
            flagZ = 0
            regA = regA Or RAM(regH * 256& + regL)
            If regA = 0 Then
                flagZ = 1
            End If
        Case 183 '0xB7 OR A
            flagC = 0
            flagH = 0
            flagZ = 0
            regA = regA Or regA
            If regA = 0 Then
                flagZ = 1
            End If
        Case 184 '0xB8 CP B
            flagZ = 0
            flagC = 0
            flagH = 0
            If regA = regB Then
                flagZ = 1
            End If
            If regA < regB Then
                flagC = 1
            End If
            If (regA And 15) < (regB And 15) Then
                flagH = 1
            End If
        Case 185 '0xB9 CP C
            flagZ = 0
            flagC = 0
            flagH = 0
            If regA = regC Then
                flagZ = 1
            End If
            If regA < regC Then
                flagC = 1
            End If
            If (regA And 15) < (regC And 15) Then
                flagH = 1
            End If
        Case 186 '0xBA CP D
            flagZ = 0
            flagC = 0
            flagH = 0
            If regA = regD Then
                flagZ = 1
            End If
            If regA < regD Then
                flagC = 1
            End If
            If (regA And 15) < (regD And 15) Then
                flagH = 1
            End If
        Case 187 '0xBB CP E
            flagZ = 0
            flagC = 0
            flagH = 0
            If regA = regE Then
                flagZ = 1
            End If
            If regA < regE Then
                flagC = 1
            End If
            If (regA And 15) < (regE And 15) Then
                flagH = 1
            End If
        Case 188 '0xBC CP H
            flagZ = 0
            flagC = 0
            flagH = 0
            If regA = regH Then
                flagZ = 1
            End If
            If regA < regH Then
                flagC = 1
            End If
            If (regA And 15) < (regH And 15) Then
                flagH = 1
            End If
        Case 189 '0xBD CP L
            flagZ = 0
            flagC = 0
            flagH = 0
            If regA = regL Then
                flagZ = 1
            End If
            If regA < regL Then
                flagC = 1
            End If
            If (regA And 15) < (regL And 15) Then
                flagH = 1
            End If
        Case 190 '0xBE CP (HL)
            flagZ = 0
            flagC = 0
            flagH = 0
            If regA = RAM(regH * 256& + regL) Then
                flagZ = 1
            End If
            If regA < RAM(regH * 256& + regL) Then
                flagC = 1
            End If
            If (regA And 15) < (RAM(regH * 256& + regL) And 15) Then
                flagH = 1
            End If
        Case 191 '0xBF CP A
            flagZ = 0
            flagC = 0
            flagH = 0
            If regA = regA Then
                flagZ = 1
            End If
            If regA < regA Then '¿?
                flagC = 1
            End If
            If (regA And 15) < (regA And 15) Then '¿?
                flagH = 1
            End If
        Case 192 '0xC0 RET NZ
            If flagZ = 0 Then
                temp = regSP \ 256
                temp2 = regSP And 255
                regPC = (RAM((temp * 256&) + temp2) * 256&) + RAM((temp * 256&) + temp2 - 1) - OpCodeBytes(opCode)
                regSP = regSP + 2
            End If
        Case 193 '0xC1 POP BC
            regB = RAM(regSP)
            regC = RAM(regSP - 1)
            regSP = regSP + 2
        Case 194 '0xC2 JP NZ, a16
            If flagZ = 0 Then
                regPC = (RAM(regPC + 2) * 256&) + RAM(regPC + 1) - OpCodeBytes(opCode)
            End If
        Case 195 '0xC3 JP a16
            regPC = (RAM(regPC + 2) * 256&) + RAM(regPC + 1) - OpCodeBytes(opCode) 'Le resto lo que luego voy a sumar al PC
        Case 196 '0xC4 CALL NZ, a16
            If flagZ = 0 Then
                regSP = regSP - 2
                temp = regPC + OpCodeBytes(opCode)
                RAM(regSP) = temp \ 256
                RAM(regSP - 1) = temp And 255
                regPC = (RAM(regPC + 2) * 256&) + RAM(regPC + 1) - OpCodeBytes(opCode)
            End If
        Case 197 '0xC5 PUSH BC
            regSP = regSP - 2
            RAM(regSP) = regB
            RAM(regSP - 1) = regC
        Case 198 '0xC6 ADD A, d8
            flagZ = 0
            flagH = 0
            flagC = 0
            temp = CInt(regA) + CInt(RAM(regPC + 1))
            If temp > 255 Then
                flagC = 1
                regA = temp - 256
            Else
                regA = temp
            End If
            If regA = 0 Then
                flagZ = 1
            End If
            If regA And 15 Then
                flagH = 1
            End If
        Case 199 '0xC7 RST 0
            regSP = regSP - 2
            temp = regPC + OpCodeBytes(opCode)
'            RAM(regSP) = Application.WorksheetFunction.Bitrshift(temp, 8)
            RAM(regSP) = temp \ 256
            RAM(regSP - 1) = temp And 255
            regPC = 0 - OpCodeBytes(opCode) 'OJO con esto....
        Case 200 '0xC8 RET Z
            If flagZ = 1 Then
                temp = regSP \ 256
                temp2 = regSP And 255
                regPC = (RAM((temp * 256&) + temp2) * 256&) + RAM((temp * 256&) + temp2 - 1) - OpCodeBytes(opCode)
                regSP = regSP + 2
            End If
        Case 201 '0xC9 RET
            temp = regSP \ 256
            temp2 = regSP And 255
            regPC = (RAM((temp * 256&) + temp2) * 256&) + RAM((temp * 256&) + temp2 - 1) - OpCodeBytes(opCode)
            regSP = regSP + 2
        Case 202 '0xCA JP Z, A16
            If flagZ = 1 Then
                regPC = (RAM(regPC + 2) * 256&) + RAM(regPC + 1) - OpCodeBytes(opCode)
            End If
        Case 203 'CB 'INSTRUCCION ESPECIAL, EL OPCODE A CONTINUACION
            Dim opCode2 As Integer
            opCode2 = RAM(regPC + 1)
            Select Case opCode2
                Case 17 '0x11 RL C
                    flagZ = 0
                    flagH = 0
                    temp = regC * 2
                    temp = temp + flagC
                    regC = temp And 255
                    flagC = (temp And 256) \ 256
                    If regC = 0 Then
                        flagZ = 1
                    End If
                Case 25 '0x19 RR C
                    flagZ = 0
                    flagH = 0
                    regC = regC \ 2
                    regC = regC And 127
                    If flagC = 1 Then
                        regC = regC Or 128
                    End If
                    If regC = 0 Then
                        flagZ = 1
                    End If
                Case 26 '0x20 RR D
                    flagZ = 0
                    flagH = 0
                    regD = regD \ 2
                    regD = regD And 127
                    If flagC = 1 Then
                        regD = regD Or 128
                    End If
                    If regD = 0 Then
                        flagZ = 1
                    End If
                Case 27 '0x21 RR D
                    flagZ = 0
                    flagH = 0
                    regD = regD \ 2
                    regD = regD And 127
                    If flagC = 1 Then
                        regD = regD Or 128
                    End If
                    If regD = 0 Then
                        flagZ = 1
                    End If
                Case 39 '0x27 SLA A
                    flagZ = 0
'                    flagC = Application.WorksheetFunction.Bitrshift(regA, 7) And 1
                    flagC = (regA \ 128) And 1
'                    temp = Application.WorksheetFunction.Bitlshift(regA, 1)
                    temp = regA * 2
                    regA = temp And 254
                Case 51 '0x33 SWAP E
                    flagC = 0
                    flagH = 0
                    flagZ = 0
                    temp = regE And 15 'Parte baja
                    temp2 = (regE \ 16) And 15
                    temp3 = (temp * 16&) + temp2
                    regE = CByte(temp3)
                    If regE = 0 Then
                        flagZ = 1
                    End If
                Case 55 '0x37 SWAP A
                    flagC = 0
                    flagH = 0
                    flagZ = 0
                    temp = regA And 15 'Parte baja
'                    temp2 = Application.WorksheetFunction.Bitrshift(regA, 4) And 15
                    temp2 = (regA \ 16) And 15
                    temp3 = (temp * 16&) + temp2
                    regA = CByte(temp3)
                    If regA = 0 Then
                        flagZ = 1
                    End If
                Case 56 '0x38 SRL B
                    flagZ = 0
                    flagH = 0
                    flagC = regB And 1
'                    regB = Application.WorksheetFunction.Bitrshift(regA, 1)
                    regB = regB \ 2
                    regB = regB And 127
                    If regB = 0 Then
                        flagZ = 1
                    End If
                Case 63 '0x3F SRL A
                    flagZ = 0
                    flagH = 0
                    flagC = regA And 1
'                    regA = Application.WorksheetFunction.Bitrshift(regA, 1)
                    regA = regA \ 2
                    regA = regA And 127
                    If regA = 0 Then
                        flagZ = 1
                    End If
                Case 64 '0x40 BIT 0, B
                    temp = regB
                    temp = temp And 1
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 65 '0x41 BIT 0, C
                    temp = regC
                    temp = temp And 1
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 70 '0x46 BIT 0, (HL)
                    temp = RAM(regH * 256& + regL)
                    temp = temp And 1
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 71 '0x47 BIT 0, A
                    temp = regA
                    temp = temp And 1
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 72 '0x48 BIT 1, B
                    temp = regB
                    temp = temp And 2
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 80 '0x50 BIT 2, B
                    temp = regB
                    temp = temp And 4
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 87 '0x57 BIT 2, A
                    temp = regA
                    temp = temp And 4
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 88 '0x58 BIT 3, B
                    temp = regB
                    temp = temp And 8
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 95 '0x5F BIT 3, A
                    temp = regA
                    temp = temp And 8
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 96 '0x60 BIT 4, B
                    temp = regB
                    temp = temp And 16
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 97 '0x61 BIT 4, C
                    temp = regC
                    temp = temp And 16
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 104 '0x68 BIT 5, B
                    temp = regB
                    temp = temp And 32
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 105 '0x69 BIT 5, C
                    temp = regC
                    temp = temp And 32
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 111 '0x6F BIT 5, A
                    temp = regA
                    temp = temp And 32
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 112 '0x70 BIT 6, B
                    temp = regB
                    temp = temp And 64
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 113 '0x71 BIT 6, C
                    temp = regC
                    temp = temp And 64
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 114 '0x72 BIT 6, D
                    temp = regD
                    temp = temp And 64
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 115 '0x73 BIT 6, E
                    temp = regE
                    temp = temp And 64
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 116 '0x74 BIT 6, H
                    temp = regH
                    temp = temp And 64
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 117 '0x75 BIT 6, L
                    temp = regL
                    temp = temp And 64
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 118 '0x76 BIT 6, (HL)
                    temp = RAM(regH * 256& + regL)
                    temp = temp And 64
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 119 '0x77 BIT 6, A
                    temp = regA
                    temp = temp And 64
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 120 '0x78 BIT 7, B
                    temp = regB
                    temp = temp And 128
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 121 '0x78 BIT 7, C
                    temp = regC
                    temp = temp And 128
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 122 '0x78 BIT 7, D
                    temp = regD
                    temp = temp And 128
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 123 '0x78 BIT 7, E
                    temp = regE
                    temp = temp And 128
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 124 '0x78 BIT 7, H
                    temp = regH
                    temp = temp And 128
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 125 '0x78 BIT 7, L
                    temp = regL
                    temp = temp And 128
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 126 '0x7E BIT 7, (HL)
                    temp = RAM(regH * 256& + regL)
                    temp = temp And 128
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                    Call writeRAM((regH * 256&) + regL, RAM((regH * 256&) + regL))
                Case 127 '0x7E BIT 7, A
                    temp = regA
                    temp = temp And 128
                    If temp > 0 Then
                        flagZ = 0
                    Else
                        flagZ = 1
                    End If
                Case 134 '0x86 RES 0, (HL)
                    temp = RAM((regH * 256&) + regL)
                    temp = temp And 254
                    RAM((regH * 256&) + regL) = temp
                    Call writeRAM((regH * 256&) + regL, RAM((regH * 256&) + regL))
                Case 135 '0x87 RES 0, A
                    regA = regA And 254
                Case 158 '0x9E RES 3, (HL)
                    temp = RAM((regH * 256&) + regL)
                    temp = temp And 247
                    RAM((regH * 256&) + regL) = temp
                    Call writeRAM((regH * 256&) + regL, RAM((regH * 256&) + regL))
                Case 190 '0xBE RES 7, (HL)
                    temp = RAM((regH * 256&) + regL)
                    temp = temp And 127
                    RAM((regH * 256&) + regL) = temp
                    Call writeRAM((regH * 256&) + regL, RAM((regH * 256&) + regL))
                Case 198 '0xC6 SET 0, (HL)
                    'temp = RAM((regH * 256&) + regL) Or 1
                    RAM((regH * 256&) + regL) = RAM((regH * 256&) + regL) Or 1
                    Call writeRAM((regH * 256&) + regL, RAM((regH * 256&) + regL))
                Case 216 '0xD8 SET 3, B
                    regB = regB Or 8
                Case 222 '0xDE SET 3, (HL)
                    'temp = RAM((regH * 256&) + regL) Or 8
                    RAM((regH * 256&) + regL) = RAM((regH * 256&) + regL) Or 8
                    Call writeRAM((regH * 256&) + regL, RAM((regH * 256&) + regL))
                Case 248 '0xF8 SET 7, B
                    regB = regB Or 128
                Case 254 '0xFE SET 7, (HL)
                    RAM((regH * 256&) + regL) = RAM((regH * 256&) + regL) Or 128
                    Call writeRAM((regH * 256&) + regL, RAM((regH * 256&) + regL))
                Case Else
                    Debug.Print "Opcode CB " & RAM(regPC + 1) & " NO IMPLEMENTADA"
            End Select
        Case 204 'CALL Z, a16
            If flagZ = 1 Then
                regSP = regSP - 2
                temp = regPC + OpCodeBytes(opCode)
                RAM(regSP) = temp \ 256
                RAM(regSP - 1) = temp And 255
                regPC = (RAM(regPC + 2) * 256&) + RAM(regPC + 1) - OpCodeBytes(opCode)
            End If
        Case 205 '0xCD CALL a16
            regSP = regSP - 2
            temp = regPC + OpCodeBytes(opCode)
            RAM(regSP) = temp \ 256
            RAM(regSP - 1) = temp And 255
            regPC = (RAM(regPC + 2) * 256&) + RAM(regPC + 1) - OpCodeBytes(opCode)
        Case 206 '0xCE ADC A, d8
            flagZ = 0
            flagH = 0
            temp = CInt(regA) + CInt(RAM(regPC + 1)) + flagC
            flagC = 0
            If temp > 255 Then
                flagC = 1
                regA = temp - 256
            Else
                regA = temp
            End If
            If regA = 0 Then
                flagZ = 1
            End If
            If regA And 15 Then
                flagH = 1
            End If
        Case 207 '0xCF RST 1
            regSP = regSP - 2
            temp = regPC + OpCodeBytes(opCode)
            RAM(regSP) = temp \ 256
            RAM(regSP - 1) = temp And 255
            regPC = 8 - OpCodeBytes(opCode)
        Case 208 '0xD0 RET NC
            If flagC = 0 Then
                temp = regSP \ 256
                temp2 = regSP And 255
                regPC = (RAM((temp * 256&) + temp2) * 256&) + RAM((temp * 256&) + temp2 - 1) - OpCodeBytes(opCode)
                regSP = regSP + 2
            End If
        Case 209 '0xD1 POP DE
            regD = RAM(regSP)
            regE = RAM(regSP - 1)
            regSP = regSP + 2
        Case 210 '0xD2 JP NC, a16
            If flagC = 0 Then
                regPC = (RAM(regPC + 2) * 256&) + RAM(regPC + 1) - OpCodeBytes(opCode)
            End If
        Case 211 '0xD3
            Debug.Print "Opcode 211 (0xD3) no existe. RegPC: " + regPCHex
        Case 212 '0xD4 CALL NC, a16
            If flagC = 0 Then
                regSP = regSP - 2
                temp = regPC + OpCodeBytes(opCode)
                RAM(regSP) = temp \ 256
                RAM(regSP - 1) = temp And 255
                regPC = (RAM(regPC + 2) * 256&) + RAM(regPC + 1) - OpCodeBytes(opCode)
            End If
        Case 213 '0xD5 PUSH DE
            regSP = regSP - 2
            RAM(regSP) = regD
            RAM(regSP - 1) = regE
        Case 214 '0xD6 SUB D8
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - RAM(regPC + 1)
            If temp < 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor RAM(regPC + 1) Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 215 '0xD7 RST 2
            regSP = regSP - 2
            temp = regPC + OpCodeBytes(opCode)
            RAM(regSP) = temp \ 256
            RAM(regSP - 1) = temp And 255
            regPC = 16 - OpCodeBytes(opCode)
        Case 216 '0xD8 RET C
            If flagC = 1 Then
                temp = regSP \ 256
                temp2 = regSP And 255
                regPC = (RAM((temp * 256&) + temp2) * 256&) + RAM((temp * 256&) + temp2 - 1) - OpCodeBytes(opCode)
                regSP = regSP + 2
            End If
        Case 217 '0xD9 RETI
            temp = regSP \ 256
            temp2 = regSP And 255
            regPC = (RAM((temp * 256&) + temp2) * 256&) + RAM((temp * 256&) + temp2 - 1) - OpCodeBytes(opCode)
            regSP = regSP + 2
            flagIME = True
        Case 218 '0xDA JP C, a16
            If flagC = 1 Then
                regPC = (RAM(regPC + 2) * 256&) + RAM(regPC + 1) - OpCodeBytes(opCode)
            End If
        Case 219 '0xDB
            Debug.Print "Opcode 219 (0xDB) no existe. RegPC: " + regPCHex
        Case 220 '0xDC CALL C, a16
            If flagC = 1 Then
                regSP = regSP - 2
                temp = regPC + OpCodeBytes(opCode)
                RAM(regSP) = temp \ 256
                RAM(regSP - 1) = temp And 255
                regPC = (RAM(regPC + 2) * 256&) + RAM(regPC + 1) - OpCodeBytes(opCode)
            End If
        Case 221 '0xDD
            Debug.Print "Opcode 221 (0xDD) no existe. RegPC: " + regPCHex
        Case 222 '0xDE SBC A, d8
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = CInt(regA) - RAM(regPC + 1) - flagC
            If temp And 256 > 0 Then
                flagC = 1
            End If
            temp = temp And 255
            If temp = 0 Then
                flagZ = 1
            End If
            If ((regA Xor RAM(regPC + 1) Xor temp) And 16) > 0 Then
                flagH = 1
            End If
            regA = temp
        Case 223 '0xDF RST 3
            regSP = regSP - 2
            temp = regPC + OpCodeBytes(opCode)
            RAM(regSP) = temp \ 256
            RAM(regSP - 1) = temp And 255
            regPC = 24 - OpCodeBytes(opCode)
        Case 224 'E0 LD (a8), A 'FF00 + a8 = A
            RAM(dirIOPorts + RAM(regPC + 1)) = regA
            Call writeRAM(dirIOPorts + RAM(regPC + 1), regA)
        Case 225 'E1 POP HL
            regH = RAM(regSP)
            regL = RAM(regSP - 1)
            regSP = regSP + 2
        Case 226 '0xE2 LD (C), A 'FF00 + C = A
            RAM(dirIOPorts + regC) = regA
            Call writeRAM(dirIOPorts + regC, regA)
        Case 227 '0xE3
            Debug.Print "Opcode 227 (0xE3) no existe. RegPC: " + regPCHex
        Case 228 '0xE4
            Debug.Print "Opcode 228 (0xE4) no existe. RegPC: " + regPCHex
        Case 229 'E5 PUSH HL
            regSP = regSP - 2
            RAM(regSP) = regH
            RAM(regSP - 1) = regL
        Case 230 'E6 AND d8
            flagZ = 0
            flagH = 1
            flagC = 0
            regA = regA And RAM(regPC + 1)
            If regA = 0 Then
                flagZ = 1
            End If
        Case 231 'E7 RST 4
            regSP = regSP - 2
            temp = regPC + OpCodeBytes(opCode)
            RAM(regSP) = temp \ 256
            RAM(regSP - 1) = temp And 255
            regPC = 32 - OpCodeBytes(opCode)
        Case 232 'E8 ADD SP, s8
            flagZ = 0
            flagH = 0
            flagC = 0
            temp = RAM(regPC + 1)
            If temp > 127 Then
                temp = -(((Not temp) + 1) And 255)
            End If
            temp2 = regSP + temp
            If (regSP Xor temp Xor temp2) And 16 Then
                flagH = 1
            End If
            If (regSP Xor temp Xor temp2) And 256 Then
                flagC = 1
            End If
            regSP = temp2 And 65535
        Case 233 'E9 JP HL
            regPC = (regH * 256&) + regL - OpCodeBytes(opCode)
        Case 234 'EA LD (a16), A
            temp = (RAM(regPC + 2) * 256&) + RAM(regPC + 1)
            If temp = 8192 Then 'Estan escribiendo a la direccion 2000h para hacer switching de banco de memoria MBC (NO IMPLEMENTADO)
                
            Else
                temp = (RAM(regPC + 2) * 256&) + RAM(regPC + 1)
                RAM(temp) = regA
                Call writeRAM(temp, regA)
                If temp >= 65280 And temp <= 65355 Or temp = 65535 Then
                    Debug.Print "RegPC:" & Hex(regPC) & " Opcode:" & Hex(opCode) & " Escritura al registro: " & temp
                End If
            End If
        Case 235 'EB
            Debug.Print "Opcode 235 (0xEB) no existe. RegPC: " + regPCHex
        Case 236 'EC
            Debug.Print "Opcode 236 (0xEC) no existe. RegPC: " + regPCHex
        Case 237 'ED
            Debug.Print "Opcode 237 (0xED) no existe. RegPC: " + regPCHex
        Case 238 '0xEE XOR D8
            flagC = 0
            flagH = 0
            flagZ = 0
            regA = regA Xor RAM(regPC + 1)
            If regA = 0 Then
                flagZ = 1
            End If
        Case 239 '0xEF RST 5
            regSP = regSP - 2
            temp = regPC + OpCodeBytes(opCode)
            RAM(regSP) = temp \ 256
            RAM(regSP - 1) = temp And 255
            regPC = 40 - OpCodeBytes(opCode)
        Case 240 'F0 LD A, (a8)
            regA = RAM(dirIOPorts + RAM(regPC + 1))
        Case 241 '0xF1 POP AF
            regA = RAM(regSP)
            regF = RAM(regSP - 1)
            regSP = regSP + 2
            flagZ = (regF And 128) \ 128
            flagN = (regF And 64) \ 64
            flagH = (regF And 32) \ 32
            flagC = (regF And 16) \ 16
        Case 242 '0xF2 LD A, (C)
            regA = RAM(dirIOPorts + regC)
        Case 243 '0xF3 DI
            flagIME = False
        Case 244 '0xF4
            Debug.Print "Opcode 244 (0xF4) no existe. RegPC: " + regPCHex
        Case 245: '0xF5 PUSH AF
            regF = 0
            regF = regF Or (flagZ * 128&)
            regF = regF Or (flagN * 64&)
            regF = regF Or (flagH * 32&)
            regF = regF Or (flagC * 16&)

            regSP = regSP - 2
            RAM(regSP) = regA
            RAM(regSP - 1) = regF
        Case 246 '0xF6 OR d8
            regA = regA Or RAM(regPC + 1)
        Case 247 '0xF7 RST 6
            regSP = regSP - 2
            temp = regPC + OpCodeBytes(opCode)
            RAM(regSP) = temp \ 256
            RAM(regSP - 1) = temp And 255
            regPC = 48 - OpCodeBytes(opCode)
        Case 248 '0xF8 LD HL, SP+s8
            flagZ = 0
            flagC = 0
            flagH = 0
            temp = RAM(regPC + 1)
            If temp > 127 Then
                temp = -(((Not temp) + 1) And 255)
            End If
            temp2 = (temp + regSP) And 65535
            regH = (temp2 \ 256) And 256
            regL = temp2 And 255
            If ((regSP Xor temp Xor temp2) And 256) = 256 Then
                flagC = 1
            End If
            If ((regSP Xor temp Xor temp2) And 16) = 16 Then
                flagH = 1
            End If
        Case 249 '0xF9 LD SP, HL
            regSP = regH * 256& + regL
        Case 250 '0xFA LD A, (a16)
            regA = RAM((RAM(regPC + 2) * 256&) + RAM(regPC + 1))
        Case 251 '0xFB EI
            setInterruptsInNextCycle = True 'Es que el IME = 1 no se hace aqui, sino después de la siguiente instrucción
            'flagIME = True
        Case 252 '0xFC
            Debug.Print "Opcode 252 (0xFC) no existe. RegPC: " + regPCHex
        Case 253 '0xFD
            Debug.Print "Opcode 253 (0xFD) no existe. RegPC: " + regPCHex
        Case 254 '0xFE CP d8
            flagZ = 0
            flagC = 0
            flagH = 0
            If regA = RAM(regPC + 1) Then
                flagZ = 1
            End If
            If regA < RAM(regPC + 1) Then
                flagC = 1
            End If
            If (regA And 15) < (RAM(regPC + 1) And 15) Then
                flagH = 1
            End If
        Case 255 '0xFF RST 7
            regSP = regSP - 2
            temp = regPC + OpCodeBytes(opCode)
            RAM(regSP) = temp \ 256
            RAM(regSP - 1) = temp And 255
            regPC = 56 - OpCodeBytes(opCode)
        Case Else
            Debug.Print "OPCODE " & RAM(regPC) & " no implementado"
    End Select
    
    'Negative flag
    Select Case opCode
        Case 4, 7, 9, 12, 15, 20, 23, 25, 28, 31, 36, 41, 44, 52, 55, 57, 60, 63
            flagN = 0
        Case 5, 13, 21, 29, 37, 45, 47, 53, 61
            flagN = 1
        Case 128 To 143
            flagN = 0
        Case 144 To 159
            flagN = 1
        Case 160 To 183
            flagN = 0
        Case 184 To 191
            flagN = 1
        Case 198, 206, 230, 232, 238, 246, 248
            flagN = 0
        Case 214, 222, 254
            flagN = 1
    End Select
    
    
    regPC = regPC + OpCodeBytes(opCode)
    'CPUCyclesCounter = CPUCyclesCounter + Cycles(opCode)
    tickNextInstruction = tick + Cycles(opCode)
    
    Call checkInterrupts
    If setInterruptsInNextCycle = True Then 'Así dejo que se ejecute la siguiente instrucción antes de que salte a una interrupción
        setInterruptsInNextCycle = False
        flagIME = True
    End If
End Sub

Private Sub PPU()
    Static PPUTick As Integer
    
    Dim posX, posY, pixelColor, SpriteColor As Byte
    Dim spritePosX, spritePosY, spriteID, spriteParams As Byte
    Dim screenPixelColor As Byte
    Dim i As Byte
    Dim numPixelsEsteColor As Byte
    Dim numArrayDentroDeEseColor As Byte
    Dim backgroundX, backgroundY, tileID As Integer
    Dim tileDataAddress As Long
    Dim prioridad As Boolean
    Dim paleta1 As Boolean
    
    Static dirOfSpriteToBeDrawn As Long
    Dim dirOAMOfObject As Long
    
    Dim currentPaintedRow As Byte
    
    If LCDEnabled Then 'Solo si el LCD está habilitado
        currentPaintedRow = RAM(regLY)
        If PPUTick >= 0 And PPUTick < 20 Then 'OAM SEARCH
            If PPUTick = 0 Then
                RAM(regSTAT) = (RAM(regSTAT) And 252) Or 2 'Pongo el modo 2
                If (RAM(regSTAT) And 32) Then 'Si está habilitada la interrupción para este modo
                    RAM(regIF) = RAM(regIF) Or 2 'Set el flag de interrupción de STAT
                End If
            End If
            If spriteEnabled And spritesDirty Then
                'Compruebo dos sprites por tick (ya que hay 20 ticks y 40 sprites)
                'Primer sprite
                dirOAMOfObject = dirOAM + ((PPUTick * 2) * 4)
                spritePosY = RAM(dirOAMOfObject + 0)
                spritePosX = RAM(dirOAMOfObject + 1)
                If currentPaintedRow + 16 >= spritePosY And currentPaintedRow + 16 < spritePosY + 8 Then 'Añadimos este sprite al array de sprites a pintar en esta linea
                    spritesDirtyBuffer(spritePosX) = dirOAMOfObject
                End If
                'Segundo sprite
                dirOAMOfObject = dirOAMOfObject + 4
                spritePosY = RAM(dirOAMOfObject + 0)
                spritePosX = RAM(dirOAMOfObject + 1)
                If currentPaintedRow + 16 >= spritePosY And currentPaintedRow + 16 < spritePosY + 8 Then 'Añadimos este sprite al array de sprites a pintar en esta linea
                    spritesDirtyBuffer(spritePosX) = dirOAMOfObject
                End If
            End If
        ElseIf PPUTick >= 20 And PPUTick < 60 Then 'PIXEL TRANSFER (MODE 3)
            If PPUTick = 20 Then
                RAM(regSTAT) = (RAM(regSTAT) And 252) Or 3 'Pongo el modo 3
            End If
        
            If (RAM(regLY) < 144) And (BGDirty = True Or spritesDirty = True) And (BGEnabled = True Or spriteEnabled = True) Then
                pixelColor = 0
                Dim BGTileBufferOffset As Byte
                If dirBGTileMap = 38912 Then '0x9800
                    BGTileBufferOffset = 0
                Else
                    BGTileBufferOffset = 32
                End If
                
                'Dibujo 4 pixeles consecutivos 'Como hay 40 ticks en este modo por cada tick dibujo 4 pixeles
                Dim pixelValue1, pixelValue2, pixelValueFinal, temp1, temp2 As Byte
                Dim posXInTile, posYInTile As Byte

                For i = 0 To 3
                    posX = ((PPUTick - 20) * 4) + i
                    posY = RAM(regLY)
                    If BGEnabled Then
                        backgroundX = posX \ 8
                        backgroundY = (posY + RAM(regSCY)) \ 8
                        tileID = CInt(RAM(dirBGTileMap + (backgroundY * 32) + backgroundX))
                        If tileID <> 0 Then
'                            Debug.Print "TILE=" & tileID & ":X=" & posX & ":Y=" & posY & ":backX=" & backgroundX & ":backY=" & backgroundY & ":SCY=" & RAM(regSCY)
                        End If
                        If dirBGTileData = 36864 And tileID < 128 Then 'Si el modo de direccionamiento es 0x9000
                            tileID = tileID + 256 'Estos empiezan a partir de 0x9000
                        End If
                        tileDataAddress = dirBGTileData + (tileID * 16&)
                        posXInTile = 7 - (posX - (backgroundX * 8&))
                        posYInTile = posY + RAM(regSCY) - (backgroundY * 8&)
                        If posYInTile <> 7 And tileID <> 0 Then
'                            Debug.Print "posYInTile=" & posYInTile
                        End If

                        pixelValueFinal = tilesBuffer(tileID, posXInTile, posYInTile)
                        pixelColor = pixelValueFinal
                    End If
                    If spriteEnabled And spritesDirty Then
                        If spritesDirtyBuffer(posX + 8) <> 0 Then
                            dirOfSpriteToBeDrawn = spritesDirtyBuffer(posX + 8) 'Nuevo sprite a dibujar
                        End If
                        If dirOfSpriteToBeDrawn <> 0 Then
                            spritePosX = RAM(dirOfSpriteToBeDrawn + 1) - 8
                            spritePosY = RAM(dirOfSpriteToBeDrawn + 0) - 16
                            spriteID = RAM(dirOfSpriteToBeDrawn + 2)
                            spriteParams = RAM(dirOfSpriteToBeDrawn + 3)
                            If (spriteParams And 128) = 0 Then
                                prioridad = True
                            Else
                                prioridad = False
                            End If
                            If (spriteParams And 16) > 0 Then
                                paleta1 = True
                            Else
                                paleta1 = False
                            End If

                            tileDataAddress = dirVRAM + (spriteID * 16&)
                            If ((posX - spritePosX) < 8) And ((posY - spritePosY) < 8) And ((posX - spritePosX) >= 0) And ((posY - spritePosY) >= 0) Then  'Todavia no hemos pintado el sprite entero
                                posXInTile = 7 - (posX - spritePosX)
                                posYInTile = posY - spritePosY
        
                                pixelValueFinal = tilesBuffer(spriteID, posXInTile, posYInTile)
                                SpriteColor = pixelValueFinal
                            Else
                                dirOfSpriteToBeDrawn = 0
                            End If
                        End If
                    End If
                    
                    'Pixel mixer (mezcla background con sprites)
                    If BGEnabled Then
                        If spriteEnabled And dirOfSpriteToBeDrawn <> 0 Then
                            If SpriteColor = 0 Then 'Si es 0 es transparente y sólo se dibuja el background
                                pixelColor = BGP(pixelColor)
                            Else
                                If (pixelColor = 0) Or (prioridad = True) Then 'Tiene prioridad sobre el BG así que se dibuja el sprite
                                    If paleta1 Then
                                        pixelColor = OBP1(SpriteColor)
                                    Else
                                        pixelColor = OBP0(SpriteColor)
                                    End If
                                Else
                                    pixelColor = BGP(pixelColor)
                                End If
                            End If
                        Else
                            pixelColor = BGP(pixelColor)
                        End If
                    End If
                                        
                    'Si el pixel no es del mismo color que ya había lo añado al array que se va a pintar cuando se refresque la pantalla
                    screenPixelColor = screen(posX, posY)
                    If pixelColor <> screenPixelColor Then
                        screen(posX, posY) = pixelColor
                        screenDirty = True
                        
                        pixelsPaintedPerFPS = pixelsPaintedPerFPS + 1
                        
                        numPixelsEsteColor = contadorPixelsDeCadaColor(pixelColor)
                        numArrayDentroDeEseColor = Int(numPixelsEsteColor / 20)
                            
                        If (numPixelsEsteColor Mod 20 <> 0) Then
                            colorArrayrangeStr(pixelColor, numArrayDentroDeEseColor) = colorArrayrangeStr(pixelColor, numArrayDentroDeEseColor) + ","
                        End If
                        colorArrayrangeStr(pixelColor, numArrayDentroDeEseColor) = colorArrayrangeStr(pixelColor, numArrayDentroDeEseColor) + ColumnLetter(posX + 1) + CStr(posY + 1)
                        contadorPixelsDeCadaColor(pixelColor) = contadorPixelsDeCadaColor(pixelColor) + 1
                    End If
                    
                Next i
            End If
                
        ElseIf PPUTick = 63 Then 'H-BLANK (MODE 0)
            RAM(regSTAT) = (RAM(regSTAT) And 252) Or 0 'Pongo el modo 0
            If (RAM(regSTAT) And 8) Then 'Si está habilitada la interrupción para este modo
                RAM(regIF) = RAM(regIF) Or 2 'Set el flag de interrupción de STAT
            End If
        ElseIf PPUTick = 114 Then 'Es mayor que 114, ha acabado la linea y tiene que empezar una nueva
            Erase spritesDirtyBuffer
        
            PPUTick = -1
            'spritesDirty = False
            Call updateLY
        End If
        
        PPUTick = PPUTick + 1
    Else
        'Debug.Print "LCD No habilitado"
    End If
End Sub

Private Sub updateLY()
    Dim i, j As Byte 'i: Del 1 al 160 (posicion X del pixel ---- j: Del 1 al 4 (colores)
    Dim colorXRange(3) As Range
    
    If LCDEnabled = True Then 'Si el LCD está habilitado
        If screenDirty = True Then
            For j = 0 To 3
                If colorArrayrangeStr(j, 0) <> "" Then
                    Set colorXRange(j) = Range(colorArrayrangeStr(j, 0))
                    For i = 1 To 8
                        If colorArrayrangeStr(j, i) <> "" Then
                            Set colorXRange(j) = Union(colorXRange(j), Range(colorArrayrangeStr(j, i)))
                        End If
                    Next i
                    colorXRange(j).Interior.Color = rgbColors(j)
                End If
            Next j
            
            Erase colorXRange
            Erase colorArrayrangeStr
            Erase contadorPixelsDeCadaColor
        End If
        
        RAM(regLY) = RAM(regLY) + 1 'INCREMENTO EL LCD Y COORDINATE
        If RAM(regLY) >= 0 And RAM(regLY) < 144 Then 'Filas de la 0 a la 143
            'pintaLineaConRango (RAM(regLY))
        ElseIf RAM(regLY) = 144 Then 'VBLANK
            RAM(regIF) = RAM(regIF) Or 1 'Pongo a 1 el flag del vblank
            RAM(regSTAT) = (RAM(regSTAT) And 252) Or 1 'Pongo el modo 1
        ElseIf RAM(regLY) > 153 Then 'Refresco la pantalla
            RAM(regLY) = 0
            
            FPS = FPS + 1
                        
            If screenDirty = True Then
                
                Application.ScreenUpdating = True
                screenDirty = False
                BGDirty = False
                spritesDirty = False
                
                DoEvents
                
                Application.ScreenUpdating = False
            Else
                DoEvents
            End If
            
            If GetAsyncKeyState(vbKeyEscape) Then
                teclaPulsadaEscape = True
            End If

            If GetAsyncKeyState(vbKeyA) Then
                teclaPulsadaA = True
            Else
                teclaPulsadaA = False
            End If
            If GetAsyncKeyState(vbKeyS) Then
                teclaPulsadaB = True
            Else
                teclaPulsadaB = False
            End If
            If GetAsyncKeyState(vbKeyReturn) Then
                teclaPulsadaEnter = True
            Else
                teclaPulsadaEnter = False
            End If
            If GetAsyncKeyState(vbKeyControl) Then
                teclaPulsadaSelect = True
            Else
                teclaPulsadaSelect = False
            End If
            If GetAsyncKeyState(vbKeyRight) Then
                teclaPulsadaFlechaDcha = True
            Else
                teclaPulsadaFlechaDcha = False
            End If
            If GetAsyncKeyState(vbKeyLeft) Then
                teclaPulsadaFlechaIzda = True
            Else
                teclaPulsadaFlechaIzda = False
            End If
            If GetAsyncKeyState(vbKeyUp) Then
                teclaPulsadaFlechaArriba = True
            Else
                teclaPulsadaFlechaArriba = False
            End If
            If GetAsyncKeyState(vbKeyDown) Then
                teclaPulsadaFlechaAbajo = True
            Else
                teclaPulsadaFlechaAbajo = False
            End If
        End If
        
        If (RAM(regLY) = RAM(regLYC)) And (RAM(regSTAT) And 64) Then 'LY = LYC y está habilitada la interrupción?
            RAM(regIF) = RAM(regIF) Or 2 'Set el flag de interrupción de STAT
        End If
    End If
End Sub

Private Sub checkInterrupts()
    Dim vBlankSet As Boolean
    Dim statSet As Boolean
    Dim timerSet As Boolean
    Dim serialSet As Boolean
    Dim joypadSet As Boolean
    
    If flagIME = True Then
        vBlankSet = (RAM(regIE) And 1) And (RAM(regIF) And 1)
        statSet = (RAM(regIE) And 2) And (RAM(regIF) And 2)
        timerSet = (RAM(regIE) And 4) And (RAM(regIF) And 4)
        serialSet = (RAM(regIE) And 8) And (RAM(regIF) And 8)
        joypadSet = (RAM(regIE) And 16) And (RAM(regIF) And 16)
        
        If vBlankSet Or statSet Or timerSet Or serialSet Or joypadSet Then
            regSP = regSP - 2
            RAM(regSP) = regPC \ 256
            RAM(regSP - 1) = regPC And 255
            flagIME = False 'Deshabilito las interrupciones
        End If
        
        If vBlankSet Then 'INTERRUPCION VBLANK
            regPC = 64 '-> Jump to interrupt vector 0x40
            RAM(regIF) = RAM(regIF) And 254 'Reseteo el flag
        ElseIf statSet Then
            regPC = 72 '-> Jump to interrupt vector 0x48
            RAM(regIF) = RAM(regIF) And 253 'Reseteo el flag
        ElseIf timerSet Then
            regPC = 80 '-> Jump to interrupt vector 0x50
            RAM(regIF) = RAM(regIF) And 251 'Reseteo el flag
        ElseIf serialSet Then
            regPC = 88 '-> Jump to interrupt vector 0x58
            RAM(regIF) = RAM(regIF) And 247 'Reseteo el flag
        ElseIf joypadSet Then
            regPC = 96 '-> Jump to interrupt vector 0x60
            RAM(regIF) = RAM(regIF) And 239 'Reseteo el flag
        End If
    End If
End Sub

Private Sub writeRAM(RAMAddress As Long, RAMValue As Byte)
    Dim backgroundX, backgroundY As Integer
    Dim initialDir As Long
    
    Select Case RAMAddress
        Case 8192: 'MBC Switch (no implementado)
        Case 0 To 32767 '0 to 7FFF (ROM)
            Debug.Print "Escritura a ROM no permitida"
        Case 32768 To 38911
            If (RAMAddress And 1) = 1 Then
                Dim numTile As Integer
                Dim posXInTile, posYInTile As Byte
                Dim pixelValue1, pixelValue2, pixelValueFinal As Byte
                Dim temp1, temp2 As Byte
                
                numTile = (RAMAddress - 32768) \ 16
                posYInTile = (RAMAddress - (32768 + (numTile * 16))) \ 2
                pixelValue1 = RAM(RAMAddress - 1)
                pixelValue2 = RAM(RAMAddress)
                For posXInTile = 7 To 0 Step -1
                    temp1 = (pixelValue1 \ (2 ^ posXInTile)) And 1 'ATENCION QUE LA BARRA DE DIVISION VA AL REVES "\" en vez de "/"
                    temp2 = (pixelValue2 \ (2 ^ posXInTile)) And 1
                    pixelValueFinal = (temp2 * 2) + temp1
                    tilesBuffer(numTile, posXInTile, posYInTile) = pixelValueFinal
                Next posXInTile
            End If
        Case 38912 To 38912 + 1023 '0x9800 (BGTileMap)
            If RAMValue <> 0 Then
                initialDir = 38912
                backgroundY = (RAMAddress - initialDir) \ 32
                backgroundX = (RAMAddress - initialDir) Mod 32
                BGDirty = True
            End If
        Case 39936 To 39936 + 1023 '0x9C00 (BGTileMap2)
            If RAMValue <> 0 Then
                initialDir = 39936
                backgroundY = (RAMAddress - initialDir) \ 32
                backgroundX = (RAMAddress - initialDir) Mod 32
                BGDirty = True
            End If
        Case 65024 To 65183 'FE00 to FE9F (OAM)
            spritesDirty = True
            'Debug.Print "OAM table changed"
        Case regJOYPAD: 'FF00 Joypadb
                If RAM(regJOYPAD) And 32 Then 'P14
                    'RAM(regJOYPAD) = RAM(regJOYPAD) Or 15
                    RAM(regJOYPAD) = RAM(regJOYPAD) Or 47
                    If teclaPulsadaFlechaAbajo Then
                        RAM(regJOYPAD) = RAM(regJOYPAD) And 247 'Pongo el bit 3 (P13) a nivel bajo (flecha abajo)
                    End If
                    If teclaPulsadaFlechaArriba Then
                        RAM(regJOYPAD) = RAM(regJOYPAD) And 251 'Pongo el bit 2 (P12) a nivel bajo (flecha arriba)
                    End If
                    If teclaPulsadaFlechaIzda Then
                        RAM(regJOYPAD) = RAM(regJOYPAD) And 253 'Pongo el bit 1 (P11) a nivel bajo (flecha izquierda)
                    End If
                    If teclaPulsadaFlechaDcha Then
                        RAM(regJOYPAD) = RAM(regJOYPAD) And 46 'Pongo el bit 0 (P10) a nivel bajo (flecha derecha)
                    End If
                ElseIf RAM(regJOYPAD) And 16 Then 'P15
                    RAM(regJOYPAD) = 31 'Pongo todos a nivel alto, es decir, no presionados
                    If teclaPulsadaEnter Then
                        RAM(regJOYPAD) = RAM(regJOYPAD) And 247 'Pongo el bit 3 (P13) a nivel bajo (start presionado)
                    End If
                    If teclaPulsadaSelect Then
                        RAM(regJOYPAD) = RAM(regJOYPAD) And 251 'Pongo el bit 2 (P12) a nivel bajo (select presionado)
                    End If
                    If teclaPulsadaB Then
                        RAM(regJOYPAD) = RAM(regJOYPAD) And 253  'Pongo el bit 1 (P11) a nivel bajo (B presionado)
                    End If
                    If teclaPulsadaA Then
                        RAM(regJOYPAD) = RAM(regJOYPAD) And 254 'Pongo el bit 0 (P10) a nivel bajo (A presionado)
                    End If
                End If
        Case regSB 'Serial port, no implementado
        Case regSC 'Serial port, no implementado
        Case regDIV 'Divider register
            RAM(regDIV) = 0
        Case regTMA
            RAM(regTMA) = RAMValue
            Debug.Print Hex(regPC) & ":RAMAddress FF06 (TMA TIMER MODULO)" & " Valor: " & Hex(RAMValue)
        Case regTAC 'Timer Control
            RAM(regTAC) = RAMValue
            Debug.Print Hex(regPC) & ":RAMAddress FF07 (TIMER CONTROL)" & " Valor: " & Hex(RAMValue)
            If RAMValue And 4 > 0 Then
                TimerEnabled = True
            Else
                TimerEnabled = False
            End If
            Select Case (RAMValue And 3)
                Case 0
                    TimerControlClockSelectedValue = 1024
                Case 1
                    TimerControlClockSelectedValue = 16
                Case 2
                    TimerControlClockSelectedValue = 64
                Case 3
                    TimerControlClockSelectedValue = 256
            End Select
        Case regIF 'FF0F IF
            'Debug.Print Hex(regPC) & ":RAMAddress FF0F (Interrupt flag)" & " Valor: " & Hex(RAMValue)
        Case regNR51 'Sonido, no implementado
        Case regLCDC 'LCDC (FF40)
            'Debug.Print Hex(regPC) & ":RAMAddress FF40 (LCDC register)" & " Valor: " & Hex(RAMValue)
            If (RAM(regLCDC) And 128) = 128 Then 'Si el LCD está habilitado
                If LCDEnabled = False Then
                    Debug.Print "LCD ENABLED"
                End If
                LCDEnabled = True
            Else
                If LCDEnabled = True Then
                    Debug.Print "LCD DISABLED"
                End If
                LCDEnabled = False
            End If
            If (RAMValue And 32) = 0 Then
                If windowEnabled = True Then
                    Debug.Print "WINDOW DISABLED"
                End If
                windowEnabled = False
            Else
                If windowEnabled = False Then
                    Debug.Print "WINDOW ENABLED"
                End If
                windowEnabled = True
            End If
            If (RAMValue And 16) = 0 Then
                If dirBGTileData <> 36864 Then
                    Debug.Print "OLD: " & dirBGTileData & " dirBGTileData = 0x9000"
                End If
                dirBGTileData = 36864 '0x9000
            Else
                If dirBGTileData <> 32768 Then
                    Debug.Print "OLD: " & dirBGTileData & " dirBGTileData = 0x8000"
                End If
                dirBGTileData = dirVRAM
            End If
            If (RAMValue And 8) = 0 Then
                If dirBGTileMap <> 38912 Then
                    Debug.Print "dirBGTileMap = 0x9800"
                End If
                dirBGTileMap = 38912 '0x9800
                BGDirty = True
            Else
                If dirBGTileMap <> 39936 Then
                    Debug.Print "dirBGTileMap = 0x9C00"
                End If
                dirBGTileMap = 39936 '0x9C00
                BGDirty = True
            End If
            If (RAMValue And 4) = 0 Then
                If spritesSize = True Then '8x16
                    Debug.Print "SPRITES SIZE = 8x8"
                End If
                spritesSize = False
            Else
                If spritesSize = False Then
                    Debug.Print "SPRITES SIZE = 8X16"
                End If
                spritesSize = True
            End If
            If (RAMValue And 2) = 0 Then
                If spriteEnabled = True Then
                    Debug.Print "SPRITES DISABLED"
                End If
                spriteEnabled = False
            Else
                If spriteEnabled = False Then
                    Debug.Print "SPRITES ENABLED"
                End If
                spriteEnabled = True
            End If
            If (RAMValue And 1) = 0 Then
                If BGEnabled = True Then
                    Debug.Print "BACKGROUND DISABLED"
                End If
                BGEnabled = False
            Else
                If BGEnabled = False Then
                    Debug.Print "BACKGROUND ENABLED"
                End If
                BGEnabled = True
            End If
        Case regSTAT 'STAT (FF41)
            Debug.Print Hex(regPC) & ":RAMAddress FF41 (STAT register)" & " Valor: " & Hex(RAMValue)
        Case regSCY
            If RAMValue <> 0 Then
                Debug.Print "SCY: " & RAMValue
            End If
        Case regSCX
            If RAMValue <> 0 Then
                Debug.Print "SCX: " & RAMValue
            End If
        Case regBGP
            Debug.Print "BGP: " & RAMValue
        Case regOBP0
            Debug.Print "OBP0: " & RAMValue
        Case regOBP1
            Debug.Print "OBP1: " & RAMValue
        Case regBIOS
            Debug.Print "BIOS: " & RAMValue
            If RAMValue <> 0 Then
                Debug.Print "BIOS DISABLED"
                'Recargo la ROM para quitar de la RAM los primeros 100 bytes de la BIOS
                Call loadROM
            End If
        Case regLY 'LY
            Debug.Print "REG_LY: " & RAMValue
        Case regLYC
            Debug.Print "REG_LYC: " & RAMValue
        Case regDMA 'DMA
            Call writeToOAM
        Case regIE
            Debug.Print "RegIE: " & Hex(RAMValue)
        Case 65505 'Variable GAME_STATUS guardada en memoria
                Debug.Print "Tick: " & tick & " regPC: " & Hex(regPC) & ":GAME STATUS: " & Hex(RAMValue)
        Case Else
            If RAMAddress > 65280 And RAMAddress <= 65355 Then
                'Debug.Print Hex(regPC) & ":RAMAddress " & Hex(RAMAddress); "(" & IOPortsNames(RAMAddress - 65280) & ")" & " no implementado. Nuevo valor: " & Hex(RAMValue)
            ElseIf RAMAddress > 65356 And RAMAddress = 65451 Then 'Pause status
                Static pausevalue As Byte
                
                Debug.Print Hex(regPC) & ":RAMAddress " & Hex(RAMAddress) & " Valor: " & Hex(RAMValue)
                If RAMValue = 0 Then
                    
                    If pausevalue = 1 Then 'Estabamos en pausa pero lo quitamos
                        'debug2Excel = True
                    End If
                End If
                pausevalue = RAMValue
            End If
    End Select
End Sub

Private Sub Timers()
    If tick Mod 256 = 0 Then
        If RAM(regDIV) = 255 Then
            RAM(regDIV) = 0
        Else
            RAM(regDIV) = RAM(regDIV) + 1
        End If
    End If
    If TimerEnabled Then
        If (tick Mod TimerControlClockSelectedValue) = 0 Then
            If RAM(regTIMA) = 255 Then
                RAM(regTIMA) = RAM(regTMA)
                'Set la interrupcion del timer
                RAM(regIF) = RAM(regIF) Or 4 'Pongo a 1 el flag del timer
            Else
                RAM(regTIMA) = RAM(regTIMA) + 1
            End If
        End If
    End If
End Sub

Private Sub writeToOAM()
    Dim i, x, y As Byte
    Dim spritePosX As Byte
    Dim spritePosY As Byte
    Dim spriteID As Integer
    Dim spriteParams As Byte

    Static oldCRC As Long
    Dim newCRC As Long
        
    
    spritesDirty = False
    
    For i = 0 To 39
        spritePosY = RAM((RAM(regDMA) * 256&) + (i * 4) + 0)
        spritePosX = RAM((RAM(regDMA) * 256&) + (i * 4) + 1)
        spriteID = RAM((RAM(regDMA) * 256&) + (i * 4) + 2)
        spriteParams = RAM((RAM(regDMA) * 256&) + (i * 4) + 3)
        newCRC = newCRC + spritePosX + spritePosY + spriteID + spriteParams
        RAM(dirOAM + (i * 4) + 0) = spritePosY
        RAM(dirOAM + (i * 4) + 1) = spritePosX
        RAM(dirOAM + (i * 4) + 2) = spriteID
        RAM(dirOAM + (i * 4) + 3) = spriteParams
    Next i
    
    If newCRC <> oldCRC Then
        oldCRC = newCRC
        spritesDirty = True
    End If
End Sub

Private Sub inicializaVideoRAM()
    Dim x, y As Byte
    
    For y = 1 To 144
        For x = 1 To 160
            screen(x, y) = 255
        Next x
    Next y
End Sub

Private Function ColumnLetter(ColumnNumber As Byte) As String
    Dim N As Byte
    Dim c As Byte
    Dim s As String
    
    N = ColumnNumber
    Do
        c = ((N - 1) Mod 26)
        s = Chr(c + 65) & s
        N = (N - c) \ 26
    Loop While N > 0
    ColumnLetter = s
End Function

Private Sub changeColumnWidth()
    'MsgBox ActiveWindow.PointsToScreenPixelsX(Columns(1).Width)
    Worksheets("Hoja1").Range("A:" & ColumnLetter(160)).ColumnWidth = 0.1
    Worksheets("Hoja1").Range("A1:A144").RowHeight = 1
End Sub

Private Sub initScreen()
    Dim x, y As Long
    
    Application.ScreenUpdating = False
    For y = 1 To 144
        For x = 1 To 160
            Range("A1").Offset(y, x).Interior.Color = rgbColors(0)
        Next x
    Next y
    Application.ScreenUpdating = True
End Sub

Private Sub loadROM()
    'Open binary ROM
    Dim fileInt As Integer
    fileInt = FreeFile
     
    Open "tetris.gb" For Binary Access Read As #fileInt
    Get #fileInt, , RAM
    
    Close #fileInt
End Sub

Private Sub loadBIOS()
    'Open binary ROM
    Dim fileInt As Integer
    fileInt = FreeFile
     
    Open "C:\Users\c20895\Documents\Documentos\Backups\MacrosBackups\a\bio.jpg" For Binary Access Read As #fileInt
    Get #fileInt, , BIOS
    
    Close #fileInt
End Sub

Private Sub initializeRegisters()
    regA = 1 '0x01
    regF = 176 '0xB0
    regB = 0 '0x00
    regC = 19 '0x13
    regD = 0 '0x00
    regE = 216 '0xD8
    regH = 1 '0x01
    regL = 77 '0x4D
    regSP = 65534 '0xFFFE
'    regPC = 256 '0x0100
    regPC = 0 'Para ejecutar la BIOS
    flagZ = 1
    flagN = 0
    flagH = 1
    flagC = 1
    flagIME = False
    RAM(regIF) = 225 '0xE1
    RAM(regLCDC) = 145 '0x91
    RAM(regIE) = 0 'All disabled
    RAM(regTIMA) = 0
    RAM(regTMA) = 0
    RAM(regTAC) = 0
    
    dirBGTileMap = 38912 '0x9800
    dirBGTileData = dirVRAM '0x8000
    
    LCDEnabled = False
    BGEnabled = False
    windowEnabled = False
    spriteEnabled = False
    spritesSize = False
    BGDirty = False
    windowDirty = False
    spritesDirty = True
    TimerEnabled = False
End Sub

Private Sub createCyclesArray()
    Dim CyclesString As String
    Dim CyclesSplit As Variant
    Dim i As Integer
    
    CyclesString = ""
    CyclesString = CyclesString & " 1,3,2,2,1,1,2,1,5,2,2,2,1,1,2,1"  '0
    CyclesString = CyclesString & ",1,3,2,2,1,1,2,1,3,2,2,2,1,1,2,1" '1
    CyclesString = CyclesString & ",3,3,2,2,1,1,2,1,3,2,2,2,1,1,2,1" '2
    CyclesString = CyclesString & ",3,3,2,2,3,3,3,1,3,2,2,2,1,1,2,1" '3
    CyclesString = CyclesString & ",1,1,1,1,1,1,2,1,1,1,1,1,1,1,2,1" '4
    CyclesString = CyclesString & ",1,1,1,1,1,1,2,1,1,1,1,1,1,1,2,1" '5
    CyclesString = CyclesString & ",1,1,1,1,1,1,2,1,1,1,1,1,1,1,2,1" '6
    CyclesString = CyclesString & ",1,1,1,1,1,1,2,1,1,1,1,1,1,1,2,1" '7
    CyclesString = CyclesString & ",1,1,1,1,1,1,2,1,1,1,1,1,1,1,2,1" '8
    CyclesString = CyclesString & ",1,1,1,1,1,1,2,1,1,1,1,1,1,1,2,1" '9
    CyclesString = CyclesString & ",1,1,1,1,1,1,2,1,1,1,1,1,1,1,2,1" 'A
    CyclesString = CyclesString & ",1,1,1,1,1,1,2,1,1,1,1,1,1,1,2,1" 'B
    CyclesString = CyclesString & ",5,3,4,4,6,4,2,4,5,4,4,2,6,6,2,4" 'C
    CyclesString = CyclesString & ",5,3,4,0,6,4,2,4,5,4,4,0,6,0,2,4" 'D
    CyclesString = CyclesString & ",3,3,2,0,0,4,2,4,4,1,4,0,0,0,2,4" 'E
    CyclesString = CyclesString & ",3,3,2,1,0,4,2,4,3,2,4,1,0,0,2,4" 'F
    
    CyclesSplit = Split(CyclesString, ",")
    
    For i = LBound(CyclesSplit) To UBound(CyclesSplit)
        Cycles(i) = CyclesSplit(i)
    Next i
    
End Sub

Private Sub createOpCodeBytesArray()
    Dim OpCodeBytesString As String
    Dim OpCodeBytesSplit As Variant
    Dim i As Integer
    
    OpCodeBytesString = "1,3,1,1,1,1,2,1,3,1,1,1,1,1,2,1" '0
    OpCodeBytesString = OpCodeBytesString & ",2,3,1,1,1,1,2,1,2,1,1,1,1,1,2,1" '1
    OpCodeBytesString = OpCodeBytesString & ",2,3,1,1,1,1,2,1,2,1,1,1,1,1,2,1" '2
    OpCodeBytesString = OpCodeBytesString & ",2,3,1,1,1,1,2,1,2,1,1,1,1,1,2,1" '3
    OpCodeBytesString = OpCodeBytesString & ",1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1" '4
    OpCodeBytesString = OpCodeBytesString & ",1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1" '5
    OpCodeBytesString = OpCodeBytesString & ",1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1" '6
    OpCodeBytesString = OpCodeBytesString & ",1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1" '7
    OpCodeBytesString = OpCodeBytesString & ",1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1" '8
    OpCodeBytesString = OpCodeBytesString & ",1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1" '9
    OpCodeBytesString = OpCodeBytesString & ",1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1" 'A
    OpCodeBytesString = OpCodeBytesString & ",1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1" 'B
    OpCodeBytesString = OpCodeBytesString & ",1,1,3,3,3,1,2,1,1,1,3,2,3,3,2,1" 'C
    OpCodeBytesString = OpCodeBytesString & ",1,1,3,0,3,1,2,1,1,1,3,0,3,0,2,1" 'D
    OpCodeBytesString = OpCodeBytesString & ",2,1,1,0,0,1,2,1,2,1,3,0,0,0,2,1" 'E
    OpCodeBytesString = OpCodeBytesString & ",2,1,1,1,0,1,2,1,2,1,3,1,0,0,2,1" 'F
    
    
    OpCodeBytesSplit = Split(OpCodeBytesString, ",")
    
    For i = LBound(OpCodeBytesSplit) To UBound(OpCodeBytesSplit)
        OpCodeBytes(i) = OpCodeBytesSplit(i)
    Next i
    
End Sub

Private Sub createOpCodeNemonics()
    Dim OpCodeNemString As String
    Dim OpCodeNemSplit As Variant
    Dim i As Integer
    
    OpCodeNemString = "NOP;LD BC,D16;LD (BC),A;INC BC;INC B;DEC B;LD B,D8;RLCA,;LD (A16),SP;ADD HL,BC;LD A,(BC);DEC BC;INC C;DEC C;LD C,D8;RRCA" '0
    OpCodeNemString = OpCodeNemString & ";STOP;LD DE,D16;LD (DE),A;INC DE;INC D;DEC D;LD D,D8;RLA;JR S8;ADD HL,DE;LD A,(DE);DEC DE;INC E;DEC E;LD E,D8;RRA" '1
    OpCodeNemString = OpCodeNemString & ";JR NZ,S8;LD HL,D16;LD (HL+),A;INC HL;INC H;DEC H;LD H,D8;DAA;JR Z,S8;ADD HL,HL;LD A,(HL+);DEC HL;INC L;DEC L;LD L,D8;CPL" '2
    OpCodeNemString = OpCodeNemString & ";JR NC,S8;LD SP,D16;LD (HL-),A;INC SP;INC (HL);DEC (HL);LD (HL),D8;SCF;JR C,S8;ADD HL,SP;LD A,(HL-);DEC SP;INC A;DEC A;LD A,D8;CCF" '3
    OpCodeNemString = OpCodeNemString & ";LD B,B;LD B,C; LD B,D;LD B,E;LD B,H;LD B,L;LD B,(HL);LD B,A;LD C,B;LD C,C;LD C,D;LD C,E;LD C,H;LD C,L;LD C,(HL);LD C,A" '4
    OpCodeNemString = OpCodeNemString & ";LD D,B;LD D,C;LD D,D;LD D,E;LD D,H;LD D,L;LD D,(HL);LD D,A;LD E,B;LD E,C;LD E,D;LD E,E;LD E,H;LD E,L;LD E,(HL);LD E,A" '5
    OpCodeNemString = OpCodeNemString & ";LD H,B;LD H,C;LD H,D;LD H,E;LD H,H;LD H,L;LD H,(HL);LD H,A;LD L,B;LD L,C;LD L,D;LD L,E;LD L,H;LD L,L;LD L,(HL);LD L,A" '6
    OpCodeNemString = OpCodeNemString & ";LD (HL),B;LD (HL),C;LD (HL),D;LD(HL),E;LD (HL),H;LD (HL),L;HALT;LD (HL),A;LD A,B;LD A,C;LD A,D;LD A,E;LD A,H;LD A,L;LD A,(HL);LD A,A" '7
    OpCodeNemString = OpCodeNemString & ";ADD A,B;ADD A,C;ADD A,D;ADD A,E;ADD A,H;ADD A,L;ADD A,(HL);ADD A,A;ADC A,B;ADC A,C;ADC A,D;ADC A,E;ADC A,H;ADC A,L;ADC A,(HL);ADC A,A" '8
    OpCodeNemString = OpCodeNemString & ";SUB B;SUB C;SUB D;SUB E;SUB H;SUB L;SUB (HL);SUB A;SBC A,B;SBC A,C;SBC A,D;SBC A,E;SBC A,H;SBC A,L;SBC A,(HL);SBC A,A" '9
    OpCodeNemString = OpCodeNemString & ";AND B;AND C;AND D;AND E;AND H;AND L;AND (HL);AND A;XOR B;XOR C;XOR D;XOR E;XOR H;XOR L; XOR (HL);XOR A" 'A
    OpCodeNemString = OpCodeNemString & ";OR B;OR C;OR D;OR E;OR H;OR L;OR (HL);OR A;CP B;CP C;CP D;CP E;CP H; CP L;CP (HL);CP A" 'B
    OpCodeNemString = OpCodeNemString & ";RET NZ;POP BC;JP NZ,A16;JP A16;CALL NZ,A16;PUSH BC;ADD A,D8;RST 0;RET Z;RET;JP Z,A16;NULL;CALL Z,A16;CALL A16;ADC A,D8;RST 1" 'C
    OpCodeNemString = OpCodeNemString & ";RET NC;POP DE;JP NC,A16;NULL;CALL NC,A16;PUSH DE;SUB D8;RST 2;RET C;RETI;JP C,A16;NULL;CALL C,A16;NULL;SBC A,D8;RST 3" 'D
    OpCodeNemString = OpCodeNemString & ";LD (A8),A;POP HL;LD (C),A;NULL;NULL;PUSH HL;AND D8;RST 4;ADD SP,S8;JP HL;LD (A16),A;NULL;NULL;NULL;XOR D8;RST 5" 'E
    OpCodeNemString = OpCodeNemString & ";LD A,(A8);POP AF;LD A,(C);DI;NULL;PUSH AF;OR D8;RST 6;LD HL,SP+S8;LD SP,HL;LD A,(A16);EI;NULL;NULL;CP D8;RST 7" 'F
            
    OpCodeNemSplit = Split(OpCodeNemString, ";")
    
    For i = LBound(OpCodeNemSplit) To UBound(OpCodeNemSplit)
        OpCodeNemonic(i) = OpCodeNemSplit(i)
    Next i
End Sub

Private Sub createIOPortsNames()
    Dim IOPortsString As String
    Dim IOPortsSplit As Variant
    Dim i As Integer
    
    IOPortsString = "P1;SB;SC;;DIV;TIMA;TMA;TAC;;;;;;;;IF;NR10;NR11;NR12;NR13;NR14;;NR21;NR22;NR23;NR24;NR30;NR31;NR32;NR33;NR34;;NR41;NR42;NR43;NR44;NR50;NR51;NR52;;;;;;;;;;WAVE;WAVE;WAVE;WAVE;WAVE;WAVE;WAVE;WAVE;WAVE;WAVE;WAVE;WAVE;WAVE;WAVE;WAVE;WAVE;LCDC;STAT;SCY;SCX;LY;LYC;DMA;BGP;OBP0;OBP1;WY;WX"
            
    IOPortsSplit = Split(IOPortsString, ";")
    
    For i = LBound(IOPortsSplit) To UBound(IOPortsSplit)
        IOPortsNames(i) = IOPortsSplit(i)
    Next i
End Sub

Private Sub DeclareKeys()
    Application.OnKey "{UP}", ""
    Application.OnKey "{DOWN}", ""
    Application.OnKey "{LEFT}", ""
    Application.OnKey "{RIGHT}", ""
    Application.OnKey "a", ""
    Application.OnKey "{ENTER}", ""
    Application.OnKey "{ESC}", ""
End Sub

Private Sub printInitialDebugInfo()
    Dim rowNumber As Long
    
    For rowNumber = 0 To 32767
        With Sheets("Hoja2")
            .Range("A1").Offset(rowNumber, 0).Value = Replace(Space(4 - Len(Hex(rowNumber))), " ", "0") & Hex(rowNumber)
            .Range("B1").Offset(rowNumber, 0).Value = Replace(Space(4 - Len(Hex(RAM(rowNumber)))), " ", "0") & Hex(RAM(rowNumber))
        End With
    Next rowNumber
End Sub
