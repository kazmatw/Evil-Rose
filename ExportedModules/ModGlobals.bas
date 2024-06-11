Attribute VB_Name = "ModGlobals"
Public Declare PtrSafe Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As LongPtr, ByVal Flags As Long) As LongPtr
'Declare Windows API Function, using "PtrSafe" keyword to run in 64bit environment

Public Const KLF_SETFORPROCESS = &H1
'This const let this code only affect current thread (in Excel)

' Keyboard Layout code
Public Const HKL_ENGLISH As LongPtr = &H4090409 ' ENG(US)
Public Const HKL_CHINESE_TRADITIONAL_PHONETIC As LongPtr = &H4040404 ' �c�餤�� (�`��)


Public Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal HWnd As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As LongPtr, _
    ByVal lpTimerFunc As LongPtr) As Long

Public Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal HWnd As LongPtr, _
    ByVal nIDEvent As LongPtr) As Long
    
'Variables Declaration

Public Type TBlo                            'Block
    Arr(6, 6) As Byte                           'Array
    BriCol As Long                              'Bright Color
    NorCol As Long                              'Normal Color
    DarCol As Long                              'Dark Color
    ColInd As Byte                              'Color Index
    Siz As Byte                                 'Size
    X As Byte                                   'Row
    Y As Byte                                   'Column
End Type

Public Type TBloSet                         'Block Set
    Blo() As Byte                               'Block
End Type

Public Type TCol                            'Color
    Bri As Long                                 'Bright
    Nor As Long                                 'Normal
    Dar As Long                                 'Dark
End Type

Public Type TPlaFie                         'Playing Field
    BacCol1 As Long                             'Background Color 1
    BacCol2 As Long                             'Background Color 2
    BorBCol As Long                             'Border Bright Color
    BorNCol As Long                             'Border Normal Color
    BorDCol As Long                             'Border Dark Color
    W As Byte                                   'Width
    H As Byte                                   'Height
    X As Byte                                   'Row
    Y As Byte                                   'Column
End Type

Public Type TSta                            'Statistics
    Blo As Integer                              'Blocks
    Lev As Long                                 'Level
    LevPro As Integer                           'Level Progress
    Row As Integer                              'Rows
    Gap As Single                               'Gapless
    GapSum As Integer                           'Gapless Sum
    Sco As Long                                 'Score
    ScoMax As Long                              'Score Max
    Qua As Integer                              'Quads
    DouQua As Integer                           'Double Quads
End Type

Public Type TStaFie                         'Statistics Field
    BacCol1 As Long                             'Background Color 1
    BacCol2 As Long                             'Background Color 2
    BorBCol As Long                             'Border Bright Color
    BorNCol As Long                             'Border Normal Color
    BorDCol As Long                             'Border Dark Color
    W As Byte                                   'Width
    H As Byte                                   'Height
    X As Byte                                   'Row
    Y As Byte                                   'Column
End Type

Public Type TTim                            'Timer
    CurPas As Byte                              'Current Pass
    ExeThr As Byte                              'Execution Threshold
    ExeThrDef As Byte                           'Execution Threshold Default
    LevTim As Byte                              'Level Timer
End Type

Public ColLib(112) As TCol                  'Color Library
Public ColSet() As Byte                     'Color Sets
Public BloLib(7) As TBlo                    'Block Library
Public BloSet() As TBloSet                  'Block Set
Public Mat() As Byte                        'Matrix
Public MatCop() As Byte                     'Matrix Copy
Public NexBlo(3) As Byte                    'Next Block
Public Tim As TTim                          'Timer
Public BloPre As Byte                       'Block Preview
Public CurBlo As TBlo                       'Current Block
Public CurBloSet As Byte                    'Current Block Set
Public CurColSet As Byte                    'Current Color Set
Public GamSta As Byte                       'Game State
                                                '0 = No Game Running
                                                '1 = Game Running
                                                '2 = Deletion Of Rows
                                                '3 = Dropping Of Rows
                                                '4 = Game Paused
                                                '5 = Game Over
Public GamSheBC As Long                     'Game Sheet Background Color
Public GamStaPas As Byte                    'Game State Pass
Public GapLes As Single                     'Gapless
Public MilSec As Single                     'Milliseconds
Public PlaFie As TPlaFie                    'Playing Field
Public Sta As TSta                          'Statistics
Public StaFie As TStaFie                    'Statistics Field
Public TimID As Long                        'Timer ID
