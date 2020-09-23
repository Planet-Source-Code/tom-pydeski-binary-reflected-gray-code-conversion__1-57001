VERSION 5.00
Begin VB.Form frmGray 
   Caption         =   "Tom Pydeski's Gray Code Conversion"
   ClientHeight    =   7395
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDecGray 
      Caption         =   "Decimal to Gray"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtDecimal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Text            =   "0"
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton cmdGrayDec 
      Caption         =   "Gray to Decimal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtGray 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4320
      TabIndex        =   0
      Text            =   "15"
      Top             =   480
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   6360
      Left            =   45
      TabIndex        =   4
      Top             =   1005
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Decimal Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Gray Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2805
      TabIndex        =   5
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmGray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim GrayCode As Integer
Dim ConvTable$
Dim Conversion$
Dim GraySum As Integer
Dim BitPos As Integer
Dim GrayBit As Integer
Dim BinArr(8) As Byte
Dim GrayArr(8) As Byte
Dim arryData() As String
'Example submitted by Tom Pydeski to convert Decimal to Binary Reflected Gray Code
'I did some poking around and found the following info, but could not find any
'examples in visual basic.  I took the C examples and converted them to VB
'
'http://mathworld.wolfram.com/GrayCode.html
'A Gray code is an encoding of numbers so that adjacent numbers have a single digit
'differing by 1. The term Gray code is often used to refer
' to a "reflected" code, or more specifically still, the binary reflected Gray code.
'
'To convert a binary d1,d2,...dn-1,dn number to its corresponding
'binary reflected Gray code, start at the right with the
'digit dn (the nth, or last, digit). If the dn-1 is 1, replace  by 1-dn-1;
'otherwise, leave it unchanged. Then proceed to dn-1 .
'Continue up to the first digit d1, which is kept the same since
'd0 is assumed to be a 0.
'The resulting number g1,g2,....gn-1,gn  is the reflected binary Gray code.
'
'To convert a binary reflected Gray code g1,g2,....gn-1,gn  to a binary number,
'start again with the nth digit, and compute
'sigma n = sum i=1 to n-1 gi (mod 2)
'
'If Sigma n is 1, replace gn by 1-gn ; otherwise, leave it the unchanged. Next compute
'sigma n-1 = sum i=1 to n-2 gi (mod 2)
'
'and so on. The resulting number d1,d2,...dn-1,dn  is the binary number corresponding
'to the initial binary reflected Gray code.
'
'The code is called reflected because it can be generated in the following manner.
'Take the Gray code 0, 1. Write it forwards, then backwards: 0, 1, 1, 0.
' Then prepend 0s to the first half and 1s to the second half: 00, 01, 11, 10.
'Continuing, write 00, 01, 11, 10, 10, 11, 01, 00 to obtain:
'000, 001, 011, 010, 110, 111, 101, 100, ...
'(Sloane's A014550 <http://www.research.att.com/projects/OEIS?Anum=A014550>).
'Each iteration therefore doubles the number of codes.
'
'0=0
'1=1
'2=11
'3=10
'4=110
'5=111
'6=101
'7=100
'8=1100
'9=1101
'10=1111
'11=1110
'12=1010
'13=1011
'14=1001
'15=1000
'
'Calculated                 From Table
'0  0   0           0   0
'1  (00000001)  1    (00000001)     1   1
'2  (00000010)  3    (00000011)     2   11
'3  (00000011)  2    (00000010)     3   10
'4  (00000100)  6    (00000110)     4   110
'5  (00000101)  7    (00000111)     5   111
'6  (00000110)  5    (00000101)     6   101
'7  (00000111)  4    (00000100)     7   100
'8  (00001000)  12   (00001100)     8   1100
'9  (00001001)  13   (00001101)     9   1101
'10 (00001010)  15   (00001111)     10  1111
'11 (00001011)  14   (00001110)     11  1110
'12 (00001100)  10   (00001010)     12  1010
'13 (00001101)  11   (00001011)     13  1011
'14 (00001110)  9    (00001001)     14  1001
'15 (00001111)  8    (00001000)     15  1000
'16 (00010000)  24   (00011000)     16  11000
'17 (00010001)  25   (00011001)     17  11001
'18 (00010010)  27   (00011011)     18  11011
'19 (00010011)  26   (00011010)     19  11010
'20 (00010100)  30   (00011110)     20  11110
'21 (00010101)  31   (00011111)     21  11111
'22 (00010110)  29   (00011101)     22  11101
'23 (00010111)  28   (00011100)     23  11100
'24 (00011000)  20   (00010100)     24  10100
'25 (00011001)  21   (00010101)     25  10101
'26 (00011010)  23   (00010111)     26  10111
'27 (00011011)  22   (00010110)     27  10110
'28 (00011100)  18   (00010010)     28  10010
'29 (00011101)  19   (00010011)     29  10011
'30 (00011110)  17   (00010001)     30  10001
'31 (00011111)  16   (00010000)     31  10000
'32 (00100000)  48   (00110000)     32  110000
'33 (00100001)  49   (00110001)     33  110001
'34 (00100010)  51   (00110011)     34  110011
'35 (00100011)  50   (00110010)     35  110010
'36 (00100100)  54   (00110110)     36  110110
'37 (00100101)  55   (00110111)     37  110111
'38 (00100110)  53   (00110101)     38  110101
'39 (00100111)  52   (00110100)     39  110100
'40 (00101000)  60   (00111100)     40  111100
'41 (00101001)  61   (00111101)     41  111101
'42 (00101010)  63   (00111111)     42  111111
'43 (00101011)  62   (00111110)     43  111110
'44 (00101100)  58   (00111010)     44  111010
'45 (00101101)  59   (00111011)     45  111011
'46 (00101110)  57   (00111001)     46  111001
'47 (00101111)  56   (00111000)     47  111000
'48 (00110000)  40   (00101000)     48  101000
'49 (00110001)  41   (00101001)     49  101001
'50 (00110010)  43   (00101011)     50  101011
'51 (00110011)  42   (00101010)     51  101010
'52 (00110100)  46   (00101110)     52  101110
'53 (00110101)  47   (00101111)     53  101111
'54 (00110110)  45   (00101101)     54  101101
'55 (00110111)  44   (00101100)     55  101100
'56 (00111000)  36   (00100100)     56  100100
'57 (00111001)  37   (00100101)     57  100101
'58 (00111010)  39   (00100111)     58  100111
'59 (00111011)  38   (00100110)     59  100110
'60 (00111100)  34   (00100010)
'61 (00111101)  35   (00100011)
'62 (00111110)  33   (00100001)
'63 (00111111)  32   (00100000)
'64 (01000000)  96   (01100000)
'65 (01000001)  97   (01100001)
'66 (01000010)  99   (01100011)
'67 (01000011)  98   (01100010)
'68 (01000100)  102  (01100110)
'69 (01000101)  103  (01100111)
'70 (01000110)  101  (01100101)
'71 (01000111)  100  (01100100)
'72 (01001000)  108  (01101100)
'73 (01001001)  109  (01101101)
'74 (01001010)  111  (01101111)
'75 (01001011)  110  (01101110)
'76 (01001100)  106  (01101010)
'77 (01001101)  107  (01101011)
'78 (01001110)  105  (01101001)
'79 (01001111)  104  (01101000)
'80 (01010000)  120  (01111000)
'81 (01010001)  121  (01111001)
'82 (01010010)  123  (01111011)
'83 (01010011)  122  (01111010)
'84 (01010100)  126  (01111110)
'85 (01010101)  127  (01111111)
'86 (01010110)  125  (01111101)
'87 (01010111)  124  (01111100)
'88 (01011000)  116  (01110100)
'89 (01011001)  117  (01110101)
'90 (01011010)  119  (01110111)
'91 (01011011)  118  (01110110)
'92 (01011100)  114  (01110010)
'93 (01011101)  115  (01110011)
'94 (01011110)  113  (01110001)
'95 (01011111)  112  (01110000)
'96 (01100000)  80   (01010000)
'97 (01100001)  81   (01010001)
'98 (01100010)  83   (01010011)
'99 (01100011)  82   (01010010)
'100    (01100100)  86   (01010110)
'101    (01100101)  87   (01010111)
'102    (01100110)  85   (01010101)
'103    (01100111)  84   (01010100)
'104    (01101000)  92   (01011100)
'105    (01101001)  93   (01011101)
'106    (01101010)  95   (01011111)
'107    (01101011)  94   (01011110)
'108    (01101100)  90   (01011010)
'109    (01101101)  91   (01011011)
'110    (01101110)  89   (01011001)
'111    (01101111)  88   (01011000)
'112    (01110000)  72   (01001000)
'113    (01110001)  73   (01001001)
'114    (01110010)  75   (01001011)
'115    (01110011)  74   (01001010)
'116    (01110100)  78   (01001110)
'117    (01110101)  79   (01001111)
'118    (01110110)  77   (01001101)
'119    (01110111)  76   (01001100)
'120    (01111000)  68   (01000100)
'121    (01111001)  69   (01000101)
'122    (01111010)  71   (01000111)
'123    (01111011)  70   (01000110)
'124    (01111100)  66   (01000010)
'125    (01111101)  67   (01000011)
'126    (01111110)  65   (01000001)
'127    (01111111)  64   (01000000)
'128    (10000000)  192  (11000000)
'129    (10000001)  193  (11000001)
'130    (10000010)  195  (11000011)
'131    (10000011)  194  (11000010)
'132    (10000100)  198  (11000110)
'133    (10000101)  199  (11000111)
'134    (10000110)  197  (11000101)
'135    (10000111)  196  (11000100)
'136    (10001000)  204  (11001100)
'137    (10001001)  205  (11001101)
'138    (10001010)  207  (11001111)
'139    (10001011)  206  (11001110)
'140    (10001100)  202  (11001010)
'141    (10001101)  203  (11001011)
'142    (10001110)  201  (11001001)
'143    (10001111)  200  (11001000)
'144    (10010000)  216  (11011000)
'145    (10010001)  217  (11011001)
'146    (10010010)  219  (11011011)
'147    (10010011)  218  (11011010)
'148    (10010100)  222  (11011110)
'149    (10010101)  223  (11011111)
'150    (10010110)  221  (11011101)
'151    (10010111)  220  (11011100)
'152    (10011000)  212  (11010100)
'153    (10011001)  213  (11010101)
'154    (10011010)  215  (11010111)
'155    (10011011)  214  (11010110)
'156    (10011100)  210  (11010010)
'157    (10011101)  211  (11010011)
'158    (10011110)  209  (11010001)
'159    (10011111)  208  (11010000)
'160    (10100000)  240  (11110000)
'161    (10100001)  241  (11110001)
'162    (10100010)  243  (11110011)
'163    (10100011)  242  (11110010)
'164    (10100100)  246  (11110110)
'165    (10100101)  247  (11110111)
'166    (10100110)  245  (11110101)
'167    (10100111)  244  (11110100)
'168    (10101000)  252  (11111100)
'169    (10101001)  253  (11111101)
'170    (10101010)  255  (11111111)
'171    (10101011)  254  (11111110)
'172    (10101100)  250  (11111010)
'173    (10101101)  251  (11111011)
'174    (10101110)  249  (11111001)
'175    (10101111)  248  (11111000)
'176    (10110000)  232  (11101000)
'177    (10110001)  233  (11101001)
'178    (10110010)  235  (11101011)
'179    (10110011)  234  (11101010)
'180    (10110100)  238  (11101110)
'181    (10110101)  239  (11101111)
'182    (10110110)  237  (11101101)
'183    (10110111)  236  (11101100)
'184    (10111000)  228  (11100100)
'185    (10111001)  229  (11100101)
'186    (10111010)  231  (11100111)
'187    (10111011)  230  (11100110)
'188    (10111100)  226  (11100010)
'189    (10111101)  227  (11100011)
'190    (10111110)  225  (11100001)
'191    (10111111)  224  (11100000)
'192    (11000000)  160  (10100000)
'193    (11000001)  161  (10100001)
'194    (11000010)  163  (10100011)
'195    (11000011)  162  (10100010)
'196    (11000100)  166  (10100110)
'197    (11000101)  167  (10100111)
'198    (11000110)  165  (10100101)
'199    (11000111)  164  (10100100)
'200    (11001000)  172  (10101100)
'201    (11001001)  173  (10101101)
'202    (11001010)  175  (10101111)
'203    (11001011)  174  (10101110)
'204    (11001100)  170  (10101010)
'205    (11001101)  171  (10101011)
'206    (11001110)  169  (10101001)
'207    (11001111)  168  (10101000)
'208    (11010000)  184  (10111000)
'209    (11010001)  185  (10111001)
'210    (11010010)  187  (10111011)
'211    (11010011)  186  (10111010)
'212    (11010100)  190  (10111110)
'213    (11010101)  191  (10111111)
'214    (11010110)  189  (10111101)
'215    (11010111)  188  (10111100)
'216    (11011000)  180  (10110100)
'217    (11011001)  181  (10110101)
'218    (11011010)  183  (10110111)
'219    (11011011)  182  (10110110)
'220    (11011100)  178  (10110010)
'221    (11011101)  179  (10110011)
'222    (11011110)  177  (10110001)
'223    (11011111)  176  (10110000)
'224    (11100000)  144  (10010000)
'225    (11100001)  145  (10010001)
'226    (11100010)  147  (10010011)
'227    (11100011)  146  (10010010)
'228    (11100100)  150  (10010110)
'229    (11100101)  151  (10010111)
'230    (11100110)  149  (10010101)
'231    (11100111)  148  (10010100)
'232    (11101000)  156  (10011100)
'233    (11101001)  157  (10011101)
'234    (11101010)  159  (10011111)
'235    (11101011)  158  (10011110)
'236    (11101100)  154  (10011010)
'237    (11101101)  155  (10011011)
'238    (11101110)  153  (10011001)
'239    (11101111)  152  (10011000)
'240    (11110000)  136  (10001000)
'241    (11110001)  137  (10001001)
'242    (11110010)  139  (10001011)
'243    (11110011)  138  (10001010)
'244    (11110100)  142  (10001110)
'245    (11110101)  143  (10001111)
'246    (11110110)  141  (10001101)
'247    (11110111)  140  (10001100)
'248    (11111000)  132  (10000100)
'249    (11111001)  133  (10000101)
'250    (11111010)  135  (10000111)
'251    (11111011)  134  (10000110)
'252    (11111100)  130  (10000010)
'253    (11111101)  131  (10000011)
'254    (11111110)  129  (10000001)
'255    (11111111)  128  (10000000)
'
'

Private Sub Form_Load()
Debug.Print (DectoGray(57)), ConBin(DectoGray(57))
'let's build a table for all values so we can compare
For i = 1 To 255
    'simple way to convert decimal to GrayCode
    GrayCode = i Xor (i \ 2)
    Conversion$ = i & vbTab & "(" & ConBin(i) & ") " & vbTab & GrayCode & vbTab & " (" & ConBin(GrayCode) & ")"
    ConvTable$ = ConvTable$ & Conversion$ & vbCrLf
    List1.AddItem Conversion$
Next i
'Debug.Print ConvTable$
Clipboard.Clear
Clipboard.SetText ConvTable$
cmdGrayDec_Click
List1.TopIndex = 0
End Sub

Public Function DectoGray(DecimalIn As Integer) As Integer
'To convert binary to Gray, it is only necessary to XOR the original
'unsigned binary with a copy of itself that has been right shifted one place.
DectoGray = DecimalIn Xor (DecimalIn \ 2)
End Function

Private Sub cmdGrayDec_Click()
'convert the value in txtgray to decimal
txtDecimal = ""
DoEvents
Refresh
txtDecimal = GraytoDec(txtGray)
'highlight the pair in the list for comparison
For i = 0 To List1.ListCount - 1
    'the data in the list is tab delimited
    arryData() = Split(List1.List(i), vbTab)
    If txtGray = arryData(2) Then
        List1.ListIndex = i
        List1.TopIndex = i - 1
        Exit For
    End If
Next i
End Sub

Private Sub cmdDecGray_Click()
'convert the value in txtDecimal to Gray Code
txtGray = ""
DoEvents
Refresh
txtGray = DectoGray(txtDecimal)
'highlight the pair in the list for comparison
For i = 0 To List1.ListCount - 1
    'the data in the list is tab delimited
    arryData() = Split(List1.List(i), vbTab)
    If txtDecimal = arryData(0) Then
        List1.ListIndex = i
        List1.TopIndex = i - 1
        Exit For
    End If
Next i
End Sub

Public Function GraytoDec(GrayIn As Integer) As Integer
'from http://www.dspguru.com/comp.dsp/tricks/alg/grayconv.htm
'The cannonical way to convert Gray to binary is to XOR the bits one at a time,
'starting with the two highest bits, using the newly calculated bit in the next XOR.
GraytoDec = 0
'initialize binary array
For i = 7 To 0 Step -1
    GrayArr(i) = ((2 ^ i) And GrayIn) / (2 ^ i)
    Debug.Print GrayArr(i);
Next i
Debug.Print "=",
BinArr(7) = GrayArr(7)
For i = 6 To 0 Step -1
    'exclusive or the previous bit with the current bit
    'this becomes the current output bit
     BinArr(i) = BinArr(i + 1) Xor GrayArr(i)
Next i
'convert binary to decimal
For i = 7 To 0 Step -1
    GraytoDec = GraytoDec + (BinArr(i) * (2 ^ i))
    Debug.Print BinArr(i);
Next i
Debug.Print " = ";
End Function

Function ConBin(bNum As Integer) As String
'convert to binary string
Dim X As Long
Dim bIn As String
X = CLng(bNum)
bIn = ""
Do
    bIn = (X And 1) & bIn
    X = X \ 2
Loop While X
Do Until Len(bIn) > 7
    bIn = 0 & bIn
Loop
ConBin = bIn
End Function

Private Sub List1_Click()
'get the data from the list and put it in the
'appropriate text boxes for display
arryData() = Split(List1.Text, vbTab)
txtDecimal = arryData(0)
txtGray = arryData(2)
End Sub

Private Sub txtDecimal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdDecGray_Click
End If
End Sub

Private Sub txtGray_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrayDec_Click
End If
End Sub

