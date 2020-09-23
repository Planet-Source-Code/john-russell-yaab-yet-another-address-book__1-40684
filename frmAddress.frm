VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAddress 
   Caption         =   "Address Book"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12945
   Icon            =   "frmAddress.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   12945
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":19DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   7905
      Width           =   12945
      _ExtentX        =   22834
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14605
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignm Heg       =   5
            Alignm Heg       =   5
     5605
omeStreet
        End If
  mF   5ntabject.Width           =   14605
  "vd If
  mF   5ntabject.Width         If
  mF   5ntaD1-1a        =   ""
         EndProperty
      EndProp   rop
  eP   EndPropertyoa.rop
      EndProp   rop
  eP   EndPropertyoa.rop
   s"vd If
  mF   5ntabject.Width         If
  mF     DstatusBar   ndP92t=   11    D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB4D  IB6b1ga"AddressloThen  32
      MaskColor     op   rop
  eP   EndProo eP     IB4s 
      vb        =  c4S6
            tbContac #pertyoa.rop
   s"vd If
  mF   5ntl  DstatusBar IBg "1ga"Addres5O
      Ali,irlRStyle           =   5
        "vd If
  mF   5ntl  DstatusBar IBg "1ga"Addr IBrrop
   sBg "1ga"           Height       horop
   sf
            AutoSize        =   1
        =lor2i     Td "1ga"Addr IB}}eCd If
  mF   5ntl  DstatusBar IBg e.Widt   NumPanels       =   4
         BeginProperty Panel1 {8E3aneloperty Panel1 Ruu"loperty Panel1 Ruu"loperty anels g    =  nel1 Ruu"loperty ananels g    =  nel1 Ruu"loperty ananels g      rop
  ePCaE3ane  495
      Left            =   0anel1 Ruu"loperty an=   sECT ContactID, LastName, FirstNalopertBtoCT Coty an=   sECT ContactID, LastNa  "frmAddress.frx":0CE6
        ai 4
!anel1 Ruu"loperty anels g    =  nel1 Ruu"loperty t Ruu"loperty an=   sECT CoAd     i     Td "1gaels i 4elsentY        l Image6 e As Recordset
Dim rsCallType As Recordset

Dimperty
         BeginProperty Panel3 {8E3867AB-  "vd If
  mF   5ntl  DstatusBar IBg "1ga"Addr IBrrop
   sBg "1ga"           HaDstatusBar IBg "1ga"Addr IBrrop
   sBg "1ga"           HaDstatusBar IBg "1ga"AC HaDstatusBR   s       =   4
         B-00Cp
   sBg "1ga"       As'  mF   5ntl  DstatusBar IBg "1ga"Addr IBrrop
   sBg "1ga"8E3867AB-  "vd I    4s      HaDstat    nd v atusTI    4s     rop t'"1ga"Addr v atusTI    4s     r4ga"8E3867AB-  "vd I    4s      HaDstat    ndAa"A   4s      HaDstat    ndAa"A   4s      HaDstat    ndAa"A    4sBRtatusBar IBg "1ga"AC HaDstatusBR   s       =tsTI    4s     r4ga"8E3867AB-  "       sEIdneN7AB8 4s      HaDstat    ndAa"A    4sBRtatusBar2cAs Recordset

Dimperty
         BeginProperty Panel3 {8E3867AB-  "vd If
  mF   5ntl  DstatusBar IBg "1ga"Addr IBrrop
   sBg "1ga"           HaDstatusB-c     =  a      HaDstatuEb586-ordset

DimpertactID, LasAdd11g,T CoAd     i     Td "1gaels i 4elsentY        l Image6 e As Recordset
Dim rsCallType As Reset
Dim rsCallT33 {8E3867AB-  "vd If
 llTdIBrrop
   sBg "1gas mF   5ntl tatet
Dstatge64d-V"
r  s0.Tab = 0              '-- make the 1 mF   5ntl r IBg "1ga"AC HaDsxxwlEwlEwi    a    Alignm Heg       =   5
       d If
 llTdI4cC     Begi,T CoAd     i     Td "1gaels i 4elsentY    ed
Dim rsCallType A9gaels 0A  4elsentY    ed
Dim renD.  IfPanel1 Ruu"loperty 4DADNn         =   "Delete"
            ImageIndex      =   4
  T2e"
            ImageInd "1ga"Addr IB}}eCd If
k" 5
            Alignm Heg      }tBRuu"loperty an=   sECT CoAd     i      entY    ed
Dim renD.  IfPanel1 Ruu"lopert   ed
Dim renD.  If renD.  IfPanel1 Ru
Dy
         BeginProperty ListImage3 {2C247F27-8591-11D1- renD.  If renD.  IfPanel1 Ru
Dl   ed
Dim r   ba"
       tdroperty ListImaod IBg "1fe    td
!5.  If renD.  IfPanel1 2    =   1200
"loperty an=   sEmi0 rep3 {2C24pert IfPanel1 Ru
Dl   ed41fe 4    ed24pert IfPanel1 Ru
Dl        Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture       22222222222222222222222222222lete"
   wa   - renD.  I1222222Zture 2 (   1s2122f            Picture  irthDay, 4, 6)
        End If
el2 C - renD.  I12229lfPanel1 Ru
Dl        Key  False
            Toolbar1.Buttons(bEdit).ey                    ToolbdEi7fPanel1 Ru
DF ndP    Width                6)
        End If
el2 C - renD.  I12229lfP
            Tool  irthDals gOlign F ndP    Width     ane    WC  I12229lfP
            Tool  irthDalID, L  WC  Il'EndPropertyy         I12229lfP
            Tool  irthDalID, L  WC  Il'EndPropertyy         I12229lfP
            Tool  irthDalIL              =   5
        "vd If
  mF   5ntl  DstatusBar Ir(1B04 {2statge64d-V"
r  s0.Tab = 0    NB 49lfP
         d     I12229hDasBg  L  WC  Il'End        9lfP
r          Picture  ia"8En22222222lete"ool  nle    gehDalllllllllllllll8E3867AB-  "vd I    4s      HaDstat    ndAa"A  a     =   5
        "vd If
  mF   5ntl  DstatusBar Ir(1B04 {2statge64d-V"
r  s0.Tab = 0    NB 49lfP
         d     I12229hDasBg  L  WC  Ie2229hDasBg  L  WC  Ie2229hDasBgyy       ge4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
  s   5
   04 {2stat     9l m r   ba"
   } 
  s   5
   04 {28turg  "
   } 
  s   5XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
   04 {28turontactID, LastName,108     9l m r ertBtoCT Coty an=   553 ge4 {2C2u"lopert  I12229hDasfP
   fP
   DasBg  L  WC   _ExtentY        = l'Eabase     s
   04 {28tu  Ie2229hDasBgyy    193 ge4 {2C247F27-8591-11D1-B16A-00-7476F0283628} 
  s   5
   04 {2sta36   04 {28turontactID, LastName,372     9l m r ertBtoCT Coty an=   565 ge4 {2C247F12229hDasd-V"   IliskReadsDasBg  L  WC47F1867erl1 Ruu"lop    12staixed SingleasBg  L  WC47F  Ie2229hDasBgyy     5 ge4 {2C247F a"Addr IBrrop
   sBg "32     9l m r  
  s   5
   04 {2sta4n MSComctlLib.SrontactID, LastName,48     9l m r  
        EndProperty
 1 ge4 {2C247Fu"lopert  I 7F12229hDasd-V"   IliskWriIl'DasBg  L  WC47F1867erl1 Ruu"lop    12staixed SingleasBg  L  WC47F  Ie2229hDasBgyy     5 ge4 {2C247F a"Addr IBrrop
   sBg "32     9l m r  
  s   5
   04 {2sta47 MSComctlLib.SrontactID, LastName,96     9l m r  
        EndProperty
 1 ge4 {2C247Fu"lopert  I 7F12229hDasd-V"   IReadCacheDasBg  L  WC47F1867erl1 Ruu"lop    12staixed SingleasBg  L  WC47F  Ie2229hDasBgyy     5 ge4 {2C247F a"Addr IBrrop
   sBg "32     9l m r  
  s   5
   04 {2sta46 MSComctlLib.SrontactID, LastName,144     9l m r  
        EndProperty
 1 ge4 {2C247Fu"lopert  I 7F12229hDasd-V"   IReadAheadDasBg  L  WC47F1867erl1 Ruu"lop    12staixed SingleasBg  L  WC47F  Ie2229hDasBgyy     5 ge4 {2C247F a"Addr IBrrop
   sBg 4  =   "frmAAAAAAA  s   5
   04 {2sta45 MSComctlLib.SrontactID, LastName,48     9l m r  
        EndProperty
 1 ge4 {2C247Fu"lopert  I 7F12229hDasd-V"   ILocksPlacedDasBg  L  WC47F1867erl1 Ruu"lop    12staixed SingleasBg  L  WC47F  Ie2229hDasBgyy     5 ge4 {2C247F a"Addr IBrrop
   sBg 4  =   "frmAAAAAAA  s   5
   04 {2sta44 MSComctlLib.SrontactID, LastName,96     9l m r  
        EndProperty
 1 ge4 {2C247Fu"lopert  I 7F12229hDasd-V"   IReleaseLocksDasBg  L  WC47F1867erl1 Ruu"lop    12staixed SingleasBg  L  WC47F  Ie2229hDasBgyy     5 ge4 {2C247F a"Addr IBrrop
   sBg 4  =   "frmAAAAAAA  s   5
   04 {2sta43 MSComctlLib.SrontactID, LastName,144     9l m r  
        EndProperty
 1 ge4 {2C247Fu"lopert  I 7F12229hDasd-V" d-V"
7DasBg  L  WC47F   _ExtentY        = lisk Reads"asBg  L  WC47FFore:0894
         &H00FF3932&asBg  L  WC47F  Ie2229hDasBgyy     5 ge4 {2C247F a"Addr IBrrop
   sBg 24     9l m r  
  s   5
   04 {2sta42    9l m r  
 ontactID, LastName,48     9l m r  
        EndProperty97 ge4 {2C247Fu"lopert  I 7F12229hDasd-V" d-V"
8DasBg  L  WC47F   _ExtentY        = lisk WriIl'"asBg  L  WC47FFore:0894
         &H00FF3932&asBg  L  WC47F  Ie2229hDasBgyy     5 ge4 {2C247F a"Addr IBrrop
   sBg 24     9l m r  
  s   5
   04 {2sta4
            tbontactID, LastName,96     9l m r  
        EndProperty97 ge4 {2C247Fu"lopert  I 7F12229hDasd-V" d-V"
9DasBg  L  WC47F   _ExtentY        = Read Cache"asBg  L  WC47FFore:0894
         &H00FF3932&asBg  L  WC47F  Ie2229hDasBgyy     5 ge4 {2C247F a"Addr IBrrop
   sBg 24     9l m r  
  s   5
   04 {2sta4     9l m r  
 ontactID, LastName,144     9l m r  
        EndProperty97 ge4 {2C247Fu"lopert  I 7F12229hDasd-V" d-V"20DasBg  L  WC47F   _ExtentY        = Read Ahead"asBg  L  WC47FFore:0894
         &H00FF3932&asBg  L  WC47F  Ie2229hDasBgyy     5 ge4 {2C247F a"Addr IBrrop
   sBg 288     9l m r  
  s   5
   04 {2sta39    9l m r  
 ontactID, LastName,48     9l m r  
        EndProperty
 1 ge4 {2C247Fu"lopert  I 7F12229hDasd-V" d-V"21DasBg  L  WC47F   _ExtentY        = LocksDPlaced"asBg  L  WC47FFore:0894
         &H00FF3932&asBg  L  WC47F  Ie2229hDasBgyy     5 ge4 {2C247F a"Addr IBrrop
   sBg 288     9l m r  
  s   5
   04 {2sta3n MSComctlLib.SrontactID, LastName,96     9l m r  
        EndProperty
 1 ge4 {2C247Fu"lopert  I 7F12229hDasd-V" d-V"22DasBg  L  WC47F   _ExtentY        = LocksDReleased"asBg  L  WC47FFore:0894
         &H00FF3932&asBg  L  WC47F  Ie2229hDasBgyy     5 ge4 {2C247F a"Addr IBrrop
   sBg 288     9l m r  
  s   5
   04 {2sta37 MSComctlLib.SrontactID, LastName,144     9l m r  
        EndProperty
 1 ge4 {2C247Fu"lopert  Iu"lopert  I12229hDasBg  L  WC    Il'DasBg  L  WC  Ie2229hDasBgyy     29 ge4 {2C247F27-8591-11D1-B16A-00-7464     9l m r Locked -11D1-B16A-00-BeginProperty PaaaaMultity P1D1-B16A-00-BeginProperty PaaaaScrollBarsD1-B16A-003egiBothperty Paaaa  s   5
   04 {2sta29    9l m r rontactID, LastName,372     9l m r ertBtoCT Coty an=   805 ge4 {2C2u"lopert  I12229htyle        _Ext   5       DasBg  L  WC  Ie2229hDasBgyy     89 ge4 {2C247F27-8591-11D1-B16A-00-7464     9l m r   s   5
   04 {2sta28    9l m r rontactID, LastName,60     9l m r ertBtoCT Coty an=   805 ge4 {2C2 ctID, LastNa  "frmAddre4  8ge4 {2C2 ctID, Las   ai 4
!ane5106 MSComctlLi   52sta2ge4 {2C247F2-V"Wrap2sta-BeginProperty PaaaaHideSelec_Extent2sta-BeginProperty Paaaa       =   4
         BeginProper 7FFore:0894
         F27-8591-1     9l m r ty ListImage2 {2C247F27-8591-11D1-B16A- 7F1867erl1 Ruu"lop    1D1-B16A- 7Ferty anels g    =  nel1 RuuuuuuuNumItemomeStree =  n     9l mu"lopert  I12229hDasBg  L  WC  , 4, 6)
DasBg  L  WC  Ie2229hDasBgyy     8 ge4 {2C247F27-8591-11D1-B16A-0024     9l m r MaxLengBtoCT Cot!ane5     9l m r   s   5
   04 {2sta1     9l m r   ba"
   } 
  s   5
   04 {28turg  "
   } 
  s   5XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
   04 {28turontactID, LastName,492     9l m r ertBtoCT Coty an=   469 ge4 {2C2u"lopert  I12229hDasBg  L  WC  , 4,22222DasBg  L  WC  Ie2229hDasBgyy     8 ge4 {2C247F27-8591-11D1-B16A-004  =   "frmAAAAMaxLengBtoCT Cot!ane2ge4 {2C247F  s   5
   04 {2sta9    9l m r r ba"
   } 
  s   5
   04 {28turg  "
   } 
  s   5XX
   04 {28turontactID, LastName,312     9l m r ertBtoCT Coty an=   PCaE3ane  4u"lopert  I12229hDasBg  L  WC  , 4,ge4 DasBg  L  WC  Ie2229hDasBgyy     8 ge4 {2C247F27-8591-11D1-B16A-00264     9l m r MaxLengBtoCT Cot!ane12ge4 {2C247F  s   5
   04 {2sta8    9l m r r ba"
   } 
  s   5
   04 {28turg  "
   } 
  s   5XXXXXXXXXXXX
   04 {28turontactID, LastName,312     9l m r ertBtoCT Coty an=   145 ge4 {2C2u"lopert  I12229hDasBg  L  WC  , 4,22 RuDasBg  L  WC  Ie2229hDasBgyy     8 ge4 {2C247F27-8591-11D1-B16A-0024     9l m r MaxLengBtoCT Cot!ane2     9l m r   s   5
   04 {2sta7    9l m r   ba"
   } 
  s   5
   04 {28turg  "
   } 
  s   5XXXXXXXXXXXXXXXXXXXX
   04 {28turontactID, LastName,312     9l m r ertBtoCT Coty an=    29 ge4 {2C2u"lopert  I12229hDasBg  L  WC    =   12DasBg  L  WC  Ie2229hDasBgyy     8 ge4 {2C247F27-8591-11D1-B16A-0036 =   "frmAAAAMaxLengBtoCT Cot!ane2     9l m r   s   5
   04 {2sta6   04 {28tur ba"
   } 
  s   5
   04 {28turg  "
   } 
  s   5XXXXXXXXXXXXXXXXXXXX
   04 {28turontactID, LastName,204     9l m r ertBtoCT Coty an=    29 ge4 {2C2u"lopert  I12229hDasBg  L  WC  Ru
Dl   ed
DasBg  L  WC  Ie2229hDasBgyy     8 ge4 {2C247F27-8591-11D1-B16A-00264     9l m r MaxLengBtoCT Cot!ane1    9l m r   s   5
   04 {2sta5    9l m r   ba"
   } 
  s   5
   04 {28turg  "
   } 
  s   5X
   04 {28turontactID, LastName,204     9l m r ertBtoCT Coty an=    5 ge4 {2C2u"lopert  I12229hDasBg  L  WC  l1 Ru
DyDasBg  L  WC  Ie2229hDasBgyy     8 ge4 {2C247F27-8591-11D1-B16A-0024     9l m r MaxLengBtoCT Cot!ane15    9l m r   s   5
   04 {2sta4    9l m r   ba"
   } 
  s   5
   04 {28turg  "
   } 
  s   5XXXXXXXXXXXXXXX
   04 {28turontactID, LastName,204     9l m r ertBtoCT Coty an=   169 ge4 {2C2u"lopert  I12229hMSfrx".frx"EdL  WT CoAd     DasBg  L  WC  Ie2229hDasBgyy     8 ge4 {2C247F27-8591-11D1-B16A-0024     9l m r   s   5
   04 {2sta11    9l m r   ba"
   } 
  s   5
   04 {28turdress.frx":0442
   52     9l m r ertBtoCT Coty an=   
 1 ge4 {2C247FinProperty Panel3 {8E111D1-B16A- 7FID, Las   ai 4
!ane501D1-B16A- 7FI anels g    =  nel1 Ruu"loperty t r MaxLengBtoCT Cot!ane1     9l m r Mask            s   5##/##/####
   04 {28tuPromptChar      s   5_
   04 {2u"lopert  I12229hMSfrx".frx"EdL  WT C    ImageInd DasBg  L  WC  Ie2229hDasBgyy     8 ge4 {2C247F27-8591-11D1-B16A-00432     9l m r   s   5
   04 {2sta12ge4 {2C247F  ba"
   } 
  s   5
   04 {28turdress.frx":0442
  408     9l m r ertBtoCT Coty an=   157 ge4 {2C247FinProperty Panel3 {8E778ge4 {2C2 ctID, Las   ai 4
!ane501D1-B16A- 7FI anels g    =  nel1 Ruu"loperty t r PromptInclude=  nel10 2statge64d-V"
r r MaxLengBtoCT Cot!ane15    9l m r Mask            s   5999999999999999
   04 {28tuPromptChar      s   5_
   04 {2u"lopert  I12229hMSfrx".frx"EdL  WT C    FaxDasBg  L  WC  Ie2229hDasBgyy     8 ge4 {2C247F27-8591-11D1-B16A-00228     9l m r   s   5
   04 {2sta13    9l m r   ba"
   } 
  s   5
   04 {28turdress.frx":0442
  408     9l m r ertBtoCT Coty an=   157 ge4 {2C247FinProperty Panel3 {8E778ge4 {2C2 ctID, Las   ai 4
!ane501D1-B16A- 7FI anels g    =  nel1 Ruu"loperty t r PromptInclude=  nel10 2statge64d-V"
r r MaxLengBt    =  nel115    9l m r Mask            s   5999999999999999
   04 {28tuPromptChar      s   5_
   04 {2u"lopert  I12229hMSfrx".frx"EdL  WT C    eInd DasBg  L  WC  Ie2229hDasBgyy     8 ge4 {2C247F27-8591-11D1-B16A-0024     9l m r   s   5
   04 {2sta14    9l m r   ba"
   } 
  s   5
   04 {28turdress.frx":0442
  408     9l m r ertBtoCT Coty an=   157 ge4 {2C247FinProperty Panel3 {8E778ge4 {2C2 ctID, Las   ai 4
!ane501D1-B16A- 7FI anels g    =  nel1 Ruu"loperty t r ClipMod       If
  mF   5ntaD1-PromptInclude=  nel10 2statge64d-V"
r r PromptChar      s   5_
   04 {2u"lopert  I12229hMSfrx".frx"EdL  WT C    ZipDasBg  L  WC  Ie2229hDasBgyy     8 ge4 {2C247F27-8591-11D1-B16A-00468     9l m r   s   5
   04 {2sta15    9l m r   ba"
   } 
  s   5
   04 {28turontactID, LastName,312     9l m r ertBtoCT Coty an=   
 1 ge4 {2C247FinProperty Panel3 {8E111D1-B16A- 7FID, Las   ai 4
!ane501D1-B16A- 7FI anels g    =  nel1 Ruu"loperty t r MaxLengBtoCT Cot!ane1     9l m r Mask            s   5#####-####
   04 {28tuPromptChar      s   5_
   04 {2u"lopert  I12229hDasd-V" d-V"
DasBg  L  WCWidth         If
  -BeginProperty Paaaa   _ExtentY        =  Toolbar1.Button tu  Ie2229hDasBgyy    19 ge4 {2C247F27-8591-11D1-B16A-0024     9l m r   s   5
   04 {2sta49    9l m r r ba"
   } 
  s   5C To
   04 {28turontactID, LastName,72     9l m r ertBtoCT Coty an=   66     9l mu"lopert  I12229hDasd-V"   Il'End      DasBg  L  WC1867erl1 Ruu"lop    12staixed SingleasBg  L  WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-00-7308     9l m r   s   5
   04 {2sta3 ge4 {2C247FrontactID, LastName,108     9l m r ertBtoCT Coty an=   361 ge4 {2C2u"lopert  I12229hDasd-V"   I2222222leteDasBg  L  WC1867erl1 Ruu"lop    12staixed SingleasBg  L  WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-00-7308     9l m r   s   5
   04 {2sta34ge4 {2C247FrontactID, LastName,168     9l m r ertBtoCT Coty an=   361 ge4 {2C2u"lopert  I12229hDasd-V"   I8E3867AB-  DasBg  L  WC1867erl1 Ruu"lop    12staixed SingleasBg  L  WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-00-7308     9l m r   s   5
   04 {2sta33ge4 {2C247FrontactID, LastName,228     9l m r ertBtoCT Coty an=   361 ge4 {2C2u"lopert  I12229hDasd-V" a"A  a DasBg  L  WC   _ExtentY        = l'Eabase d        9lfP
r  7FFore:0894
         &H00FF3932&asBg  L  WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-00-7476F0283628} 
  s   5
   04 {2sta32ge4 {2C247F ontactID, LastName,108     9l m r ertBtoCT Coty an=   145 ge4 {2C2u"lopert  I12229hDasd-V" a"A  a5DasBg  L  WC   _ExtentY        = 2222 222lete"ool  nl 7FFore:0894
         &H00FF3932&asBg  L  WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-00-7476F0283628} 
  s   5
   04 {2sta31    9l m r  ontactID, LastName,168     9l m r ertBtoCT Coty an=   157 ge4 {2C2u"lopert  I12229hDasd-V" a"A  a6DasBg  L  WC   _ExtentY        = NB 4u"lo8E3867s"ool  nl 7FFore:0894
         &H00FF3932&asBg  L  WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-00-7476F0283628} 
  s   5
   04 {2sta3F0283628} 
 ontactID, LastName,228     9l m r ertBtoCT Coty an=   157 ge4 {2C2u"lopert  I12229hDasd-V" a"A  a0DasBg  L  WC   _ExtentY        = wi    a    A WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-00468     9l m r   s   5
   04 {2sta270283628} 
 ontactID, LastName,276F0283628} 
ertBtoCT Coty an=   
 1 ge4 {2C2u"lopert  I12229hDasd-V" a"A  8DasBg  L  WC   _ExtentY        = ge4 {2C247F27 WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-00264     9l m r   s   5
   04 {2sta26   04 {28turontactID, LastName,276F0283628} 
ertBtoCT Coty an=   
45 ge4 {2C2u"lopert  I12229hDasd-V" a"A  7DasBg  L  WC   _ExtentY        =      l1 Ru
Dl       WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-0024     9l m r   s   5
   04 {2sta2 ge4 {2C247FrontactID, LastName,276F0283628} 
ertBtoCT Coty an=    29 ge4 {2C2u"lopert  I12229hDasd-V" a"A  a3DasBg  L  WC   _ExtentY        = 2222   1200
"lope WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-0036 =   "frmAAAA  s   5
   04 {2sta24ge4 {2C247FrontactID, LastName,180     9l m r ertBtoCT Coty an=    29 ge4 {2C2u"lopert  I12229hDasd-V" a"A  a2DasBg  L  WC   _ExtentY        = M. I.00
"lope WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-00264     9l m r   s   5
   04 {2sta23ge4 {2C247FrontactID, LastName,180     9l m r ertBtoCT Coty an=   PCaE3ane  4u"lopert  I12229hDasd-V" a"A  a DasBg  L  WC   _ExtentY        = l1 Ru   1200
"lope WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-0024     9l m r   s   5
   04 {2sta22ge4 {2C247F ontactID, LastName,180     9l m r ertBtoCT Coty an=   169 ge4 {2C2u"lopert  I12229hDasd-V" a"A  9DasBg  L  WC   _ExtentY        = 222222222lete" WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-004  =   "frmAAAA  s   5
   04 {2sta21    9l m r  ontactID, LastName,276F0283628} 
ertBtoCT Coty an=   49 ge4 {2C2u"lopert  I12229hDasd-V" a"A  6DasBg  L  WC   _ExtentY        = Nell eInd "1ga"Addr WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-00432     9l m r   s   5
   04 {2sta2     9l m r  ontactID, LastName,384     9l m r ertBtoCT Coty an=   157 ge4 {2C2u"lopert  I12229hDasd-V" a"A  5DasBg  L  WC   _ExtentY        =      perty 4DADNn WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-00228     9l m r   s   5
   04 {2sta19    9l m r rontactID, LastName,384     9l m r ertBtoCT Coty an=   157 ge4 {2C2u"lopert  I12229hDasd-V" a"A   DasBg  L  WC   _ExtentY        =      eInd "1ga"Addr WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-0024     9l m r   s   5
   04 {2sta18    9l m r rontactID, LastName,384     9l m r ertBtoCT Coty an=   157 ge4 {2C2u"lopert  I12229hDasd-V" a"A  3DasBg  L  WC   _ExtentY        =       6)
D5
    "1ga"Addr WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-0024     9l m r   s   5
   04 {2sta170283628} 
 ontactID, LastName,468     9l m r ertBtoCT Coty an=   3 5 ge4 {2C2u"lopert  I12229hDasd-V"   IoAd     DasBg  L  WC1867erl1 Ruu"lop    12staixed SingleasBg  L  WC  Ie2229hDasBgyy     5 ge4 {2C247F27-8591-11D1-B16A-00192     9l m r   s   5
   04 {2sta"loperty t r   ba"
   } 
  s   5
   04 {28turontactID, LastName, 52     9l m r ertBtoCT Coty an=   301 ge4 {2C2u"lopertu"lopert12229htyle        Toolbar Toolbar DasBg  L   AutoSize        = 12st  AutoronasBg  L   Ie2229hDasBgyy    81     9l m27-8591-11D1-B16A-00     9l m  s   5
   04 {2sta     9l m ontactID, LastName,     9l mertBtoCT Coty an=   
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F028361429    9l mButtonertBtoCT C28361323ge4 {2C2Button  Ie2229hD28361371 Ruu"loperty anels g    =  nel1 Ru nel1 Ruu"loperty Panelel1 Ru neD1-B16A-00erty Panel"C0F0283628
   04 {2I anels g    =  nel1 Ruu"loperty t12229  IB4D  IButtons {66833FE8ress3oThen  32
      MaskColor     op  NumButtons   =  nel112ge4 {2C247F12229  IB4D  IButton1 {66833FEAress3oThen  32
      MaskColor     op        _ExtentY        =   1 New"or     op     C0F02   5
   04anelel1 Ru ne rtu"l  IB4D  IB4D  IB4D  IB4D  IB4D  IButton2 {66833FEAress3oThen  32
      MaskColor     op     l1 Ruu"loperty Panel3l1 Ru ne rtu"l  IB4D  IB4D  IB4D  IB4D  IB4D  IButton3 {66833FEAress3oThen  32
      MaskColor     op        _ExtentY        = Nnelsl"or     op     C0F02   5
   04anel2l1 Ru ne rtu"l  IB4D  IB4D  IB4D  IB4D  IB4D  IButton4 {66833FEAress3oThen  32
      MaskColor     op     l1 Ruu"loperty Panel3l1 Ru ne rtu"l  IB4D  IB4D  IB4D  IB4D  IB4D  IButton5 {66833FEAress3oThen  32
      MaskColor     op        _ExtentY        = 2ave"asBg  L  WC47FC0F02   5
   04anel3l1 Ru ne rtu"l  IB4D  IB4D  IB4D  IB4D  IB4D  IButton6 {66833FEAress3oThen  32
      MaskColor     op     l1 Ruu"loperty Panel3l1 Ru ne rtu"l  IB4D  IB4D  IB4D  IB4D  IB4D  IButton7 {66833FEAress3oThen  32
      MaskColor     op        _ExtentY        = lelete"asBg  L  WC47FC0F02   5
   04anel4l1 Ru ne rtu"l  IB4D  IB4D  IB4D  IB4D  IB4D  IButton8 {66833FEAress3oThen  32
      MaskColor     op     l1 Ruu"loperty Panel3l1 Ru ne rtu"l  IB4D  IB4D  IB4D  IB4D  IB4D  IButton9 {66833FEAress3oThen  32
      MaskColor     op        _ExtentY        = Edit"asBg  L  WC47FC0F02   5
   04anel5l1 Ru ne rtu"l  IB4D  IB4D  IB4D  IB4D  IB4D  IButton10 {66833FEAress3oThen  32
      MaskColor     op     l1 Ruu"loperty Panel3l1 Ru ne rtu"l  IB4D  IB4D  IB4D  IB4D  IB4D  IButton11 {66833FEAress3oThen  32
      MaskColor     op        _ExtentY        = Quit"asBg  L  WC47FC0F02   5
   04anel6l1 Ru ne rtu"l  IB4D  IB4D  IB4D  IB4D  IB4D  IButton12 {66833FEAress3oThen  32
      MaskColor     op     l1 Ruu"loperty Panel3l1 Ru ne rtu"l  IB4D  IB4D  IBu"l  IB4D  IB4D u"lopert12229hDasMenu mnuPopupDasBg  L    _ExtentY        = PopupDMenu"asBg  L VisibRuu"lopertyame,  2statge64d-V"
r12229hDasMenu mnu  1NewDasBg  L  WC   _ExtentY        = &  1 New"or     ou"lopert  I12229hDasMenu mnuleleteDasBg  L  WC   _ExtentY        = &lelete"asBg  L u"lopert  I12229hDasMenu mnu Toolba
DyDasBg  L  WC   _ExtentY        = By Ie2229hD  1200
"lopeu"lopert  I12229hDasMenu Mnu  =   12DasBg  L  WC   _ExtentY        = By   = n 12,l1 Ru
Dy
       u"lopertu"lopu"lopAttributeDVB_  12D= =   5
    "opAttributeDVB_Global  12Spac2D= atge64dAttributeDVB_d    r IrD= atge64dAttributeDVB_PredeclaredIdD= nPropeAttributeDVB_ExposedD= atge64dO _ExteExplicit

Dim cB 4u"lNod  As Nod  ' usedDto popul222D1 Ru cB 4rol
Dim rs  Il'Tr IrDAso8E3867set
Dim rs    TyprDAso8E3867set

Dim iCurr-B122222DAsoInteger ' curr-B1 s2222Dof program, ie. adding, editing etc...
Dim lCurr-B1NB 4u"l3628AsoLong ' unique IDDof cB 4u"l
Dim sCurr-B1NB 4u"l  12DAsoString ' cB 4u"l's n 12
Dim bFieldsPopul222dDAsoBooleaten'fl bato sRu if the fields hac #d  r and whether they should bu cleared
Dim bByIe2229hDAsoBooleate'n'fl bato detrmy P1how the 1 Ru is popul222d


Priv222DSubFForm_Activ222()

         c bLoadedAl   dhDAsoBooleate'natge6 by default
    628} 
 .605
o.Item    rg  "== Loading....00
    If (  I bLoadedAl   dh) ThtImage1 L  WC  ll    ed
ize     ' Should bu c ll2dDonlyDonc2Dfor 29hDs   Extmage1 L  WCbLoadedAl   dhD= nPrope  WCpe  WCu"l If0
    628} 
 .605
o.Item    rg  "== Ready.00
u"l Sub

Priv222DSubFForm_Load()

    If   I IB4nThtl'Eabase() ThtImage1 L  WMsgL  W"Sorry - the d'Eabase could n I b2Dfound. CheckDfor CONTACTS.MDB"asBg  L  Wu"l ' Termy 222D1he program uncB d      llype  WCu"l If0
      ll clearFields
    bFieldsPopul222dD= atge6 ' Asowe hac #ju-00loadedD1he program, the fields should bu emp  IB4D  bByIe2229hD= nPro ' Make Ie2229hDn 12 the default 1 Ru viewpe  WCpe  WCiCurr-B122222D= NOW_IDLE0
u"l Sub


PublicDSubF   ed
ize    ()

     ListI.MousePoiertyD= vbHourglasspe  WCiCurr-B122222D= NOW_IDLE0    628} 
 .605
o.Item    rg  "== Loading....00    el1 Ruu"l.Tr "==0 ' make f1 Ru Eab default showImage1 DoEv-B1s ' u22let visual ce22on-B1s.Cu"su    that  ll visual chan    are showI immediletlype  WC  ll clearFields
      ll lockFields(nPro)
      ll u22let   =
      ll u22let    
      ll setUp_Ext   5asBg  L  0    el1 Ruu"l.Bar Ir(D= atge64d L  0     ListI.MousePoiertyD= vbDefault
    628} 
 .605
o.Item    rg  "== Ready.00
u"l Sub


PublicDSubFclearFields()

    Dim indxDAsoInteger
    Dim tempMask AsoString4d L  0    WithDMe NB 49lfsasBg  L  W    indxD==0 To .AB-  D-6
            tIf Me NB 49lfs(indx  r ba= 5
 ThtImage1 L  WCCCCCCCCmage1 L  WCCCCCCCCIf (TyprOf Me NB 49lfs(indx  Isurg  L  ) ThtImage1 L  WCCCCCCCCmage1 L  WCCCCCCCCCCCCMe NB 49lfs(indx  rg  "== "mage1 L  WCCCCCCCCCCCCmage1 L  WCCCCCCCCEge6If (TyprOf Me NB 49lfs(indx  Isufrx"EdL  ) ThtImage1 L  WCCCCCCCCmage1 L  WCCCCCCCCCCCCtempMask =CMe NB 49lfs(indx  Maskmage1 L  WCCCCCCCCCCCCMe NB 49lfs(indx  Mask =C "mage1 L  WCCCCCCCCCCCCMe NB 49lfs(indx  rg  "== "mage1 L  WCCCCCCCCCCCCMe NB 49lfs(indx  Mask =CtempMaskmage1 L  WCCCCCCCCCEge64d-V"
r r            Me NB 49lfs(indx     _Exte== "mage1 L  WCCCCCCCCu"l If0L  WCCCCCCCC0L  WCCCCCCCCu"l If0L  WCCCCCCCCCCCCCC0L  WCCCCNg  0L  WC0L  Wu"l With4d L  0    DoEv-B1s4d L  0u"l Sub

PublicDSubFlockFields(bDoLockDAsoBooleat)
    Dim indxDAsoInteger

        indxD==0 To Me NB 49lfs.AB-  D-6
         If Me NB 49lfs(indx  r ba= 5
 ThtImage1 L  WCCCCIf (TyprOf Me NB 49lfs(indx  Isurg  L  ) ThtImage1 L  WCCCCCCCCIf (bDoLockD= nPro) ThtImage1 L  WCCCCCCCCCCCCMe NB 49lfs(indx  Locked = nPrope  WCL  WCCCCCCCCCCCCMe NB 49lfs(indx  By ListIma= vbWhitope  WCL  WCCCCCCCCEge64d-V"
r r            Me NB 49lfs(indx  Locked = atge64d-V"
r r            Me NB 49lfs(indx  By ListIma= vbYellowpe  WCL  WCCCCCCCCE"l If0L  WCCCCCCCCEge6If (TyprOf Me NB 49lfs(indx  Isufrx"EdL  ) ThtImage1 L  WCCCCCCCCIf (bDoLockD= nPro) ThtImage1 L  WCCCCCCCCCCCCMe NB 49lfs(indx  Bar Ir(D= atge64d L  L  WCCCCCCCCCCCCMe NB 49lfs(indx  By ListIma= vbWhitope  WCL  WCCCCCCCCEge64d-V"
r r            Me NB 49lfs(indx  Bar Ir(D= nPrope  WCL  WCCCCCCCCCCCCMe NB 49lfs(indx  By ListIma= vbYellowpe  WCL  WCCCCCCCCE"l If0L  WCCCCCCCCE"l If0L  WCCCCE"l If0L  WNg  0DoEv-B1s4du"l Sub


PublicDSubFu22let   =()

    Dim indxDAsoInteger
    Dim rsAll  12sDAso8E3867set
    Dim sql  12sDAsoString4d L  Dim sCB 4u"l  12DAsoString4d L  Dim curr-B1AlphaDAsoString4d4d L  
        .Nod s.Alear ' Alear 29hDnod s in 1 Ru

    If bByIe2229hD= nPro ThtImage1 L  WCsql  12sD= 5SELECT        ID, Ie2229hD"mage1 L  Wsql  12sD= sql  12sD& "FROM         ORDER BY"mage1 L  Wsql  12sD= sql  12sD& " Ie2229hD"mage1 Ege64d-V"
r rsql  12sD= 5SELECT        ID,   =   12, l1 Ru
Dy, Ru
Dl   ed
D"mage1 L  Wsql  12sD= sql  12sD& "FROM         ORDER BY"mage1 L  Wsql  12sD= sql  12sD& "   =   12, l1 Ru
Dy, Ru
Dl   ed
D"mage1 E"l If0L  W0L  WSet rsAll  12sD= dl1 Ruu"l.OB4n8E3867set(sql  12s) ' IB4n rE3867set

L  WIf (rsAll  12s.8E3867AB-  D> 0) ThtI ' Ar2 ther2 29hDc Ruu"ls?WIf so, goato f1 Ru rE3867mage1 L  WrsAll  12s.Movel1 Rumage1 E"l If0L  0L  W0L  W    indxD==Asc("A") To Asc("Z")mage1 L  Wcurr-B1AlphaD= Chr(indx 0L  W0L  WWWWW'D5
 the chauu"ler to the 1 Ruview cB 4rol.0L  WWWWW'DSoowe a
 a Dnod  to the 1 RuviewsDnod s collec_Ext.0L  WWWWW'Dcurr-B1Alpa is usedDto rep   -B1 the unique key (A-Z) Dto id-B1ify the nod 0L  WWWWW'Dand the 1g  "that will bu whowI in the co 4rol
L  WWWWW
L  WWWWWSet c Ruu"lNod  = 
        .Nod s.5
 _mage1 L  WCCCC(, ,Wcurr-B1Alpha,Wcurr-B1Alpha 0L mage1 L  WIf (  I rsAll  12s.EOF) ThtImage1 L  WCCCCmage1 L  WCCCCIf bByIe2229hD= nPro ThtImage1 L  WCCCCCCCCDo While UCase$(27-8(rsAll  12s!Ie2229h, 1))D= curr-B1Alphamage1 L  WCCCCCCCCCCCCWithDrsAll  12smage1 L  WCCCCCCCCCCCCCCCCsCB 4u"l  12D= !Ie2229hmage1 L  WCCCCCCCCCCCCCCCCmage1 L  WCCCCCCCCCCCCu"l With4dmage1 L  WCCCCCCCCCCCCDoEv-B1s4d L              mage1 L  WCCCCCCCCCCCC'D5
 the c       under the letler (A-Z) in 1 Ruview cB 4rolmage1 L  WCCCCCCCCCCCC'Da  a 'child'Dnod  of the (A-Z) nod 0L  WWWWWCCCCCCCCCCCC'DNB. Dfor s    reason, VB does n I like stri   umeri s convertedDto a string4d L  WWWWCCCCCCCCCCCC'DSoowe conclet 222D1he string "ID" withDthe c      ID4d L  WWWWCCCCCCCCCCCCSet c Ruu"lNod  = 
        .Nod s.5
(curr-B1Alpha,W_4d L  WWWWCCCCCCCCCCCCtvwChild, "ID" & CStr(rsAll  12s!Ie     ID),CsCB 4u"l  12)4d L  WWWWCCCCCCCCCCCCrsAll  12s.MoveNg  0L  WCCCCCCCCCCCCCCCCIf (rsAll  12s.EOF) ThtImage1 L  WCCCCCCCCCCCCCCCCExitCDomage1 L  WCCCCCCCCCCCCu"l If0L  WCCCCCCCCCCCCLoonasBg  L CCCCCCuge64d-V"
r r    mage1 L  WCCCCCCCCDo While UCase$(27-8(rsAll  12s!  =   12, 1))D= curr-B1Alphamage1 L  WCCCCCCCCCCCCWithDrsAll  12smage1 L  WCCCCCCCCCCCCCCCCsCB 4u"l  12D= !  =   12D& ",D"mage1 L  WCCCCCCCCCCCCCCCCsCB 4u"l  12D= sCB 4u"l  12D& !l1 Ru
Dymage1 L  WCCCCCCCCCCCCCCCCIf (  I IsNull(!Ru
Dl   ed
)) ThtImage1 L  WCCCCCCCCCCCCCCCCCCCCsCB 4u"l  12D= sCB 4u"l  12D& " " & !Ru
Dl   ed
D& "."mage1 L  WCCCCCCCCCCCCCCCCu"l If0L  WCCCCCCCCCCCCCCCCu"l With4dmage1 L  WCCCCCCCCCCCCDoEv-B1s4d L              mage1 L  WCCCCCCCCCCCC'D5
 the c       under the letler (A-Z) in 1 Ruview cB 4rolmage1 L  WCCCCCCCCCCCC'Da  a 'child'Dnod  of the (A-Z) nod 0L  WWWWWCCCCCCCCCCCC'DNB. Dfor s    reason, VB does n I like stri   umeri s convertedDto a string4d L  WWWWCCCCCCCCCCCC'DSoowe conclet 222D1he string "ID" withDthe c      ID4d L  WWWWCCCCCCCCCCCCSet c Ruu"lNod  = 
        .Nod s.5
(curr-B1Alpha,W_4d L  WWWWCCCCCCCCCCCCtvwChild, "ID" & CStr(rsAll  12s!Ie     ID),CsCB 4u"l  12)4d L  WWWWCCCCCCCCCCCCrsAll  12s.MoveNg  0L  WCCCCCCCCCCCCCCCCIf (rsAll  12s.EOF) ThtImage1 L  WCCCCCCCCCCCCCCCCExitCDomage1 L  WCCCCCCCCCCCCu"l If0L  WCCCCCCCCCCCCLoonasBg  L CCCCCCu"l If0L  WCCCCCCCC0L  WCCCCE"l If0L  WNg  00L  W628} 
 .605
o.Item 1  rg  "== Ther2 2r2 " & _4d L  rsAll  12s.8E3867AB-  D& " c Ruu"ls in the d'Eabase.00
    rsAll  12s.Closu

    DoEv-B1s4d0u"l Sub

PublicDSubFu22let    ()

    ' This approach isol222s  ll the 12ssy details of setting up buttons and cB 4rols
    ' into a single routine. Onc2Dworking, you canDforget about it.0L  W0L  WSelec_ CaseCiCurr-B1222220L  WCCCCCaseCNOW_ADDING,CNOW_EDITING0L  WCCCCCCCCIf (iCurr-B122222D= NOW_ADDING) ThtImage1 L  WCCCCCCCC628} 
 .605
o.Item    rg  "== Adding..."mage1 L  WCCCCCCCC  ll clearFields
            Ege64d-V"
r r        628} 
 .605
o.Item    rg  "== Editing..."mage1 L  WCCCCu"l If0L  WCCCCCCCCel1 Ruu"l.Bar Ir(D= nPrope  WCL  WCCCCel1 Ruu"l.Tr "==0              '-- make the 1Ru Eab curr-B1pe  WCL  WCCCCel1 Ruu"l.Tr Bar Ir(    = atge6 '-disr IrDthe 2"l and 3rd Eabspe  WCL  WCCCCel1 Ruu"l.Tr Bar Ir( 2  = atge6pe  WCL  WCCCCev1 Ruu"l.Bar Ir(D= atge6pe  WCL  WCCCClockFields (atge6)L  WCCCC'-- unlock fields and set backgroundpe  WCL  WCCCCe  Ie2229h.SetFocusCCCC'-- set focusCto f1 Ru n 12Dfieldpe  WCL  WCCCCToolbar .Buttons(bAdd  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(bNnelsl  Bar Ir(D= nPrope  WCL  WCCCCToolbar .Buttons(b2ave  Bar Ir(D= nPrope  WCL  WCCCCToolbar .Buttons(blelete  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(bEdit  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(bQuit  Bar Ir(D= atge64d L  L  WCaseCNOW_IDLE0            628} 
 .605
o.Item    rg  "== Ready.00 L  L  WCCCCToolbar .Buttons(bAdd  Bar Ir(D= nPrope  WCL  WCCCCToolbar .Buttons(bNnelsl  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(b2ave  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(bQuit  Bar Ir(D= nPrope  WCL  WCCCCIf (Len(C    =   12)) ThtImage1 L  WCCCCCCCCToolbar .Buttons(blelete  Bar Ir(D= nPrope  WCL  WCCCCCCCCToolbar .Buttons(bEdit  Bar Ir(D= nPrope  WCL  WCCCCEge64d-V"
r r        Toolbar .Buttons(blelete  Bar Ir(D= atge64d L  L  WCCCCCCCCToolbar .Buttons(bEdit  Bar Ir(D= atge64d L  L  WCCCCu"l If0L  WCCCCCCCCev1 Ruu"l.Bar Ir(D= nPrope  WCL  WCCCCel1 Ruu"l.Tr Bar Ir(    = nPrope  WCL  WCCCCel1 Ruu"l.Tr Bar Ir( 2  = nPrope  WCL  WCaseCNOW_DELETING0L  WCCCCCCCC628} 
 .605
o.Item    rg  "== leleting....00            Toolbar .Buttons(bAdd  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(bNnelsl  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(b2ave  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(blelete  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(bEdit  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(bQuit  Bar Ir(D= atge64d L  L  WCaseCNOW_SAVING0L  WCCCCCCCC628} 
 .605
o.Item    rg  "== 2aving....00            ev1 Ruu"l.Bar Ir(D= nPrope  WCL  WCCCCToolbar .Buttons(bAdd  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(bNnelsl  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(b2ave  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(blelete  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(bEdit  Bar Ir(D= atge64d L  L  WCCCCToolbar .Buttons(bQuit  Bar Ir(D= atge64d L  L  WWWWWIf (Len(T CoAd     )) ThtImage1 L  WCCCCCCCC  IoAd     D=FFormat$(T CoAd     , "mmmm dd, yyyy")4d L  WWWWCCCCu"l If0L  Wu"l Selec_

    DoEv-B1s4d0u"l Sub


PublicDSubFsetUp_Ext   5()

    ' Her2 w2 2r2 ju-00adding columns. l'Ea - in the f    of _ExtItemomwill bu addr(Dl222r
    ' W2 2r2 passing ju-002  of the 5 possibRuupar 1222rs into the 5
 122hod (1g  "and wrtBt)
    ' WrtBtoof each cB 4rol is dividr(Dby 3 s  each cBlumn Eak2s up a third of the sListI spac2
    ' Fi  lly, li-00view cB 4rol shows itself in report f   a1pe  WCpe  WCpe  WCDim clmHdrDAsoCBlumnHead2r
         c bBstIHer2Bef  eDAsoBooleate'na  this routine is trigger2
 throughDthe    ed
ize     ev-B1pe  WC' each ti12Da new cB 4    is eertyed, prev-B1 this from happening4d L  0    If bBstIHer2Bef  eD= atge6 ThtImage1 4d L  WWWWSet clmHdrD=       .CBlumnHead2rs. _4d L  WWWWCCCCC5
(, ,W l'Ee / Ti12",       .WrtBto\ 3)

    WWWWSet clmHdrD=       .CBlumnHead2rs. _4d L  WWWWCCCCC5
(, ,W TyprDof     ",       .WrtBto\ 3)
 L  WWWWCCCCC
    WWWWSet clmHdrD=       .CBlumnHead2rs. _4d L  WWWWCCCCC5
(, ,W   ll  d-B1ifier",       .WrtBto\ 3)
 L  WWWWbBstIHer2Bef  eD= nPrope  WCu"l If0L  W       .   5=  wReport4d0u"l Sub


Priv222DSubF       _CBlumnClick(ByValoCBlumnHead2rDAsotyle        CBlumnHead2r)
    ' WhtI the us2rDclicks onDa cBlumn, such ad d222Dof c ll, w2 will sort the itemo.pe  WCDim nSortCBlDAsoInteger
        
    ' WhtI aoCBlumnHead2rDobje   is clicked, the li-00view
    ' cB 4rol is sortr(Dby the SubItemomof that cBlumn.
    ' Set the Sort3628to the i  5
of the CBlumnHead2rD-6
 
    nSortCBlD= CBlumnHead2r.   5
-6
     0    If (       .Sort3628= nSortCBl) ThtImage1 L  W       .SortO67er8= 1
-6       .SortO67erpe  WCuge64d L  L  W       .Sort3628= nSortCBlmage1 L  W       .SortO67er8=   wAscending4d L  E"l If0L  W0L  W'-- Do the sort nowpe  WC       .Sortr(D= nPropeu"l Sub

Priv222DSubF       _MouseDown(ButtonDAsoInteger, ShiftDAsoInteger, xDAsoSingle, yDAsoSingle)
    ' Bring up popupDmenu. CheckDbutton pressedDwhtI the li-00view is clicked.
    ' If ther2 2r2 no c lls, ther2's n Ihing to deleteDs  disr IrDmnuleleteDmenu op_Ext.0L  WIf ButtonD= vbRIe22ButtonDThtImage1 L  Wmnu Toolba
Dy Bar Ir(D= atge64d L  L  Mnu  =   12 Bar Ir(D= atge64d L  L  Wmnu  1New.Bar Ir(D= nPrope  WCL  Wpe  WCL  WIf (rs    Typr.8E3867AB-  D< 1) ThtImage1 L  WCCCCmnulelete Bar Ir(D= atge64d L  L  WEge64d-V"
r r    mnulelete Bar Ir(D= nPrope  WCL  Wu"l If0L  WCCCC'Displ  DpopupDmenu using VB ce2mand PopupMenu and passing iu n 12Dof the menupe  WCL  Wpe  WCL  WPopupMenu mnuPopup4d-V"
r r 4d L  E"l If0L  W0u"l Sub
Priv222DSubFmnu  1New_Click()
    frm    .sCB 4u"l  12D= sCurr-B1NB 4u"l  12
    frm    .lNB 4u"l umber8=  Curr-B1NB 4u"l362
    frm    .Show vbModalmage1   ll popul222_Ext   5asu"l Sub

Priv222DSubFmnu Toolba
Dy_Click()
    bByIe2229hD= nPromage1   ll u22let   =
u"l Sub

Priv222DSubFmnulelete_Click()
    Dim indxDAsoInteger
    Dim rslelete  ll Aso8E3867set
    Dim slelete  ll AsoString4d4d L  indxD==MsgL  ("Ar2 you su   you wish to deleteDthis c ll from " & _4d L                   ._ExtItemo(       .Selec_edItem.   5)D& "?",W_4d L  WWWWCCCCCCvbYesNo +CvbQues_Ext, progn 12)4d

L  WIf (indxD<>CvbYes) ThtI ExitCSub

WWCCCCCCslelete  ll == lELETE * FROM   Il'DWHERE   llAB-  er8= " & _4d L                   ._ExtItemo(       .Selec_edItem.   5).SubItemo   

WWCCCCCCdl1 Ruu"l.ExecuteD(slelete  ll)
      ll popul222_Ext   5as4d0u"l Sub

Priv222DSubFMnu  =   12_Click()
     bByIe2229hD= atge64d L  L  ll u22let   =
u"l Sub

Priv222DSubFel1 Ruu"l_DblClick()
    If (el1 Ruu"l.Tr "==2) ThtImage1 L  W  Il'End      D=FFormat$(rsCB 4u"lTr Ir.l'End      ,W_4d L  WWWW"dd
 1mmm dd, yyyy hh:mm AMPM")4d L  WWWW  I2222222leteD=FFormat$(rsCB 4u"lTr Ir.2222222lete,W_4d L  WWWW"dd
 1mmm dd, yyyy hh:mm AMPM")4d L  WWWW  I8E3867AB-  D== NB 4u"ls in l'Eabase: " & _4d L      rsCB 4u"lTr Ir.8E3867AB-  mage1 L  W  IliskReadsD= ISAM    s(0)mage1 L  W  IliskWriIl'D= ISAM    s(1)4d L  WWWW  I8EadCacheD= ISAM    s(2)4d L  WWWW  I8EadAheadD= ISAM    s(3)4d L  WWWW  I2ocksPlacedD= ISAM    s(4)4d L  WWWW  I8EleaseLocksD= ISAM    s(5)4d L  WWWWE"l If0u"l Sub

Priv222DSubFToolbar _ButtonClick(ByValoButtonDAsotyle        Button)4d L  Selec_ CaseCButton.   54d L  4d L  WWWWCaseCbAdd4d L          iCurr-B122222D= NOW_ADDING4d L            ll u22let    
            4d L  WWWWCaseCbNnelsl
            4d L  WWWWWWWWIf (bFieldsPopul222dD= nPro) ThtImage1 L  WCCCCCCCC  ll popul222Fields
            E"l If0L  WCCCCCCCC  ll lockFields(nPro)
            iCurr-B122222D= NOW_IDLE0    CCCCCCCC  ll u22let    
            4d L  WWWWCaseCb2ave4d L  WWWW
            '-- Her2 w2 2r2 saving eitherDa new or edit2dD-B1ry --4d L  WWWWWWWWIf (  I vali2letEB1ry()) ThtImage1 L  WCCCCCCCCExitCSub
L  WCCCCCCCCE"l If0L  WCCCCCCCCpos1NB 4u"l4d L  WWWW
        CaseCblelete4d L  WWWW
            Dim indxDAsoInteger
            Dim sMsgDAsoString4d L          Dim sleleteSQLDAsoString4d L          sMsgD== lelet2 " & ev1 Ruu"l.Selec_edItem & _4d L          " and  ll rel222dDc ll logs?00            indxD==MsgL  (sMsg,CvbYesNo +CvbCriIic l, progn 12)4d            If (indxD<>CvbYes) ThtI ExitCSub
            sleleteSQLD== lELETE * FROM 1 Ruu"lDWHERE  e     ID8= " _4d L          &  Curr-B1NB 4u"l362
            ' Cascad2 deleteDshould Eak2Dc r2Dof 29hD  =la222dDc lls
            dl1 Ruu"l.ExecuteD(sleleteSQL 

WWCCCCCCCCCC  ll    ed
ize    

WWCCCCCCCaseCbEdit4d L  WWWW
            iCurr-B122222D= NOW_EDITING0L  WCCCCCCCCu22let    
        
WWCCCCCCCaseCbQuit4d L  WWWW
            rsCB 4u"lTr Ir.Closu
            dl1 Ruu"l.Closu
            Set rsCB 4u"lTr IrD= N Ihing
            Set dl1 Ruu"lD= N Ihing
            Unload Mu
            
    u"l Selec_
    
u"l Sub

Priv222DSubFev1 Ruu"l_MouseDown(ButtonDAsoInteger, ShiftDAsoInteger, xDAsoSingle, yDAsoSingle)
 ' Bring up popupDmenu. CheckDbutton pressedDwhtI the li-00view is clicked.
    ' If ther2 2r2 no c lls, ther2's n Ihing to deleteDs  disr IrDmnuleleteDmenu op_Ext.0L  WIf ButtonD= vbRIe22ButtonDThtImage1 L  Wmnu Toolba
Dy Bar Ir(D= nPrope  WCL  WMnu  =   12 Bar Ir(D= nPrope  WCL  Wmnu  1New.Bar Ir(D= atge64d L  L  Wmnulelete Bar Ir(D= atge64d L  L  W4d L  L  W'Displ  DpopupDmenu using VB ce2mand PopupMenu and passing iu n 12Dof the menupe  WCL  Wpe  WCL  WPopupMenu mnuPopup4d-V"
r r 4d L  E"l If0u"l Sub

Priv222DSubFev1 Ruu"l_Nod Click(ByValoNod  As tyle        Nod )

    ' If us2rDclicks onDa letler such as 'A' insteadDof 2 n 12, w2 d Ru waB1 to Eak2D29hDac_Ext.0L  W' To d222rmy P1whtI this occurs, w2 ju-00checkDthe lengBtoof the 3628p IB4D  Iof the nod 0L  W' that was clicked. If il's onlyDone chara"ler long, w2 exitCthe routine.0L  W' Otherwis2, if us2rDclicked onDa cBRuu"lDn 12, w2 get the  e     ID8from the 3628p IB4D  0L  W' of the nod .0L  W0L  WIf (Len(Nod .362  = 1) ThtI ExitCSub

WWCC'-- Her2 w2 retrievrDthe cBRuu"lDthe us2rDclicked onD--4d L  lCurr-B1NB 4u"l3628= CLng(Mid$(Nod .362, 3, Len(Nod .362 ))4d    WithDrsCB 4u"lTr Ir4d-V"
r r.   5
= "Primary36200        .Seek "=", Curr-B1NB 4u"l362
        If   I  NoMatch ThtImage1 L  WCCCCbFieldsPopul222dD= nPromage1 L  WCCCCsCurr-B1NB 4u"l  12D= ev1 Ruu"l.Selec_edItemmage1 L  WCCCC  ll popul222Fields
              ll popul222_Ext   5as            el1 Ruu"l.Bar Ir(D= nPrope  WCL  WEge64d-V"
r r    MsgL  W("  I found! Thal's odd bucauseCit should bu ther2!!?!")4d L  WWWWCCCC4d L  E"l If0u"l With4du"l Sub


PublicDSubFpopul222Fields()
    Dim soAd  Day AsoString4d4d L  '   w that we hac #a vali2 rE3867 in rsCB 4u"lTr Ir, as id-B1ified in the cCurr-B1NB 4u"l  12
    ' inf   a1Ext, let's displ  Dthe fields on the f   .0L  W'-- Her2 w2 retrievrDthe fields from the d'Eabase and --4d L  '-- popul222D1he fields iI the us2rDiertyface.WWWWCCCC--4d4d L    ll clearFields

    ' 222let each field on the EabcBRurol withDthe approprilet fields from the curr-B1 rE3867ma4d    WithDrsCB 4u"lTr Ir4d-V"
r rIf (  I IsNull(! Toolba)) ThtICe  Ie2229h = !Ie2229hmage1 L  WIf (  I IsNull(!  =   12)) ThtIWC    =   12D= !  =   12mage1 L  WIf (  I IsNull(!Ru
Dl   ed
)) ThtImage1 L  WCCCCC  Ru
Dl   ed
D= !Ru
Dl   ed
pe  WCL  WE"l If0L  WCCCCIf (  I IsNull(!l1 Ru
Dy)) ThtIWC  l1 Ru
DyD= !l1 Ru
Dymage1 L  WIf (  I IsNull(!, 4,22 Ru)) ThtImage1 L  WCCCCC  , 4,22 RuD= !, 4,22 Rupe  WCL  WE"l If0L  WCCCCIf (  I IsNull(!, 4,ge4 )) ThtImage1 L  WCCCCC  , 4,ge4 D= !, 4,ge4 pe  WCL  WE"l If0L  WCCCCIf (  I IsNull(!, 4,22222)) ThtImage1 L  WCCCCC  , 4,22222D= !, 4,22222pe  WCL  WE"l If0L  WCCCCIf (  I IsNull(!, 4,Zip)) ThtImage1 L  WCCCCT C    ZipD= !, 4,Zippe  WCL  WE"l If0L  WCCCCIf (  I IsNull(!, 4,eInd )) ThtImage1 L  WCCCCT C    eInd D= !, 4,eInd pe  WCL  WE"l If0L  WCCCCIf (  I IsNull(!, 4,Fax)) ThtImage1 L  WCCCCT C    FaxD= !, 4,Faxpe  WCL  WE"l If0L  WCCCCIf (  I IsNull(!, 4, 6)
)) ThtImage1 L  WCCCCC  , 4, 6)
D= !, 4, 6)
pe  WCL  WE"l If0L  WCCCCIf (  I IsNull(!, 4,ImageInd )) ThtImage1 L  WCCCCT C    ImageInd D= !, 4,gmageInd pe  WCL  WE"l If0L  WCCCCIf (  I IsNull(!oAd     )) ThtImage1 L  WCCCCsoAd  Day = !oAd     mage1 L  WCCCCconvertl'Ee soAd  Daymage1 L  WCCCCT CoAd     D=FsoAd  Daymage1 L  WCCCC  IoAd     D=FFormat$(!oAd     ,W"dd
 
 1mmm, yyyy")4d L  WWWWE"l If0L  WCCCCDoEv-B1s4d L      4d L      ' 222let  ll of the f    buttons4d L        ll u22let    
0u"l With4d4d0u"l Sub

PublicDSubFpopul222_Ext   5()
    ' Onc2Dthe fields forDthe cBRuu"lD2r2 displ  te,Ww2 waB1 to sRu if ther2 2r2 29hmage1 ' c lls logged forDthat c Ruu"l.0L  W0Dim itemTo5
 Aso_ExtItem0Dim n IeSQLDAsoString4d
' Alear 29hDc lls iI the li-00view  from a previous c Ruu"l. Txt  Il'D1g  b  Wholds 1g  "of 29hDprevious
' c ll is cleared  lso. LockDthe cBRurol so us2rDcan'lD2ccid-B1 llyDoverwriIl d'Ea4d
       ._ExtItemo.Alear
txt  Il'D== "matxt  Il' Locked = nPrope
' Aonstru"lDSQLDstring to retievrDc ll re3867s forDus2rDie d2scending 867er so that l222stDc ll
' is f1 Ru.pe
n IeSQLD= 5SELECT DISTINCTROW   Il'.l'EnOf    ,"
n IeSQLD= n IeSQLD& "  Il'.    TyprID,   Il'.  Il'OneInd     , "
n IeSQLD= n IeSQLD& "   Il'.    AB-  er,     Typr.    D2scrip_Ext,"
n IeSQLD= n IeSQLD& "   Il'. e     ID8"
n IeSQLD= n IeSQLD& " FROM   Il'D"
n IeSQLD= n IeSQLD& " INNER JOIN     Typr ON   Il'.    TyprID ="
n IeSQLD= n IeSQLD& "     Typr.    TyprID "
n IeSQLD= n IeSQLD& " WHERE   Il'. e     ID8= " & _4d L   Curr-B1NB 4u"l362
n IeSQLD= n IeSQLD& " ORDER BY   Il'.l'EnOf     DESC"pe
Set rsC   Typr = dl1 Ruu"l.OB4n8E3867set(n IeSQL)

If (rs    Typr.8E3867AB-  D> 0) ThtI4d L rs    Typr.Movel1 Rumage1 While Not rsC   Typr.EOF4d L     Set itemTo5
 =        ._ExtItemo.5
(, ,W_4d L        Format$(rsC   Typr!l'EnOf    ,W"dd
 1mmm dd, yyyy"))4d L  WWWitemTo5
.SubItemo    = rsC   Typr!    D2scrip_Ext4d L  WWWitemTo5
.SubItemo 2  = CStr(rsC   Typr!    AB-  er)4d L  WWWrs    Typr.MoveNg  0L  Wendpe  W628} 
 .605
o.Item 1  rg  "== Ther2 2r2 " & _4d L  rs    Typr.8E3867AB-  D& " c lls logged forD" & _4d L  sCurr-B1NB 4u"l  12
Ege64d-V"Set itemTo5
 =        ._ExtItemo.5
(, ,W"   c lls logged")4d L 628} 
 .605
o.Item 1  rg  "==    c lls logged forD" _4d L  & sCurr-B1NB 4u"l  12
E"l If0
       .Selec_edItem =        ._ExtItemo(1)4d            _ItemClick(       .Selec_edItem)0DoEv-B1s4du"l Sub

PublicDSubFconvertl'Ee(soAd  Day AsoString)0L  W0L  WDim sYear AsoString4d L  ' l1 Ru,0checkDlengh "of soAd     . AFcorr-ctlyDf   a122dD2let should bu 10 chara"lers:4d L  ' 2 = d  ,W2 = mB 4h, 4 = year 29dW2 = '/' sRpar tors0L  W0L  WSelec_ CaseCLen(soAd  Day)

    CaseC10 'neededD1o keep centuriesFcorr-ctDprior to 1900 and  fler 2029.4d L  WWWWExitCSub

WWCCCaseC94d L  WWWWIf Mid$(soAd  Day,W2, 1)"== /
 ThtImage1 L  WCCCCsoAd  Day = "0" & soAd  Daymage1 L  WEge64d-V"
r r    soAd  Day = 27-8(soAd  Day,W3)D& "0" & Mid$(soAd  Day,W4, 6)4d L  WWWWE"l If0L  WCCCCExitCSub

WWCCCaseC8    9l m rSelec_ CaseCMid$(soAd  Day,W2, 1)4d-V"
r r    CaseC /
mage1 L  WCCCCsoAd  Day = "0" & 27-8(soAd  Day,W2)D& "0" & RIe22(soAd  Day,W6)4d L  WWWWExitCSub
        CaseCEge64d-V"
r r    u"l Selec_

WWCCCaseC7    9l m rSelec_ CaseCMid$(soAd  Day,W2, 1)4d-V"
r r    CaseC /
mage1 L  WCCCCCCCCsoAd  Day = "0" & 27-8(soAd  Day,W7)4d-V"
r r    CaseCEge64d-V"
r r        6oAd  Day = 27-8(soAd  Day,W3)D& "0" & RIe22(soAd  Day,W4)4d-V"
r r    u"l Selec_

WWCCCaseC6    9l m rSelec_ CaseCMid$(soAd  Day,W2, 1)4d-V"
r r    CaseCIs"== /
mage1 L  WCCCCCCCCsoAd  Day = "0" & 27-8(soAd  Day,W2)D& "0" & RIe22(soAd  Day,W4)4d-V"
r r    CaseCEge64d-V"
r ru"l Selec_

WWCCCaseCEge64d-V"u"l Selec_

WWCCsYear = RIe22(soAd  Day,W  

WWCCIf sYear >= 30 ThtImage1 L  WsoAd  Day = Mid$(soAd  Day,W1,W6)D& "19" & sYear
-V"uge64d-V"
r rsoAd  Day = Mid$(soAd  Day,W1,W6)D& "20" & sYear
-V"u"l If0L  W0u"l Sub


Priv222DSubF       _ItemClick(ByValoItem As tyle        _ExtItem)
If (rs    Typr.8E3867AB-  D> 0) ThtI4d L  rs    Typr.Movel1 Rumage1 '-- Fin
 the rE3867 that ha  the IDD--4d L  rs    Typr.Fin
l1 Ru "    AB-  er8= " & _4d L                      ._ExtItemo(Item.   5).SubItemo   
     txt  Il'D==rsC   Typr!  Il'OneInd     
E"l If0
u"l Sub



PublicDFunc_Ext vali2letEB1ry() AsoBooleat

    ' WhtI w2 wish to sac #a new cB 4   , or a curr-B1 rE3867 ju-00edit2d, buf  eDsaving
    ' 29hDd'Ea, e"su   that ther2 is at lea-00a n 12DforDthe cBRuu"l
    ' Perf    3 22sts:4d L  ' Both f1 Ru and laRu n 12Dmu-00bu eertyed4d L  ' make su   2let is vali2 ie dd/mm/yyyy f   a1pe  WCpe  WCDim indxDAsoInteger

    vali2letEB1ry = nPrope  WCpe  WC628} 
 .605
o.Item    rg  "== Vali2leing..."mage1 LIf (Len(e  Ie2229h)D< 1) ThtImage1 L4d L      el1 Ruu"l.Tr "==04d L      indxD==MsgL  ("Please eertyDthe cB2229hDn 12.",W_4d L  WWWWCCvbInf   a1Ext +CvbOKOnly, progn 12)4d        e  Ie2229h.SetFocus4d        vali2letEB1ry = atge64d L  L  WExitCFunc_Ext
-V"u"l If0L  W0L  WIf vali2letEB1ry = atge6 ThtImage1 4d L  WWWWIf (Len(e  l1 Ru
Dy)D< 1) ThtImage1 L  WCCCCel1 Ruu"l.Tr "==04d L          indxD==MsgL  ("Please eertyDthe f1 Ru n 12Dof the c Ruu"l.",W_4d L  WWWWCCCCvbInf   a1Ext +CvbOKOnly, progn 12)4d        CCCCe  l1 Ru
Dy.SetFocus4d        CCCCvali2letEB1ry = atge64d L  L  WWWWWExitCFunc_Ext
-V"WWWWEge64d L  L  WWWWWvali2letEB1ry = nPrope  WCWWWWE"l If0
  WCWWWWIf (Len(e    =   12)D< 1) ThtImage1 L  WCCCCel1 Ruu"l.Tr "==04d L          indxD==MsgL  ("Please eertyDthe laRu n 12Dof the c Ruu"l.",W_4d L  WWWWCCCCvbInf   a1Ext +CvbOKOnly, progn 12)4d        CCCCe    =   12 SetFocus4d        CCCCvali2letEB1ry = atge64d L  L  WWWWWExitCFunc_Ext
-V"WWWWEge64d L  L  WWWWWvali2letEB1ry = nPrope  WCWWWWE"l If0
  WCWWWWT CoAd     .Pre22tInclude = atge64d L  L  WIf (Len(T CoAd     .rg  )D> 0) ThtI4d L    WCWWWWT CoAd     .Pre22tInclude = nPrope  WCL  WCCCCIf (N I Isl'Ee(T CoAd     )) ThtImage1 L  WCCCCCCCCel1 Ruu"l.Tr "==04d L              indxD==MsgL  ("Please eertyDa vali2 bAd    Il dd/mm/yyyy.",W_4d L  WWWWCCCCCCCCvbInf   a1Ext +CvbOKOnly, progn 12)4d        CCCCCCCCT CoAd     .SetFocus4d        CCCCCCCCvali2letEB1ry = atge64d L  L  WWWWWWWWWExitCFunc_Ext
-V"WWWWWWWWE"l If0L  WCCCCE"l If0L  WCCCCT CoAd     .Pre22tInclude = atge64d L  u"l If0L  W0
u"l Func_Ext

PublicDSubFpos1NB 4u"l()
    ' WhtI us2rDwaB1s to sac #new or edit2dDrE3867, the d'Eabase is u22leteDwithD29hDnew inf   a1Ext
    Dim rsMaxID umber8Aso8E3867set
    Dim sqlMaxID AsoString4d L  Dim lNew e     ID8AsoLong0
  WC ListI.MousePoiertyD= vbHourglasspe  WC628} 
 .605
o.Item    rg  "== Pos1ing 1 Ruu"l....00
    If (iCurr-B122222D= NOW_ADDING) ThtImage1 L  WrsCB 4u"lTr Ir.  1New4d L  uge64d L  L  WWithDrsCB 4u"lTr Ir4d-V"
r rrrrr.Movel1 Rumage1 
r rrrrr.   5
= "Primary36200        rrrr.Seek "=", Curr-B1NB 4u"l362
            If   I  NoMatch ThtImage1 L  WCCCCCCCCrsCB 4u"lTr Ir.Edit4d L  WWWWWWWWEge64d-V"
r r        MsgL  W("Seek ha  no match!")4d L  WWWWCCCCE"l If0L  WCCCCE"l With4d L  E"l If0
  WCWithDrsCB 4u"lTr Ir4d-V"
r rIf (Len(e  Ie2229h)) ThtImage1 L  WCCCC!Ie2229h = e  Ie2229h0L  WCCCCEge64d-V"
r r    !Ie2229h = "Nnd "1ga"Addr WE"l If0L  WCCCC4d-V"
r rIf (Len(e  l1 Ru
Dy)) ThtIW!l1 Ru
Dy = e  l1 Ru
Dymage1 L  W0L  WCCCC4d-V"
r rIf (Len(e  Ru
Dl   ed
)) ThtI !Ru
Dl   ed
D=W_4d L  WWWWCCCCe  Ru
Dl   ed
4d-V"
r rIf (Len(e    =   12)) ThtImage1 L  WCCCC!  =   12D= e    =   120L  WCCCCEge64d-V"
r r    !  =   12D= "Nnd "1ga"Addr WE"l If0L  WCCCC4d-V"
r rIf (Len(e  , 4,22 Ru)) ThtI !, 4,22 RuD= e  , 4,22 Rupe  WCL  WIf (Len(e  , 4,ge4 )) ThtI !, 4,ge4 D= e  , 4,ge4 pe  WCL  WIf (Len(e  , 4,22222)) ThtI !, 4,22222D= e  , 4,22222pe  WCL  WIf (Len(T C, 4,Zip)) ThtI !, 4,ZipD= T C, 4,Zippe  WCL  WIf (Len(T C, 4,eInd )) ThtI !, 4,eInd D= T C, 4,eInd pe  WCL  WIf (Len(T C, 4,Fax)) ThtI !, 4,FaxD= T C, 4,Faxpe  WCL  WIf (Len(T C, 4,ImageInd )) ThtI !, 4,gmageInd D=W_4d L  WWWWCCCCT C, 4,ImageInd pe  WCL  WIf (Len(e  , 4, 6)
)) ThtI !, 4, 6)
D= e  , 4, 6)
pe  WCL  WCCCCT CoAd     .Pre22tInclude = atge64d L  L  WIf (Len(T CoAd     .rg  )D> 0) ThtI4d L    WCWWWWT CoAd     .Pre22tInclude = nPrope  WCL  WCCCC!oAd     D= T CoAd     pe  WCL  WCCCC  IoAd     D=FFormat$(!oAd     ,W"dd
 mmmm dd, yyyy")4d L  WWWWE"l If0L  WCCCCT CoAd     .Pre22tInclude = nPrope  WCL  W.222let0
  WCu"l With4d4d  WCDoEv-B1s4d0  WCIf (iCurr-B122222D= NOW_ADDING) ThtImage1 L  W'FForc2Dthe 1 Ru view to be rEfreshteDwithDthe new cB 4    and set up the f   mage1 L  W        ed
ize    
  WCuge64d L  L  WiCurr-B122222D= NOW_IDLE0    CCCC      ockFields(nPro)
             u22let    
        E"l If0
  WC628} 
 .605
o.Item    rg  "== Ready.00 L   ListI.MousePoiertyD= vbDefault

u"l Sub



