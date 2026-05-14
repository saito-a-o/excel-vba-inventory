Attribute VB_Name = "definition"
Option Explicit

'シート名
Public Const SHEET_MITSUMORI As String = "見積書"
Public Const SHEET_NOUHINN As String = "納品書"
Public Const SHEET_SEIKYUU As String = "請求書"
Public Const SHEET_RYOUSYUU As String = "領収書"
'発行元セルアドレス
Public Const SRC_RANGE_MYCOMP As String = "$B$4"
Public Const SRC_RANGE_MYPOST3 As String = "$C$5"
Public Const SRC_RANGE_MYPOST4 As String = "$G$5"
Public Const SRC_RANGE_MYADDR1 As String = "$B$6"
Public Const SRC_RANGE_MYADDR2 As String = "$B$7"
Public Const SRC_RANGE_MYTELNUMB As String = "$B$8"
Public Const SRC_RANGE_MYFAXNUMB As String = "$B$9"
Public Const SRC_RANGE_MYTNUMB As String = "$B$10"
Public Const SRC_RANGE_MYPERSON As String = "$B$11"

'取引先セルアドレス
Public Const SRC_RANGE_OUTCOMP As String = "$B$14"
Public Const SRC_RANGE_OUTPOST3 As String = "$C$15"
Public Const SRC_RANGE_OUTPOST4 As String = "$G$15"
Public Const SRC_RANGE_OUTADDR1 As String = "$B$16"
Public Const SRC_RANGE_OUTADDR2 As String = "$B$17"
Public Const SRC_RANGE_OUTORDRNUM As String = "$B$18"
Public Const SRC_RANGE_OUTPRCDATE As String = "$B$19"

'明細セルアドレス
Public Const SRC_RANGE_DETAILITEM1 As String = "$A$23"
Public Const SRC_RANGE_DETAILNUM1 As String = "$L$23"
Public Const SRC_RANGE_DETAILPRCE1 As String = "$M$23"
Public Const SRC_RANGE_DETAILRATE1 As String = "$N$23"

'転記先取引先
Public Const TGT_RANGE_OUTCOMP As String = "$B$5"
Public Const TGT_RANGE_OUTPOST As String = "$B$6"
Public Const TGT_RANGE_OUTADDR1 As String = "$B$7"
Public Const TGT_RANGE_OUTADDR2 As String = "$B$8"

'転記先発行者
Public Const TGT_RANGE_MYCOMP As String = "$H$5"
Public Const TGT_RANGE_MYPOST As String = "$H$6"
Public Const TGT_RANGE_MYADDR As String = "$H$7"
Public Const TGT_RANGE_MYTNUM As String = "$H$8"
Public Const TGT_RANGE_MYEL As String = "$H$9"
Public Const TGT_RANGE_MYFAX As String = "$H$10"
Public Const TGT_RANGE_MYPSN As String = "$H$11"

'転記先明細
Public Const TGT_RANGE_ITEM As String = "$B$18" '見積書は$B$19
Public Const TGT_RANGE_LGT As String = "$E$18"
Public Const TGT_RANGE_NUM As String = "$F$18"
Public Const TGT_RANGE_NLBL As String = "$G$18"
Public Const TGT_RANGE_PRICE As String = "$H$18"
Public Const TGT_RANGE_RATE As String = "$I$18"

'注文情報転記先
Public Const TGT_RANGE_MIT_SUBJ As String = "$C$10"
Public Const TGT_RANGE_MIT_DLDT As String = "$C$11"
Public Const TGT_RANGE_MIT_TERM As String = "$C$12"
Public Const TGT_RANGE_MIT_EXPD As String = "$C$13"
Public Const TGT_RANGE_NOU_SUBJ As String = "$C$12"
Public Const TGT_RANGE_SEI_SUBJ As String = "$C$10"
Public Const TGT_RANGE_SEI_DUED As String = "$C$11"
Public Const TGT_RANGE_SEI_BANK As String = "$C$12"
Public Const TGT_RANGE_RYO_SUBJ As String = "$C$12"


'発行日と書類番号（4シート共通）
Public Const TGT_RANGE_COMM_DOCD As String = "$J$1"
Public Const TGT_RANGE_COMM_DOCN As String = "$J$2"


Public Type Data_To
    CompTo       As String
    PostnumTo    As String
    Addr1To      As String
    Addr2To      As String
    OrdernumTo   As String
    DateTo       As Date
End Type

Public Type Data_From
    CompFrm      As String
    PostnumFrm   As String
    Addr1Frm     As String
    Addr2Frm     As String
    TelFrm       As String
    FaxFrm       As String
    TnumFrm      As String
    PrsnFrm      As String
End Type
Public Enum CLM_DETAIL
    dtail_item = 1
    dtail_num = 12
    dtail_price = 13
    dtail_rate = 14
End Enum
