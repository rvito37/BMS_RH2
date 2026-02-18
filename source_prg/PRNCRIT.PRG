// prncrit.prg
// Function/Procedure Prototype Table  -  Last Update: 14/06/99 @ 15:13:17
// ��������������������������������������������������������������������������
// Return Value         Function/Arguments
// �������������������  �����������������������������������������������������
// .T.                  STATIC FUNCTION PUTREFTO(alist)
// cTemp                STATIC FUNCTION Turn2Str( Buffer ,lFlag )
// lRetVal              STATIC FUNCTION postBN(o)
// lRetVal              STATIC FUNCTION postProcId(o)
// lRetVal              STATIC FUNCTION postTime(o)
// lRetVal              STATIC FUNCTION postWork(o)
// Void                 STATIC PROCEDURE OpenIndex( cInd , cFile )
// Void                 STATIC PROCEDURE prnShow( cFile )
// Void                 PROCEDURE prnKillCrit
// Void                 PROCEDURE prnSetBuffer( xVal )
// Void                 PROCEDURE prnSetCritBuffer(cFile , lCode )
// cTemp                FUNCTION BuildCritData( aCut , aCutIndicators , lFlag)
// .T.                  FUNCTION DaysOnLine( o, aParamList )
// u1 == u2             FUNCTION Equal( u1, u2 )
// .T.                  FUNCTION GenericNumber(o,aParamList,cTitle,cFieldFrom,cFieldTo)
// .T.                  FUNCTION Get42ProcId( o, aParamList )
// .T.                  FUNCTION Get42WkStn( o, aParamList )
// .T.                  FUNCTION GetAllocDate( o, aParamList )
// .T.                  FUNCTION GetBN( o, aParamList,cFile,cIfile )
// .T.                  FUNCTION GetBatch( o, aParamList )
// .T.                  FUNCTION GetBookedDate( o, aParamList )
// uRet                 FUNCTION GetBuffer( cCode )
// ""                   FUNCTION GetBuffer( cCode )
// .T.                  FUNCTION GetDate( o, aParamList )
// .T.                  FUNCTION GetDialect( o, aParamList )
// .T.                  FUNCTION GetDocNo( o, aParamList,cFile,cIfile )
// .T.                  FUNCTION GetDueDate( o, aParamList )
// .T.                  FUNCTION GetEsn( o, aParamList )
// .T.                  FUNCTION GetFinDateTime(o,aParamList)
// .T.                  FUNCTION GetFinishDate( o, aParamList )
// .T.                  FUNCTION GetGeneric(o,aParamList,cTitle,cFieldFrom,cFieldTo)
// .T.                  FUNCTION GetGuy( o, aParamList )
// .T.                  FUNCTION GetMinSold( o, aParamList )
// .T.                  FUNCTION GetMonths( o, aParamList )
// .T.                  FUNCTION GetMrkDue( o, aParamList )
// .T.                  FUNCTION GetNo2OfGaps( o, aParamList)
// .T.                  FUNCTION GetNo3OfGaps( o, aParamList)
// .T.                  FUNCTION GetNoOfDays( o, aParamList )
// .T.                  FUNCTION GetNoOfGaps( o, aParamList)
// .T.                  FUNCTION GetOneBN( o, aParamList )
// .T.                  FUNCTION GetOrderDate( o, aParamList )
// .T.                  FUNCTION GetPcAndDate( o, aParamList )
// .T.                  FUNCTION GetPcAndPrior( o, aParamList )
// .T.                  FUNCTION GetPcNo( o, aParamList )
// .T.                  FUNCTION GetPcName( o, aParamList )
// .T.                  FUNCTION GetPolnid( o, aParamList )
// .T.                  FUNCTION GetProcId( o )
// .T.                  FUNCTION GetRefNo( o, aParamList )
// .T.                  FUNCTION GetRemark( o, aParamList )
// .T.                  FUNCTION GetReqDate( o, aParamList )
// .T.                  FUNCTION GetRevDate( o, aParamList )
// .T.                  FUNCTION GetRoute( o, aParamList )
// .T.                  FUNCTION GetSerNo( o, aParamList )
// .T.                  FUNCTION GetShippedDate( o, aParamList )
// .T.                  FUNCTION GetSlack( o, aParamList )
// .T.                  FUNCTION GetHours( o, aParamList )
// .T.                  FUNCTION GetStartDate( o, aParamList )
// .T.                  FUNCTION GetTime( o, aParamList )
// .T.                  FUNCTION GetToDate( o, aParamList )
// .T.                  FUNCTION GetWarehouse( o, aParamList)
// .T.                  FUNCTION GetWkStn( o )
// u1 >= u2             FUNCTION GrEq( u1, u2 )
// u1 >= u2 .AND. u...  FUNCTION GrLeEq( u1, u2, u3 )
// u1 > u2              FUNCTION Great( u1, u2 )
// cSearch $ cSource    FUNCTION InStr( cSearch, cSource )
// lRetVal              FUNCTION IsThereCond()
// TRUE                 FUNCTION ItsOk
// u1 < u2              FUNCTION Less( u1, u2 )
// u1 <= u2             FUNCTION LessEq( u1, u2 )
// .T.                  FUNCTION QtyGreaterThan( o, aParamList ,cTitle)
// lRetVal              FUNCTION ShaiCond( aTestBlocks,oScrl,cField )
// aTemp                FUNCTION aBuildCritData( aCut , aCutIndicators , lFlag)
// .T.                  FUNCTION closeFromTo( o, aParamList )
// .T.                  FUNCTION critBrowse( o, aParamList, cFile, cIndex, bKeyCol, bNameCol, bFilter)
// .T.                  FUNCTION getpccode(o, aParamList)
// .T.                  FUNCTION getpccode2(o, aParamList)
// lRetVal              Static Function PostFindate(o,ddFrom)
// cSearcher            function Searchit(cSearcher,cFile)

// G:\BMS\SOURCE\PRNCRIT.PRG
// s.b. How to add new crit ?
// 1. Add Static Var if it's single value or 2 var's if it's From-To.
// 2. Add it to the PrnKillCrit
// 3. If it's From-To Type write a Function like getTime,GetRefNo etc.
// 4. Add to the Case in PrnSetCritBuffer function.
// 5. Add  In to GetBuffer Function in the aSource & aTitle Arrays
// 6. add it in the report /query in SetCrit , SetCheck ,SetBuffer
//    SetQueryBlocks Methods
#include "avxdefs.ch"

STATIC  ;
aRecNos     ,;
cShipAs     ,;//VR 13-11-01 111132334
cAgency     ,;
cArea       ,;
cBNFrom     ,;
cBNTo       ,;
cBatchType  ,;
cBid        ,;
cTop        ,;//VR
cBloke      ,;
cBstat      ,;
cPBstat     ,;
cNBstat     ,;
cProj       ,;
cBuffer     ,;
cCompany    ,;
cCountry    ,;
cCustomer   ,;
cDOCFrom    ,;
cDialecter  ,;
cDocTo      ,;
cEngineer   ,;
cESNXX      ,;
cESNY       ,;
cEsnFrom    ,;
cEsnTo      ,;
cMoveType   ,;
cPOType     ,;
cPcCodeFrom ,;
cPcCodeTo   ,;
cPoStat     ,;
cPriorFrom  ,;
cPriorTo    ,;
cPriority   ,;
cProcFrom   ,;
cProcIDFrom ,;
cProcIDTo   ,;
cProcId     ,;
cProcTo     ,;
cProcType   ,;
cBncaps     ,;
cProductLine,;
cProductType,;
cPurpose    ,;
cRefFrom    ,;
cRefTo      ,;
cRejectId   ,;
cRemark     ,;
cRoute      ,;
cSValFile   ,;
cSValmFile  ,;
cSVMoveType ,;
cSize       ,;
cTC         ,;
cOma        ,;
cRepUnit    ,;
cEndCustomer,;
cTermination,;
cTol        ,;
cValue      ,;
cYesNo      ,;
cVoltage    ,;
cWarehouse  ,;          //YG 29/09/97 for rfg07v1
cWareh_sr   ,;
cWareh_ds   ,;
cWorkStation,;
cWorkerId   ,;
cWrkStnFrom ,;
cWrkStnTo   ,;
dAllocFrom  ,;
dAllocTo    ,;
dBookedFrom ,;
dBookedTo   ,;
dNonNon     ,;
dDueFrom    ,;
dDueTo      ,;
dEFrom      ,;
dETo        ,;
dFinFrom    ,;
dFinTo      ,;
dFrom       ,;
dMrkDueFrom ,;
dMrkDueTo   ,;
dReqFrom    ,;
dReqTo      ,;
dRevFrom    ,;
dRevTo      ,;
dShippedFrom,;
dShippedTo  ,;
dStartFrom  ,;
dStartTo    ,;
dTo         ,;
nDays       ,;
nFrom       ,;
nHoursFrom  ,;
nFromGap3   ,;
nGap1       ,;
nGap2       ,;
nGap3       ,;
nMinSold    ,;
nMonths     ,;
nWeeks      ,;
nPolnid     ,;
nQty        ,;
nQty_A      ,;
nQty_B      ,;
nSerFrom    ,;
nSerTo      ,;
nTo         ,;
nHoursTo    ,;
nToGap3     ,;
oBro        ,;
tFrom       ,;
tTo         ,;
cSonoFrom   ,;
cSonoTo     ,;
cSoordlnFrom,;
cSoordlnTo  ,;
cStage      ,;
cLoc        ,;
lM999       ,;
cFrom_To    ,;
lProm_ANV   ,;
lProm_Trans ,;
lProm_Sameesn,;
lProm_Pack,;
cPArea,;
lPromised   ;

PROCEDURE prnKillCrit
aRecNos      := NIL
cShipAs      := NIL
cAgency      := NIL
cArea        := NIL
cBNFrom      := NIL
cBNTo        := NIL
cBatchType   := NIL
cBid         := NIL
cTop         := NIL
cBloke       := NIL
cBstat       := NIL
cPBstat      := NIL
cNBstat      := NIL
cProj        := NIL
cBuffer      := NIL
cCompany     := NIL
cCountry     := NIL
cCustomer    := NIL
cDOCFrom     := NIL
cDialecter   := NIL
cDocTo       := NIL
cEngineer    := NIL
cESNXX       := NIL
cESNY        := NIL
cEsnFrom     := NIL
cEsnTo       := NIL
cMoveType    := NIL
cPOType      := NIL
cPcCodeFrom  := NIL
cPcCodeTo    := NIL
cPoStat      := NIL
cPriorFrom   := NIL
cPriorTo     := NIL
cPriority    := NIL
cProcFrom    := NIL
cProcIDFrom  := NIL
cProcIDTo    := NIL
cProcId      := NIL
cProcTo      := NIL
cProcType    := NIL
cBncaps      := NIL
cProductLine := NIL
cProductType := NIL
cPurpose     := NIL
cRefFrom     := NIL
cRefTo       := NIL
cRejectId    := NIL
cRemark      := NIL
cRoute       := NIL
cSValFile    := NIL
cSValmFile   := NIL
cSVMoveType  := NIL
cSize        := NIL
cTC          := NIL
cOma         := NIL
cRepUnit     := NIL
cEndCustomer := NIL
cTermination := NIL
cTol         := NIL
cValue       := NIL
cYesNo       := NIL
cVoltage     := NIL
cWarehouse   := NIL         //YG 29/09/97 for rfg07v1
cWareh_sr    := NIL
cWareh_ds    := NIL
cWorkStation := NIL
cWorkerId    := NIL
cWrkStnFrom  := NIL
cWrkStnTo    := NIL
dAllocFrom   := NIL
dAllocTo     := NIL
dBookedFrom  := NIL
dBookedTo    := NIL
dNonNon      := NIL
dDueFrom     := NIL
dDueTo       := NIL
dEFrom       := NIL
dETo         := NIL
dFinFrom     := NIL
dFinTo       := NIL
dFrom        := NIL
dMrkDueFrom  := NIL
dMrkDueTo    := NIL
dReqFrom     := NIL
dReqTo       := NIL
dRevFrom     := NIL
dRevTo       := NIL
dShippedFrom := NIL
dShippedTo   := NIL
dStartFrom   := NIL
dStartTo     := NIL
dTo          := NIL
nDays        := NIL
nFrom        := NIL
nHoursFrom   := NIL
nFromGap3    := NIL
nGap1        := NIL
nGap2        := NIL
nGap3        := NIL
nMinSold     := NIL
nMonths      := NIL
nPolnid      := NIL
nQty         := NIL
nQty_A       := NIL
nQty_B       := NIL
nSerFrom     := NIL
nSerTo       := NIL
nTo          := NIL
nHoursTo     := NIL
nToGap3      := NIL
oBro         := NIL
tFrom        := NIL
tTo          := NIL
cSonoFrom    := NIL
cSonoTo      := NIL
cSoordlnFrom := NIL
cSoordlnTo   := NIL
cStage       := NIL
cLoc         := NIL
lM999        := NIL
cFrom_To     := NIL
lProm_ANV    := NIL
lProm_Trans  := NIL
lProm_Sameesn:= NIL
lProm_Pack   := NIL
cPArea       := NIL
lPromised    := NIL
RETURN


STATIC PROCEDURE prnShow( cFile )
LOCAL nKey ,nRow, nCol, cSearcher, cChar ,nBroRow, i, j, lFound := .F.
cSearcher := ""
WHILE .T.
    WHILE !oBro:stabilize() ; ENDDO

    nKey := InKey(1)

    IF nKey = K_ESC
       Afill( aRecNos , "[ ]" )
       EXIT
    ELSEIF StdKeys( nKey , oBro )
    ELSEIF nKey = K_ENTER
       EXIT
    ELSEIF nKey = K_F7   // mark all
       Afill( aRecNos , "[X]" )
       oBro:refreshAll()
    ELSEIF nKey = K_SH_F7  // un mark all
       Afill( aRecNos , "[ ]" )
       oBro:refreshAll()
    ELSEIF nKey = K_ALT_F7  // invert
       FOR nKey := 1 TO Len( aRecNos )
           IF aRecNos[ nKey ] = "[X]"
              aRecNos[ nKey ] = "[ ]"
           ELSE
              aRecNos[ nKey ] = "[X]"
           ENDIF
       NEXT
       oBro:refreshAll()
    ELSEIF nKey = K_SPACE
       IF aRecNos[ (cFile)->( RecNo() ) ] = "[X]"
          aRecNos[ (cFile)->( RecNo() ) ] = "[ ]"
       ELSE
          aRecNos[ (cFile)->( RecNo() ) ] = "[X]"
       ENDIF
       oBro:refreshCurrent()
*****************************************
   ELSEIF (nKey >= 32 .AND. nKey <= 255)
          cChar := Upper( Chr( nKey ) )
          cSearcher += cChar
          cSearcher := searchIt(cSearcher,cFile)
   ELSEIF nKey = K_BS
          IF !Empty(cSearcher)
             cSearcher := SubStr(cSearcher,1,Len(cSearcher)-1)
             searchIt(cSearcher,cFile)
          END

*****************************************
    ENDIF

ENDDO
RETURN

*********************************
FUNCTION GetDate( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dFrom := dTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dFrom TO aParamList[1] , dTo TO aParamList[2]
 dFrom := aParamList[1]
 dTo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter date range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET dFrom  Valid Notempty(dFrom)
@ 5, 16 GET dTo    Valid Notempty(dTo)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
//////////////////////////vitaly 00122715/257
FUNCTION GetQtyChoice( o, aParamList,aMenuItems,cReport )

LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL aMenuAction
LOCAL nChoice

Default aMenuItems  to {"Promised Batches","Non promised Batches"}
Default cReport to " "

IF cReport == "NV Pcs to transfer"
	aMenuActions := { {|| lProm_ANV := "ABL" } , {|| lProm_ANV := "ARL" } }
ELSEIF cReport == "Actual transfer"
   aMenuActions := { {|| lProm_Trans := "Yes" } , {|| lProm_Trans := "No" } }
ELSEIF cReport == "Difference in the ESN's"
   aMenuActions := { {|| lProm_Sameesn := "Yes" } , {|| lProm_Sameesn := "No" } }
ELSEIF cReport == "Pack instruction's"
   aMenuActions := { {|| lProm_Pack := "Yes" } , {|| lProm_Pack := "No" } }
ELSE
   aMenuActions := { {|| lPromised := "Yes" } , {|| lPromised := "No" } }
ENDIF


IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb) .AND. Empty(cReport)
   lPromised := NIL
   RETURN .T.
ELSEIF !Eval(o:checkGsb) .AND. cReport == "NV Pcs to transfer"
   lProm_ANV := NIL
   RETURN .T.
ELSEIF !Eval(o:checkGsb) .AND. cReport == "Difference in the ESN's"
   lProm_Sameesn := NIL
   RETURN .T.
ELSEIF !Eval(o:checkGsb) .AND. cReport == "Pack instruction's"
   lProm_Pack := NIL
   RETURN .T.
ELSEIF !Eval(o:checkGsb) .AND. cReport == "Actual transfer"
   lProm_Trans := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF
if Empty(cReport)
   lPromised := aParamList[1]
elseif cReport == "NV Pcs to transfer"
   lProm_ANV := aParamList[1]
elseif cReport == "Actual transfer"
   lProm_Trans := aParamList[1]
elseif cReport == "Difference in the ESN's"
   lProm_Sameesn := aParamList[1]
elseif cReport == "Pack instruction's"
   lProm_Pack := aParamList[1]
endif

scrnPush()

DispBox( 4, 5 , 7 , 29 , B_SINGLE+" " )
nChoice := aChoice(5,6,6,28,aMenuItems)

IF nChoice <> 0
   EVAL(aMenuActions[nChoice])
ENDIF

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
   if Empty(cReport)
      lPromised := NIL
   elseif cReport == "NV Pcs to transfer"
      lProm_ANV := NIL
   elseif cReport == "Actual transfer"
      lProm_Trans := NIL
   elseif cReport == "Difference in the ESN's"
      lProm_Sameesn := NIL
   elseif cReport == "Pack instruction's"
      lProm_Pack := NIL
   endif
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

FUNCTION GetGapChoice( o, aParamList,aMenuItems )

LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL aMenuActions := { {|| lM999 := "Yes" } , {|| lM999 := "No" } }
LOCAL nChoice

Default aMenuItems  to {"M999 ����","M999 ���"}

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   lPromised := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

lPromised := aParamList[1]

scrnPush()
DispBox( 4, 5 , 7 , 29 , B_SINGLE+" " )
nChoice := aChoice(5,6,6,28,aMenuItems)
IF nChoice <> 0
   EVAL(aMenuActions[nChoice])
ENDIF

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ENDIF
scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

//////////////////////////
FUNCTION GetStage( o, aParamList ) // VR 27-01-01 TID:2127212
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL aMenuItems   := {"Arrive","Start","Finish"}
LOCAL aMenuActions := { {|| cStage := "A" } , {|| cStage := "S" } , {|| cStage := "F" } }
LOCAL nChoice
IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   lPromised := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

cStage := aParamList[1]

scrnPush()
DispBox( 4, 5 , 8 , 29 , B_SINGLE+" " )
nChoice := aChoice(5,6,7,28,aMenuItems)
IF nChoice <> 0
   EVAL(aMenuActions[nChoice])
ENDIF

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ENDIF
scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*********************************
FUNCTION GetDueDate( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dFrom := dTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dFrom TO aParamList[1] , dTo TO aParamList[2]
 dFrom := aParamList[1]
 dTo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter due date range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET dFrom  Valid Notempty(dFrom)
@ 5, 16 GET dTo    Valid Notempty(dTo)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*********************************
FUNCTION GetReqDate( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dReqFrom := dReqTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dFrom TO aParamList[1] , dTo TO aParamList[2]
dReqFrom := aParamList[1]
dReqTo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter Request date range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET dReqFrom  Valid Notempty(dReqFrom)
@ 5, 16 GET dReqTo    Valid Notempty(dReqTo)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*********************************
FUNCTION GetSlack( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   nFrom := nTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dFrom TO aParamList[1] , dTo TO aParamList[2]
nFrom := aParamList[1]
nTo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter start date range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET nFrom
@ 5, 16 GET nTo
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
*********************************
FUNCTION GetHours( o, aParamList ) //VR 27-01-01 TID:2127212
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   nHoursFrom := nHoursTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

nHoursFrom := aParamList[1]
nHoursTo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 35 , B_DOUBLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter hours at Rep.Stage:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET nHoursFrom PICTURE "99999.99"
@ 5, 16 GET nHoursTo   PICTURE "99999.99"
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*********************************
FUNCTION GetStartDate( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dstartFrom := dStartTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

dStartFrom := aParamList[1]
dStartTo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter start date range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET dStartFrom  Valid Notempty(dStartFrom)
@ 5, 16 GET dStartTo    Valid Notempty(dStartTo)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
*********************************
FUNCTION GetOrderDate( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dFrom := dTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dFrom TO aParamList[1] , dTo TO aParamList[2]
dFrom := aParamList[1]
dTo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter Order date range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET dFrom  Valid Notempty(dFrom)
@ 5, 16 GET dTo    Valid Notempty(dTo)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*********************************
FUNCTION GetGeneric(o,aParamList,cTitle,cFieldFrom,cFieldTo)
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   (cFieldFrom) := (cFieldTo) := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

(cFieldFrom) := aParamList[1]
(cFieldTo)   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY cTitle
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET (cFieldFrom)  //Valid Notempty((cFieldFrom))
@ 5, 16 GET (cFieldTo)    //Valid Notempty((cFieldTo))
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*********************************
FUNCTION GenericNumber(o,aParamList,cTitle,cFieldFrom,cFieldTo)
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   (cFieldFrom) := (cFieldTo) := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF


(cFieldFrom) := aParamList[1]
(cFieldTo)   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY cTitle
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET (cFieldFrom)  //Valid Notempty((cFieldFrom))
@ 5, 16 GET (cFieldTo)    //Valid Notempty((cFieldTo))
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
***************************************
FUNCTION GeTSONO(o,aParamList,cTitle,cFieldFrom,cFieldTo)
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   (cSonoFrom) := (cSonoTo) := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dFrom TO aParamList[1] , dTo TO aParamList[2]
(cSonoFrom) := aParamList[1]
(cSonoTo)   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 39 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY cTitle
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET (cSonoFrom)  //Valid Notempty((cFieldFrom))
@ 5, 22 GET (cSonoTo)    //Valid Notempty((cFieldTo))
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

***************************************
FUNCTION GeTSoordln(o,aParamList,cTitle,cFieldFrom,cFieldTo)
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   (cSoordlnFrom) := (cSoordlnTo) := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dFrom TO aParamList[1] , dTo TO aParamList[2]
(cSoordlnFrom) := aParamList[1]
(cSoordlnTo)   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 39 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY cTitle
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET (cSoordlnFrom)  //Valid Notempty((cFieldFrom))
@ 5, 22 GET (cSoordlnTo)    //Valid Notempty((cFieldTo))
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.


***************************************
FUNCTION GetAllocDate( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dDueFrom := dDueTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dFrom TO aParamList[1] , dTo TO aParamList[2]
 dallocFrom := aParamList[1]
 dallocTo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter alloc date range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET dallocFrom  Valid Notempty(dallocFrom)
@ 5, 16 GET dallocTo    Valid Notempty(dallocTo)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*********************************
FUNCTION GetRevDate( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dRevFrom := dRevTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dFrom TO aParamList[1] , dTo TO aParamList[2]
 dRevFrom := aParamList[1]
 dRevTo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 36 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter Revised due date range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET dRevFrom  Valid Notempty(dRevFrom)
@ 5, 16 GET dRevTo    Valid Notempty(dRevTo)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.


***************************************

FUNCTION GetMrkDue( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dMrkDueFrom := dMrkDueTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dFrom TO aParamList[1] , dTo TO aParamList[2]
 dMrkDueFrom := aParamList[1]
 dMrkDueTo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 36 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter marketing due date range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET dMrkDueFrom  Valid Notempty(dMrkDueFrom)
@ 5, 16 GET dMrkDueTo    Valid Notempty(dMrkDueTo)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

***************************************

FUNCTION GetBookedDate( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dBookedFrom := dBookedTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dBookedFrom TO aParamList[1] , dBookedTo TO aParamList[2]
 dBookedFrom := aParamList[1]
 dBookedTo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter date range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET dBookedFrom  Valid Notempty(dBookedFrom)
@ 5, 16 GET dBookedTo    Valid Notempty(dBookedTo)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

FUNCTION GetEDate( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dEFrom := dETo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

 dEFrom := aParamList[1]
 dETo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter date range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET dEFrom  Valid Notempty(dEFrom)
@ 5, 16 GET dETo    Valid Notempty(dETo)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.


FUNCTION GetNonDate( o, aParamList ) // Vitaly 10-06-02 ID 2610845
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dNonNon := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dBookedFrom TO aParamList[1] , dBookedTo TO aParamList[2]
dNonNon := aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter Non-Non date:"
@ 4, 6  SAY "From:     "
@ 5, 6  GET dNonNon  Valid Notempty(dNoNNoN)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

****************************************
FUNCTION GetShippedDate( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dShippedFrom := dShippedTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dShippedFrom TO aParamList[1] , dShippedTo TO aParamList[2]
 dShippedFrom := aParamList[1]
 dShippedTo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter date range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET dShippedFrom  Valid Notempty(dShippedFrom)
@ 5, 16 GET dShippedTo    Valid Notempty(dShippedTo)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
*********************************
FUNCTION closeFromTo( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dFrom := dTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dFrom TO aParamList[1] , dTo TO aParamList[2]
 dFrom := aParamList[1]
 dTo   := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Close dates:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET dFrom  Valid Notempty(dFrom)
@ 5, 16 GET dTo    Valid Notempty(dTo)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*************************************
FUNCTION GetBatch( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cBid := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cBid TO aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter Batch number:"

@ 5, 6  GET cBid PICTURE "999999"
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*************************************
FUNCTION GetGuy( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cBid := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cBloke TO aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter guy or gal of your choice:"

@ 5, 6  GET cBloke PICTURE "!AAAAAAAAA"
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
*************************************
FUNCTION GetPolnid( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cBid := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT nPolnid TO aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter Order number:"

@ 5, 6  GET nPolnid PICTURE "9999999.99"
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

***************************************
FUNCTION GetFinishDate( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dFrom := dTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

dFrom := aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter finish date:"
@ 5, 6  GET dFrom
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

***************************************
FUNCTION GetToDate( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dTo := tTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

//DEFAULT dFrom TO aParamList[1]
 dTo := aParamList[1]
 tTo := aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter finish date:"
@ 5, 6 GET dTo Valid Notempty(dTo)
@ 5, 16 GET tTo PICTURE "99:99"
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*************************************
FUNCTION GetNoOfDays( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   nDays := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT nDays TO aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter days limit:"

@ 5, 6  GET nDays PICTURE "999" VALID nDays > 0
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*************************************
FUNCTION GetMonths( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   nDays := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT nMonths TO aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter # of months:"

@ 5, 6  GET nMonths PICTURE "999"
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*************************************
FUNCTION GetWeeks( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   nDays := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT nWeeks TO aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 33 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 2  SAY "Enter # of weeks for profile:"

@ 5, 2  GET nWeeks PICTURE "99"
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

FUNCTION GetNoOfGaps( o, aParamList)
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   nDays := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT nGap1 TO aParamList[1]
DEFAULT nGap2 TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 9 , 46 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter Entry-Allocation limit:"
@ 4, 6  GET nGap1 PICTURE "999"
@ 5, 6  SAY "Enter Allocation-Shipment limit:"
@ 6, 6  GET nGap2 PICTURE "999"
@ 7, 6  say "Records which satisfy either of the"
@ 8, 6  say "above criteria will be included."
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

FUNCTION GetNo2OfGaps( o, aParamList)
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   nDays := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT nGap2 TO aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter days limit:"

@ 5, 6  GET nGap2 PICTURE "999"
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

/*********************************/
FUNCTION GetNo3OfGaps( o, aParamList)
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   nDays := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT nFromGap3 TO aParamList[1]
DEFAULT nToGap3 TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 9 , 36 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter from gap limit:"
@ 4, 6  GET nFromGap3 PICTURE "9999"
@ 6, 6  SAY "Enter to gap limit:"
@ 7, 6  GET nToGap3 PICTURE "9999"
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*************************************
FUNCTION GetRoute( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cRoute := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cRoute TO aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}

@ 3, 6  SAY "Enter route filter:"
@ 5, 6  GET cRoute PICTURE "!!!!"
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
/*********************************/
FUNCTION DaysOnLine( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   nDays := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT nDays TO aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter max days on line:"

@ 5, 6  GET nDays PICTURE "999"
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

/*********************************/
//YG 29.9.97
//Used by rfgs07v1.prg for Inventory cardfile
//YG 05.10.97 didn't use this function in the end
// used new dbf instead.
FUNCTION GetWarehouse( o, aParamList)
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor
Local i
local aWarehouse := {.T.,.T.,.T.,.T.,.T.,.T.}

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cWarehouse := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cWarehouse to aParamList[1]+"_"+aParamList[2]+;
                         "_"+aParamList[3]+"_"+aParamList[4]+"_"+aParamList[5];
                         +"_"+aParamList[6]

scrnPush()
DispBox( 1, 1 , 9 , 29 , B_SINGLE+" " )
@ 1,4 SAY "Include warehouse:(T/F)" color GETCOLORS
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
FOR i := 1 TO len(aParamlist)
     @ 1+i,2 say aParamlist[i] color GETCOLORS
     @ 1+i,8 get aWarehouse[i] color GETCOLORS
NEXT

XREAD
cWarehouse := " "
FOR i := 1 TO len(aParamlist)
     if aWarehouse[i]
          cWarehouse += aParamlist[i]+"_"
     end
     cWarehouse += "WH9"
NEXT

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*********************************
//YG 9/97
//Used by racc06v1.prg for dead/slow stock report
FUNCTION GetMinSold( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   nMinSold := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT nMinSold TO aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter Minimum Sold:"

@ 5, 6  GET nMinSold PICTURE "999999"
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*********************************
//YG 11/98
FUNCTION QtyGreaterThan( o, aParamList ,cTitle,aXY,nMode)
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

default aXY to {1, 1 , 7 , 29}

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF
IF !Eval(o:checkGsb)
   nQty := NIL
   RETURN .T.
ENDIF
IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF
DEFAULT nQty TO aParamList[1]
scrnPush()
DispBox( aXY[1],aXY[2],aXY[3],aXY[4], B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY cTitle
@ 5, 6  GET nQty PICTURE "9999999"
XREAD
IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF
scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
IF nMode == 1
	nQty_A := nQty
	nQty := 0
ELSEIF nMode == 2
	nQty_B := nQty
	nQty := 0
ENDIF
RETURN .T.

*********************************
FUNCTION GetTime( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   tFrom := tTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT tFrom TO aParamList[1] , tTo TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter time range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET tFrom PICTURE "99:99"
@ 5, 16 GET tTo   PICTURE "99:99"
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE

ENDIF

scrnPop()
GetList := aGetList
nCursor := SetCursor( nCursor )
SetPos( nRow , nCol )
RETURN .T.

********************************
FUNCTION GetEsn( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor, lRet


IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cBNFrom := cBNTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cEsnFrom TO aParamList[1] , cEsnTo TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 39 , B_SINGLE+" " )

nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 2 SAY "Enter ESN:"
@ 4, 2 SAY "From:             To:"
@ 5, 2 GET cEsnFrom  PICT "@!"                 ;
          SEND PreBlock  :=  {|o| PrePick( o, "D_ESN")}                                 ;
          SEND PostBlock :=  {|o| lret := PostPick(o,"D_ESN","ESN_id",,"iesn_id"),;
          IF(lRet,getlist[2]:Varput( o:varGet()),NIL), getlist[2]:display(),lRet}        ;
          COLOR GETCOLORS
@ 5, 20 GET cEsnTo   PICT "@!"                                                  ;
          SEND PreBlock  :=  {|o| PrePick( o, "D_ESN")}                           ;
          SEND PostBlock :=  {|o| PostPick(o,"D_ESN","ESN_id",,"iesn_id")} ;
          COLOR GETCOLORS
XREAD


IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ENDIF

cEsnTo  := Padr(alltrim(cEsnTo),16,"Z")
GetList := aGetList
nCursor := SetCursor( nCursor )
SetPos( nRow , nCol )
scrnPop()
RETURN .T.




FUNCTION GetEngineer( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor, lRet


IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cBNFrom := cBNTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cEngineer TO aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )

nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 2 SAY "Enter Engineer:"
@ 4, 2 GET cEngineer  PICT "@!"                 ;
          COLOR GETCOLORS
XREAD


IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ENDIF

GetList := aGetList
nCursor := SetCursor( nCursor )
SetPos( nRow , nCol )
scrnPop()
RETURN .T.

********************************************
FUNCTION GetBN( o, aParamList,cFile,cIfile )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor, lRet

Default cFile To "D_line"
Default cIFile To "ib_idln"

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cBNFrom := cBNTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cBNFrom TO aParamList[1] , cBNTo TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )

nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter B/N range:"
@ 4, 6 SAY "From:     To:"
@ 5, 6 GET cBNFrom PICT '999999'                                                         ;
          SEND PreBlock  :=  {|o| PrePick( o, cFile)}                                 ;
          SEND PostBlock :=  {|o|postBN(o),lRet:= PostPick(o,cFile,"B_id",,cIfile),;
          IF(lRet,getlist[2]:Varput( o:varGet()),NIL), getlist[2]:display(),lRet}        ;
          COLOR GETCOLORS
@ 5, 16 GET cBNTo   PICT '999999'                                                  ;
          SEND PreBlock  :=  {|o| PrePick( o, cFile)}                           ;
          SEND PostBlock :=  {|o| postBN(o),PostPick(o,cFile,"B_id",,cIfile)};
          COLOR GETCOLORS
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ENDIF

GetList := aGetList
nCursor := SetCursor( nCursor )
SetPos( nRow , nCol )
scrnPop()
RETURN .T.

********************************************
FUNCTION GetDocNo( o, aParamList,cFile,cIfile )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor, lRet

Default cFile To "D_SHIPH"
Default cIFile To "ishipdhn"

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF
IF !Eval(o:checkGsb)
   cDOCFrom := cDOCTo := NIL
   RETURN .T.
ENDIF
IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF
DEFAULT cDOCFrom TO aParamList[1] , cDOCTo TO aParamList[2]
scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6 SAY "Enter Invoice range:"
@ 4, 6 SAY "From:     To:"
@ 5, 6 GET cDOCFrom PICT '999999'                                      ;
          SEND PreBlock  :=  {|o| PrePick( o, cFile ) }                ;
          SEND PostBlock :=  {|o| lRet := PostPick(o,cFile,"DOC_NO",,cIfile),  ;
          IF(lRet,getlist[2]:Varput( o:varGet()),NIL), getlist[2]:display(),lRet} ;
          COLOR GETCOLORS
@ 5, 16 GET cDOCTo   PICT '999999'                                     ;
          SEND PreBlock  :=  {|o| PrePick( o, cFile ) }                ;
          SEND PostBlock :=  {|o| PostPick(o,cFile,"doc_no",,cIfile)};
          COLOR GETCOLORS
XREAD
IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ENDIF
GetList := aGetList
nCursor := SetCursor( nCursor )
SetPos( nRow , nCol )
scrnPop()
RETURN .T.

**********************************
FUNCTION GetOneBN( o, aParamList )
//special GET for QFGS30  query
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor,lRet


IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cBNFrom := cBNTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cBNFrom TO aParamList[1] , cBNTo TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )

nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter B/N range:"
@ 4, 6 SAY "From:     To:"
@ 5, 6 GET cBNFrom PICT '999999'   ;
          SEND PreBlock  :=  {|o| PrePick( o, "m_stkmv")}                        ;
          SEND PostBlock :=  {|o|postBN(o),lRet:= PostPick(o,"m_stkmv","srB_id",,"istmvbdt"),;
          IF(lRet,getlist[2]:Varput( o:varGet()),NIL), getlist[2]:display(),lRet};
          COLOR GETCOLORS
@ 5, 16 GET cBNTo   PICT '999999';
          SEND PreBlock  :=  {|o| PrePick( o, "m_stkmv")}                        ;
          SEND PostBlock :=  {|o| postBN(o),PostPick(o,"m_stkmv","srB_id",,"istmvbdt")};
          COLOR GETCOLORS
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ENDIF

GetList := aGetList
nCursor := SetCursor( nCursor )
SetPos( nRow , nCol )
scrnPop()
RETURN .T.

********************************
FUNCTION GetPcNo( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor, lRet

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cProcFrom := cProcTo:= NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cProcFrom TO aParamList[1] , cProcTo TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 45 , B_SINGLE+" " )

nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 2 SAY "Enter Process:"
@ 4, 2 SAY "From:                To:"
@ 5, 2 GET cProcFrom                   ;
          SEND PreBlock  :=  {|o| PrePick( o, "c_proc",{5,2,1},NIL,NIL,"pcprocno",NIL,NIL,NIL,5)};
          SEND PostBlock :=  {|o| PostPick(o,"c_proc","pcprocno",,"pcprocno")};
          COLOR GETCOLORS
//{|| c_proc->pcprocno}, {|| c_proc->proc_nmh }
@ 5, 23 GET cProcTo                                       ;
          SEND PreBlock  :=  {|o| PrePick( o, "c_proc",{5,2,1},NIL,NIL,"pcprocno",NIL,NIL,NIL,5)};
          SEND PostBlock :=  {|o| PostPick(o,"c_proc","pcprocno",,"pcprocno")} ;
          COLOR GETCOLORS
XREAD


IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ENDIF

GetList := aGetList
nCursor := SetCursor( nCursor )
SetPos( nRow , nCol )
alert("PC code will now be disregarded")
scrnPop()
RETURN .T.


********************************
FUNCTION GetPcName( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor, lRet

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cProcFrom := cProcTo:= NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cProcFrom TO aParamList[1] , cProcTo TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 45 , B_SINGLE+" " )

nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 2 SAY "Enter Process:"
@ 4, 2 SAY "From:                To:"
@ 5, 2 GET cProcFrom                   ;
          SEND PreBlock  :=  {|o| PrePick( o, "c_proc",{1,2,3},NIL,NIL,NIL,NIL,NIL,NIL,2)};
          SEND PostBlock :=  {|o| PostPick(o,"c_proc","proc_nme",,"procname")};
          COLOR GETCOLORS
//{|| c_proc->pcprocno}, {|| c_proc->proc_nmh }
@ 5, 23 GET cProcTo                                       ;
          SEND PreBlock  :=  {|o| PrePick( o, "c_proc",{1,2,3},NIL,NIL,NIL,NIL,NIL,NIL,2)};
          SEND PostBlock :=  {|o| PostPick(o,"c_proc","proc_nme",,"procname")} ;
          COLOR GETCOLORS
XREAD


IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ENDIF

GetList := aGetList
nCursor := SetCursor( nCursor )
SetPos( nRow , nCol )
alert("PC code will now be disregarded")
scrnPop()
RETURN .T.


//STATIC FUNCTION pre_Check(o,cAlias)
//SetF2Key( o , cAlias)
//RETURN .T.
********************************
FUNCTION GetRemark( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor, lRet


IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cRemark := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cRemark TO aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 59 , B_SINGLE+" " )

nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 2 SAY "Remark:"
@ 4, 2 GET cRemark;
          SEND reader := {|o| HebReader(o)};
          COLOR GETCOLORS
XREAD


IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ENDIF

cRemark := alltrim(cRemark)
GetList := aGetList
nCursor := SetCursor( nCursor )
SetPos( nRow , nCol )
scrnPop()
RETURN .T.

********************************************
FUNCTION GetDialect( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor, lRet


IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cRemark := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cDialecter TO aParamList[1]

scrnPush()
DispBox( 1, 1 , 7 , 59 , B_SINGLE+" " )

nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 2 SAY "Dialecter type (S / N): "
@ 3, 25 GET cDialecter picture "@!";
          send postblock := {|o| o:varGet() $ "S_N" };
          COLOR GETCOLORS
XREAD


IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ENDIF

GetList := aGetList
nCursor := SetCursor( nCursor )
SetPos( nRow , nCol )
scrnPop()
RETURN .T.

********************************************

*************************
STATIC FUNCTION postBN(o)
LOCAL lRetVal
LOCAL cBuffer := StrZero(Val(o:varGet()),6)

IF Empty( cBuffer )
   Msg24( {"10"} , 3 , .T. )
   RETURN .F.
ENDIF

   lRetVal := .T.
   o:varPut( cBuffer )
   o:display()

RETURN lRetVal
**********************************
FUNCTION GetRefNo( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cRefFrom := cRefTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cRefFrom TO aParamList[1] , cRefTo TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter Ref range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET cRefFrom PICT 'XXXXXXXXX' valid putrefto(getlist)
@ 5, 16 GET cRefTo   PICT 'XXXXXXXXX'
XREAD

cRefFrom:=Padl(Alltrim(cRefFrom),9)
cRefTo:=  Padl(Alltrim(cRefTo)  ,9)


IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ENDIF

GetList := aGetList
nCursor := SetCursor( nCursor )
SetPos( nRow , nCol )
scrnPop()
RETURN .T.
*******************************
STATIC FUNCTION PUTREFTO(alist)

aList[2]:VarPut( aList[1]:VarGet() )
aList[2]:Display()
RETURN .T.
**********************************
FUNCTION GetSerNo( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor


IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   nSerFrom := nSerTo := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT nSerFrom TO aParamList[1] , nSerTo TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter Ser range:"
@ 4, 6  SAY "From:     To:"
@ 5, 6  GET nSerFrom PICT '999999999'
@ 5, 16 GET nSerTo   PICT '999999999'

XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ENDIF

GetList := aGetList
nCursor := SetCursor( nCursor )
SetPos( nRow , nCol )
scrnPop()
RETURN .T.

**************************************************************************
FUNCTION critBrowse( o, aParamList, cFile, cIndex, bKeyCol, bNameCol, bFilter)
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
default bFilter to ""
IF o:ExitState != GE_NOEXIT

   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   prnSetCritBuffer( cFile , .T., o:VarGet() )
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF
scrnPush()

IF NetUse(cFile , 5 )
   OpenIndex( cIndex , cFile )
ELSE
    scrnPop()
   RETURN .F.
ENDIF
IF !empty(bFilter) //YG
     (cFile)->(dbsetfilter(bFilter))
     (cFile)->(dbgotop())
ENDIF
IF (cFile)->( Eof() ) .AND. (cFile)->( Bof() )
   prnSetCritBuffer( cFile , .T.,o:VarGet() )
   Eval(o:checkGsb , .F. )
   NETCLOSE (cFile)
   Tone(400,1)
   scrnPop()
   RETURN .F.
ELSE
   aRecNos := Array( (cFile)->( LastRec() ) )
ENDIF

IF aParamList[1] == "all"
   Afill( aRecNos , "[X]" )
ENDIF

oBro := (cFile)->( TBrowseDB( 2, 2 , 15 , 78 ) )

DispBox( 1, 1 , 19 , 79 , B_SINGLE+" " )

oBro:addColumn( TBColumnNEW( "*" , {|| aRecNos[(cFile)->( RecNo() )]}))
IF UPPER(cFile) == "C_TOL"
   oBro:addColumn( TBColumnNEW( "Tol" , bKeyCol ))
   oBro:addColumn( TBColumnNEW( "Part;Type" , {|| (cFile)->ptype_id } ))
/*elseif UPPER(cFile) == "C_PTYPE"
   oBro:addColumn( TBColumnNEW( "ptype_id" , bKeyCol ))
   oBro:addColumn( TBColumnNEW( "Description" , {|| IIF((cFile)->ptype_id<>'K',(cFile)->ptype_nm,(cFile)->cust_ptype) } ))*/
elseif UPPER(cFile) == "D_SVFLS"
   oBro:addColumn( TBColumnNEW( "Rep. name" , bKeyCol ))
   oBro:addColumn( TBColumnNEW( "Date" , {|| (cFile)->dsval } ))
elseif UPPER(cFile) == "D_SVMFLS"
   oBro:addColumn( TBColumnNEW( "Rep. name" , bKeyCol ))
   oBro:addColumn( TBColumnNEW( "Dates" , {|| dtoc((cFile)->OPbaldate)+"-"+dtoc((cFile)->ENDbaldate) } ))
ELSE
   oBro:addColumn( TBColumnNEW( "   " , bKeyCol ))
ENDIF
if UPPER(cFile) == "D_SVFLS"
   oBro:addColumn( TBColumnNEW( "Time" , {|| (cFile)->tsval } ))
else
     oBro:addColumn( TBColumnNEW( "Description                                                 " , bNameCol ))
end

@ 17,3 SAY Padr( "Select/Unselect :[Space] Select All :[F7]  Unselect All:[Shift]+[F7] " , 75 ) COLOR "GR+/RB"
@ 18,3 SAY Padr( "Invert Selection :[Alt]+[F7]  Up:[] Down:[] Exit:[Enter]" , 75 ) COLOR "GR+/RB"
prnShow( cFile )

IF LastKey() = K_ENTER
   cBuffer := ""
   (cFile)->( DbGoTop() )
   WHILE !(cFile)->( Eof() )
       IF aRecNos[(cFile)->( RecNo())] == "[X]"
			 IF cFile == "c_prornd"//special case for Racc07v1 report VR:01.01.03
				 cBuffer += ALLTRIM(Eval( bKeyCol ))  +"_"
			 ELSE
             cBuffer += Eval( bKeyCol )  +"_"
			 ENDIF
       ENDIF
       (cFile)->( DbSkip() )
   ENDDO
   prnSetCritBuffer(cFile , .F.,o:VarGet() )
ELSE
   cBuffer := ""
   prnSetCritBuffer(cFile , .T.,o:VarGet() )
   Eval(o:checkGsb , .F. )
ENDIF

NETCLOSE (cFile)
scrnPop()
SetPos( nRow, nCol )
RETURN .T.

PROCEDURE prnSetCritBuffer(cFile , lCode, cBUFF)

DO CASE
   CASE cFile == "c_ptype" .OR. cFile == "Product type"
        IF lCode
           cProductType := ""
        ELSE
           cProductType := cBuffer
        ENDIF
   CASE cFile == "c_pline" .OR. cFile == "Product line"
        IF lCode
           cProductLine := ""
        ELSE
           cProductLine := cBuffer
        ENDIF
   CASE cFile == "c_size" .OR. cFile == "Size"
        IF lCode
           cSize := ""
        ELSE
           cSize := cBuffer
        ENDIF
   CASE cFile == "c_value" .OR. cFile == "Value"
        IF lCode
           cValue := ""
        ELSE
           cValue := cBuffer
        ENDIF
   CASE cFile == "c_yesno" .OR. cFile == "Split batch" //Vitaly 2-9-02 ID:29121337
        IF lCode
           cYesNo := ""
        ELSE
           cYesNo := cBuffer
        ENDIF
   CASE cFile == "c_prornd" .OR. cFile == "Project"
        IF lCode
           cProj := ""
        ELSE
           cProj := cBuffer
        ENDIF
   CASE cFile == "c_tol" .OR. cFile == "Tolerance"
        IF lCode
           cTol := ""
        ELSE
           cTol := cBuffer
        ENDIF
   CASE cFile == "c_volt"  .OR. cFile == "Voltage"
        IF lCode
           cVoltage := ""
        ELSE
           cVoltage := cBuffer
        ENDIF
   CASE cFile == "c_tc" .OR. cFile == "TC"
        IF lCode
           cTC := ""
        ELSE
           cTC := cBuffer
        ENDIF
   CASE cFile == "omaesn" .OR. cFile == "OMARIM Parts"
        IF lCode
           cOma := ""
        ELSE
           cOma := cBuffer
        ENDIF
   CASE cFile == "c_repunt" .OR. cFile == "Rep.Unit"
        IF lCode
           cRepUnit := ""
        ELSE
           cRepUnit := cBuffer
        ENDIF
   CASE cFile == "c_term" .OR. cFile == "Termination"
        IF lCode
           cTermination := ""
        ELSE
           cTermination := cBuffer
        ENDIF
   CASE cFile == "c_esnxx" .OR. cFile == "ESN(XX)"
        IF lCode
           cESNXX := ""
        ELSE
           cESNXX := cBuffer
        ENDIF
   CASE cFile == "c_esny" .OR. cFile == "ESN(Y)"
        IF lCode
           cESNY := ""
        ELSE
           cESNY := cBuffer
        ENDIF
   CASE cFile == "c_btype" .OR. cFile == "Batch type"
        IF lCode
           cBatchType := ""
        ELSE
           cBatchType := cBuffer
        ENDIF
   CASE cFile == "c_stkmv" .OR. cFile == "Movement type"
        IF lCode
           cMoveType := ""
        ELSE
           cMoveType := cBuffer
        ENDIF
   CASE cFile == "c_svalmv" .OR. cFile == "Move type"
        IF lCode
           cSVMoveType := ""
        ELSE
           cSVMoveType := cBuffer
        ENDIF
   CASE cFile == "c_potype" .OR. cFile == "P/O Type"
        IF lCode
           cPOtype := ""
        ELSE
           cPOtype := cBuffer
        ENDIF

   CASE cFile == "c_PoStat" .OR. cFile == "P/O Status"
        IF lCode
           cPOstat := ""
        ELSE
           cPOstat := cBuffer
        ENDIF
   CASE cFile == "c_doctyp" .OR. cFile == "From-To"
        IF lCode
           cFrom_To := ""
        ELSE
           cFrom_To := cBuffer
        ENDIF
   CASE cFile == "c_sarea" .OR. cFile == "Area"
        IF lCode
           cArea := ""
        ELSE
           cArea := cBuffer
        ENDIF
   CASE cFile == "c_warea" .OR. cFile == "Production Area"
        IF lCode
           cPArea := ""
        ELSE
           cPArea := cBuffer
        ENDIF
   CASE cFile == "c_cntry" .OR. cFile == "Country"
        IF lCode
           cCountry := ""
        ELSE
           cCountry := cBuffer
        ENDIF
   CASE cFile == "c_agency" .OR. cFile == "Billing Agency"
        IF lCode
           cAgency:= ""
        ELSE
           cAgency := cBuffer
        ENDIF
   CASE cFile == "c_cuco" .OR. cFile == "Company"
        IF lCode
           cCompany := ""
        ELSE
           cCompany := cBuffer
        ENDIF
   CASE (cFile == "c_cust" .OR. cFile == "Customer") .AND. (cBUFF !="End Customer")
        IF lCode
           cCustomer := ""
        ELSE
           cCustomer := cBuffer
        ENDIF
   CASE cFile == "c_cust" .OR. cFile == "End Customer"
        IF lCode
           cEndCustomer := ""
           //cCustomer := ""
        ELSE
           cEndCustomer := cBuffer
           //cCustomer := cBuffer
        ENDIF
   CASE cFile == "Minimum sold"    /*YG 9/97 for racc06v1.prg*/
        IF lCode
            nMinSold := 0
        ELSE
            nMinSold := cBuffer
        ENDIF
   CASE cFile == "Quantity greater than"   .OR. ;
		  cFile == "Scrap cost greater than" .OR. ;
		  cFile == "Var from Std. greater than"
        IF lCode
            nQty := 0
        ELSE
            nQty := cBuffer
        ENDIF
   CASE cFile == "Scrap cost greater than"
        IF lCode
            nQty_A := 0
        ELSE
            nQty_A := cBuffer
        ENDIF
   CASE cFile == "Var from Std. greater than"
        IF lCode
            nQty_B := 0
        ELSE
            nQty_B := cBuffer
        ENDIF
   CASE (cFile == "Warehouse"  .OR. cFile == "c_wareh") .AND. ;
			cBUFF == "Warehouse"/*YG 9/97 for racc06v1.prg*/
        IF lCode
            cWarehouse  := ""
        ELSE
            cWarehouse := cBuffer
        ENDIF
   CASE cBUFF == "Source WH" .OR. cFile == "Source WH"
        IF lCode
            cWareh_sr  := ""
        ELSE
            cWareh_sr := cBuffer
        ENDIF
   CASE cBUFF == "Destination WH" .OR. cFile == "Destination WH"
        IF lCode
            cWareh_ds  := ""
        ELSE
            cWareh_ds := cBuffer
        ENDIF
   CASE cFile == "Location"  .OR. cFile == "c_loc"  /*YG 9/97 for racc06v1.prg*/
        IF lCode
            cLoc  := ""
        ELSE
            cLoc := cBuffer
        ENDIF


   CASE cFile == "B/N from...to"
        IF lCode
           cBNFrom  := ""
           cBNTo    := ""
        ELSE
           cBNFrom  := cBuffer[1]
           cBNTo    := cBuffer[2]
        ENDIF
   CASE cFile == "Date from...to" .OR. cFile == "Date into Warehouse";
          .OR. cFile == "Order date from...to"
        IF lCode
           dFrom  := ""
           dTo    := ""
        ELSE
           dFrom  := cBuffer[1]
           dTo    := cBuffer[2]
        ENDIF
   CASE cFile == "Start date from...to"
        IF lCode
           dStartFrom  := ""
           dStartTo    := ""
        ELSE
           dStartFrom  := cBuffer[1]
           dStartTo    := cBuffer[2]
        ENDIF
   CASE cFile == "DUE Date from...to" .OR. cFile == "Arrive/finish dates"  .OR.  cFile == "Finish dates"
        IF lCode
           dFrom  := ""
           dTo    := ""
        ELSE
           dFrom  := cBuffer[1]
           dTo    := cBuffer[2]
        ENDIF
   CASE cFile == "Allocation from...to"
        IF lCode
           dallocFrom  := ""
           dallocTo    := ""
        ELSE
           dallocFrom  := cBuffer[1]
           dallocTo    := cBuffer[2]
        ENDIF
   CASE cFile == "Revised due date from...to"
        IF lCode
           dRevFrom  := ""
           dRevTo    := ""
        ELSE
           dRevFrom  := cBuffer[1]
           dRevTo    := cBuffer[2]
        ENDIF
   CASE cFile == "Request date from...to"
        IF lCode
           dReqFrom  := ""
           dReqTo    := ""
        ELSE
           dReqFrom  := cBuffer[1]
           dReqTo    := cBuffer[2]
        ENDIF
   CASE cFile == "Marketing due date from...to"
        IF lCode
           dMrkDueFrom  := ""
           dMrkDueTo    := ""
        ELSE
           dMrkDueFrom  := cBuffer[1]
           dMrkDueTo    := cBuffer[2]
        ENDIF
   CASE cFile == "Sales period" .OR. cFile == "Billed from...to" .OR.;
               cFile == "Cancel date from...to"
        IF lCode
           dFrom  := ""
           dTo    := ""
        ELSE
           dFrom  := cBuffer[1]
           dTo    := cBuffer[2]
        ENDIF
        /*
   CASE cFile == "Booked from...to"
        IF lCode
           dFrom  := ""
           dTo    := ""
        ELSE
           dFrom  := cBuffer[1]
           dTo    := cBuffer[2]
        ENDIF
        */
   CASE cFile == "Time from...to"
        IF lCode
           tFrom  := ""
           tTo    := ""
        ELSE
           tFrom  := cBuffer[1]
           tTo    := cBuffer[2]
        ENDIF
   CASE cFile == "Invoice from...to"
        IF lCode
           cDOCFrom  := ""
           cDOCTO    := ""
        ELSE
           cDocFrom  := cBuffer[1]
           cDocTo    := cBuffer[2]
        ENDIF
   CASE cFile == "Date time finish"
        IF lCode
           dFrom := ""
           tFrom := ""
           dTo   := ""
           tTo   := ""
        ELSE
           dFrom := cBuffer[1]
           tFrom := cBuffer[2]
           dTo   := cBuffer[3]
           tTo   := cBuffer[4]
        ENDIF
   CASE cFile == "Specific Process"
        IF lCode
           cProcId := ""
        ELSE
           cProcId := cBuffer
        ENDIF
   CASE cFile == "Processes"    .OR. cFile == "c_proc"
        IF lCode
           cProcId := ""
        ELSE
           cProcId := cBuffer
        ENDIF
   CASE cFile == "Process"
        IF lCode
           cProcIdFrom := ""
           cProcIdTo   := ""
        ELSE
           cProcIdFrom := cBuffer[1]
           cProcIdTo   := cBuffer[2]
        ENDIF
   CASE cFile == "Process name"
        IF lCode
           cProcFrom := ""
           cProcTo   := ""
        ELSE
           cProcFrom := cBuffer[1]
           cProcTo   := cBuffer[2]
        ENDIF
   CASE cFile == "Process type" .OR. lower(cFile) == "c_prtype"  //YG 99011310 23/3/99
        IF lCode
           cProcType := ""
        ELSE
           cProcType := cBuffer
        ENDIF
   CASE cFile == "Caps" .OR. lower(cFile) == "c_bncaps"  //vr tapi 373
        IF lCode
           cBncaps := ""
        ELSE
           cBncaps := cBuffer
        ENDIF

   CASE cFile == "Pc_code"
        IF lCode
           cPcCodeFrom := "0000"
           cPcCodeTo   := "9999"
        ELSE
           cPcCodeFrom := cBuffer[1]
           cPcCodeTo   := cBuffer[2]
        ENDIF
   CASE cFile == "Pc_code and priority"
        IF lCode
           cPcCodeFrom := "0000"
           cPcCodeTo   := "9999"
           cPriorFrom  := ""
           cPriorTo  := ""
        ELSE
           cPcCodeFrom := cBuffer[1]
           cPcCodeTo   := cBuffer[2]
           cPriorFrom  := cBuffer[3]
           cPriorTo  := cBuffer[4]
        ENDIF
   CASE cFile == "Pc_code and date"
        IF lCode
           cPcCodeFrom := "0000"
           cPcCodeTo   := "9999"
           dFinFrom    := ctod("  /  /  ")
           dFinTo     := ctod("  /  /  ")
        ELSE
           cPcCodeFrom := cBuffer[1]
           cPcCodeTo   := cBuffer[2]
           dFinFrom   := cBuffer[3]
           dFinTo     := cBuffer[4]
        ENDIF
     CASE cFile == "Specific workstation"
         IF lCode
            cWorkStation := ""
         ELSE
            cWorkStation := cBuffer
         ENDIF

   CASE cFile == "Workstations" .OR. cFile == "c_wkstn"
         IF lCode
            cWorkStation := ""
         ELSE
            cWorkStation := cBuffer
         ENDIF

   CASE cFile == "Workstation"
        IF lCode
           cWrkStnfrom  := ""
           cWrkStnTo     := ""
        ELSE
           cWrkStnfrom  := cBuffer[1]
           cWrkStnTo     := cBuffer[2]
        ENDIF
   CASE cFile == "Ser# from...to"
        IF lCode
           nSerFrom := 0
           nSerTo   := 0
        ELSE
           nSerFrom := cBuffer[1]
           nSerTo   := cBuffer[2]
        ENDIF
   CASE cFile == "Scheduled finish date"
        IF lCode
           dFrom := NIL
        ELSE
           dFrom := cBuffer
        ENDIF
   CASE cFile == "Entry-allocation gap"
        IF lCode
           nGap1 := 0
        ELSE
           nGap1 := cBuffer
        ENDIF

   CASE cFile == "Entry-alloc-ship gaps"
        IF lCode
           nGap1 := 0
           nGap2 := 0
        ELSE
           nGap1 := cBuffer[1]
           nGap2 := cBuffer[2]
        ENDIF

CASE cFile == "Months booked" .OR. cFile == "Months"
        IF lCode
           nMonths := 0
        ELSE
           nMonths := cBuffer
        ENDIF

CASE cFile == "No of days at current proc" .OR. ;
	  cFile == "Gap in days"                .OR. ;
     cFile == "Max days on line"           .OR. ;
	  cFile == "No of days"                 .OR. ;
	  cFile == 'Max days at stage'
        IF lCode
           nDays := 0
        ELSE
           nDays := cBuffer
        ENDIF
CASE cFile == 'No. of weeks for profile'
        IF lCode
           nWeeks := 0
        ELSE
           nWeeks := cBuffer
        ENDIF
	CASE cFile == "Allocation-shipment gap"
        IF lCode
           nGap2 := 0
        ELSE
           nGap2 := cBuffer
        ENDIF
   CASE cFile == "Due-shipment gap" .OR. cFile == "Shipment-due gap"
        IF lCode
           nFromGap3 := 0
           nToGap3   := 0
        ELSE
           nFromGap3 := cBuffer[1]
           nToGap3   := cBuffer[2]
        ENDIF
   CASE cFile == "Ref# from...to"
        IF lCode
           cRefFrom := 0
           cRefTo   := 0
        ELSE
           cRefFrom := cBuffer[1]
           cRefTo   := cBuffer[2]
        ENDIF
   CASE (cFile == "c_bstat" .OR.  cFile == "Status") .AND. cBuff <> "Previous Status".AND. cBuff <> "Next Status"
        IF lCode
           cBstat := ""
        ELSE
           cBstat := cBuffer
        ENDIF
   CASE (cFile == "c_bstat" .OR.  cFile == "Status") .AND. cBuff == "Previous Status"
        IF lCode
           cPBstat := ""
        ELSE
           cPBstat := cBuffer
        ENDIF
   CASE (cFile == "c_bstat" .OR.  cFile == "Status") .AND. cBuff == "Next Status"
        IF lCode
           cNBstat := ""
        ELSE
           cNBstat := cBuffer
        ENDIF
   CASE cFile == "c_bpurp" .OR.  cFile == "Purpose"
        IF lCode
           cPurpose := ""
        ELSE
           cPurpose := cBuffer
        ENDIF
   CASE cFile == "c_prior" .OR.  cFile == "Priority"
        IF lCode
           cPriority := ""
        ELSE
           cPriority := cBuffer
        ENDIF
   CASE cFile == "d_ordrec" .or. cFile == "Booked from...to"
        IF lCode
           dBookedFrom := ""
           dBookedTo   := ""
        ELSE
           dBookedFrom := cBuffer[1]
           dBookedTo   := cBuffer[2]
        ENDIF
   CASE cFile == "WH Entry from...to"
        IF lCode
           dEFrom := ""
           dETo   := ""
        ELSE
           dEFrom := cBuffer[1]
           dETo   := cBuffer[2]
        ENDIF
   CASE cFile == "d_ordrec" .or. cFile == "Non-non date"  // Vitaly 10-06-02 ID 2610845
        IF lCode
           dNonNon := ""
        ELSE
           dNonNon := cBuffer[1]
        ENDIF
    CASE cFile == "dadd_rec" .or. cFile == "Shipped from...to"
        IF lCode
           dShippedFrom := ""
           dShippedTo   := ""
        ELSE
           dShippedFrom := cBuffer[1]
           dShippedTo   := cBuffer[2]
        ENDIF
   CASE cFile == "d_wid" .or. cFile == "Worker ID"
        IF lCode
           cWorkerId := ""
        ELSE
           cWorkerId := cBuffer
        ENDIF
   CASE cFile == "c_rejcd" .or. cFile == "Rejection code"
        IF lCode
           cRejectId := ""
        ELSE
           cRejectId := cBuffer
        ENDIF
   CASE cFile == "Remark"
        IF lCode
           cRemark := ""
        ELSE
           cRemark := cBuffer
        ENDIF
   CASE cFile == "Dialecter"
        IF lCode
           cDialecter := ""
        ELSE
           cDialecter := cBuffer
        ENDIF
   CASE cFile == "Batch number"
        IF lCode
           cBid := ""
        ELSE
           cBid := cBuffer
        ENDIF

   CASE cFile == "Order number"
        IF lCode
           nPolnId := 0
        ELSE
           nPolnId := cBuffer
        ENDIF
   CASE cFile == "Initiator" .OR. cFile == "c_tinit"
        IF lCode
           cBloke := space(10)
        ELSE
           cBloke := cBuffer
        ENDIF
   CASE cFile == "Route filter"
        IF lCode
           cRoute := ""
        ELSE
           cRoute := cBuffer
        ENDIF
   CASE cFile == "Stock valuation file"  .OR. cFile == "d_svfls"
        IF lCode
           cSValFile := ""
        ELSE
           cSValFile := cBuffer
        ENDIF
   CASE cFile == "Stock movement valuations"  .OR. cFile == "d_svmfls"
        IF lCode
           cSValmFile := ""
        ELSE
           cSValmFile := cBuffer
        ENDIF
   CASE cFile == "To date"
        IF lCode
           dTo := ""
           tTo := ""
        ELSE
           dTo := cBuffer[1]
           tTo := cBuffer[2]
        ENDIF
   CASE cFile == "Min&Max Hours at Rep.Stage"  // VR 27-01-01 TID:2127212
        IF lCode
           nHoursFrom := 0
           nHoursTo   := 0
        ELSE
           nHoursFrom := cBuffer[1]
           nHoursTo   := cBuffer[2]
        ENDIF
   CASE cFile == "Engineer"
        IF lCode
           cEngineer := ""
        ELSE
           cEngineer := cBuffer
        ENDIF
   CASE cFile == "Promised Batches"
        IF lCode
           lPromised := ""
        ELSE
           lPromised := cBuffer
        ENDIF
   CASE cFile == "NV Pcs to transfer"
        IF lCode
           lProm_ANV := ""
        ELSE
           lProm_ANV := cBuffer
        ENDIF
   CASE cFile == "Actual transfer"
        IF lCode
           lProm_Trans := ""
        ELSE
           lProm_Trans := cBuffer
        ENDIF
   CASE cFile == "Difference in the ESN's"
        IF lCode
           lProm_Sameesn := ""
        ELSE
           lProm_Sameesn := cBuffer
        ENDIF
   CASE cFile == "Pack instruction's"
        IF lCode
           lProm_Pack := ""
        ELSE
           lProm_Pack := cBuffer
        ENDIF
   CASE cFile == "Unexplained Gap"
        IF lCode
           lM999 := ""
        ELSE
           lM999 := cBuffer
        ENDIF
   CASE cFile == "Slow Moving"
        IF lCode
           lPromised := ""
        ELSE
           lPromised := cBuffer
        ENDIF
   CASE cFile == "Reporting Stage" //VR 27-01-01 TID:2127212
        IF lCode
           cStage := ""
        ELSE
           cStage := cBuffer
        ENDIF
	CASE cFile == "Ship As" .OR. cFile == "c_shipas"  //VR 13-11-01 111132334
        IF lCode
           cShipAs := ""
        ELSE
           cShipAs := cBuffer
        ENDIF
   OTHERWISE
        Tone( 400,2) ; Tone( 500 , 2 ) ; Tone( 200 , 2 )
        Alert( "Criterion title not found in Set Buffer crit list ("+cFile+")",{"CALL MIS NOW!!"})
ENDCASE
RETURN

STATIC PROCEDURE OpenIndex( cInd , cFile )
LOCAL aIndx
LOCAL oUser := GetuserInfo()

DO CASE
   CASE cFile == "c_ptype"
		  IIF(file(oUser:cTempDir+"iptype_1"+".cdx"),FERASE(oUser:cTempDir+"iptype_1"+".cdx") ,NIL )
		  CheckIndex("iptype_1",oUser:cTempDir + "iptype_1","ptype_id",.T., ,.T.)
   CASE cFile == "c_pline"
        IIF(file(oUser:cTempDir+"ipline_1"+".cdx"),FERASE(oUser:cTempDir+"ipline_1"+".cdx") ,NIL )
        CheckIndex("ipline_1",oUser:cTempDir + "ipline_1","pline_id",.T., ,.T.)
        //c_pline->(ordsetfocus("c_pline"))
        IF Empty(cProductType)

        ELSE
             c_pline->(AOFSetFilter("c_pline->ptype_id $ GetBuffer('c_ptype')"))
        ENDIF
   CASE cFile == "c_warea"
        IIF(file(oUser:cTempDir+"iparea_1"+".cdx"),FERASE(oUser:cTempDir+"iparea_1"+".cdx") ,NIL )
        CheckIndex("iparea_1",oUser:cTempDir + "iparea_1","warea_id",.T., ,.T.)
   CASE cFile == "c_size"
         c_size->(ordsetfocus("c_size"))
   CASE cFile == "c_value"
        c_value->(ordsetfocus("ivalpr"))
        IF Empty(cProductType)

        ELSE
             c_value->(AOFSetFilter("c_value->ptype_id $ GetBuffer('c_ptype')"))
        ENDIF
   CASE cFile == "c_tol"
        c_tol->(ordsetfocus("c_tol"))
        IF Empty(cProductType)

        ELSE
           c_tol->(AOFSetFilter("c_tol->ptype_id $ GetBuffer('c_ptype')"))
        ENDIF
   CASE cFile == "c_tc"
        c_tc->(ordsetfocus("c_tc"))
        IF Empty(cProductType)

        ELSE
             c_tc->(AOFSetFilter("c_tc->ptype_id $ GetBuffer('c_ptype')"))
        ENDIF
   CASE cFile == "c_esnxx"
          c_esnxx->(ordsetfocus("esnxxid"))
   CASE cFile == "c_esny"

   CASE cFile == "c_btype"
        IIF(file(oUser:cTempDir+"ibtype_1"+".cdx"),FERASE(oUser:cTempDir+"ibtype_1"+".cdx") ,NIL )
        CheckIndex("ibtype_1",oUser:cTempDir + "ibtype_1","b_type",.T., ,.T.)
        IF Empty(cProductType)

        ELSE
             c_btype->(AOFSetFilter("c_btype->ptype_id $ GetBuffer('c_ptype')"))
        ENDIF
   CASE cFile == "d_wid"
        d_wid->(ordsetfocus("iworkid"))
   OTHERWISE
        IF ! Empty(cInd)
          (cFile)->(ordsetfocus(cInd))
        ELSE
        aIndx := GetFileIndexInfo( cFile )
          IF LEN(aIndx)>0
              (cFile)->(ordsetfocus(aIndx[1][3]))
          ENDIF
        ENDIF

ENDCASE
(cFile)->(DBGOTOP())
RETURN
******************************************************
FUNCTION BuildCritData( aCut , aCutIndicators , lFlag)
LOCAL cTemp := "" , Buffer  , cType
LOCAL nLen , i

DEFAULT lFlag TO .T. // .T. prepate for file to save the criterions
DEFAULT aCutIndicators TO {}

nLen := Len( aCutIndicators )


FOR i := 1 TO nLen
    IF aCutIndicator[i]

       Buffer := GetBuffer( aCut[i] )
       cType := ValType( Buffer )

       IF cType == "A"
          Buffer := Turn2Str( Buffer , lFlag )
       ELSEIF cType == "N"
          Buffer := Str( Buffer )
       ELSEIF cType == "D"
          Buffer := Dtoc( Buffer )
       ENDIF

       IF lFlag
          cTemp += aCut[i]+"=" + "Type-"+cType+"."+ IIF(RIGHT(Buffer,1) == ".",Buffer,Buffer + ".") + Chr(10) + Chr(13)
       ELSE
          cTemp += Padr(aCut[i],20)+":" + Buffer + Chr(10) + Chr(13)
       ENDIF

    ELSE
       IF !lFlag
          cTemp += Padr(aCut[i],20)+": all "  + Chr(10) + Chr(13)
       ENDIF
    ENDIF
NEXT

RETURN cTemp
*******************************************************
FUNCTION aBuildCritData( aCut , aCutIndicators , lFlag)
LOCAL aTemp := {} , Buffer  , cType
LOCAL nLen := Len( aCutIndicators ) , i

DEFAULT lFlag TO .T. // .T. prepate for file to save the criterions

FOR i := 1 TO nLen
   IF aCutIndicator[i]

      Buffer := GetBuffer( aCut[i] )
      cType := ValType( Buffer )

      IF cType == "A"
         Buffer := Turn2Str( Buffer , lFlag )
      ELSEIF cType == "N"
         Buffer := Str( Buffer )
      ELSEIF cType == "D"
         Buffer := Dtoc( Buffer )
      ENDIF

      IF lFlag
         AADD(aTemp,aCut[i]+"=" + "Type-"+cType+"."+Buffer+"." )
      ELSE
         AADD(aTemp,Padr(aCut[i],20)+":" + Buffer)
      ENDIF

   ELSE
      IF !lFlag
         AADD(aTemp,Padr(aCut[i],20)+": all " )
      ENDIF
   ENDIF
NEXT

RETURN aTemp
*****************************************
STATIC FUNCTION Turn2Str( Buffer ,lFlag )
LOCAL nLen := Len( Buffer) , i
LOCAL cType
LOCAL cTemp := ""

DEFAULT lFlag TO .T.

FOR i:=1 TO nLen

    cType := ValType( Buffer[i] )

    IF lFlag
       cTemp += "["+cType+"]"
    ENDIF

    DO CASE
       CASE cType == "C"
            cTemp +=  Buffer[i]
       CASE cType == "N"
            cTemp += Str(Buffer[i])
       CASE cType == "D"
            cTemp += Dtoc( Buffer[i])
    ENDCASE

    IF i < nLen
       IF lFlag
          cTemp += ","   // make the buffer list
       ELSE
          cTemp += " - "
       ENDIF
    ENDIF
NEXT
RETURN cTemp

PROCEDURE prnSetBuffer( xVal )
cBuffer := xVal
RETURN

**********************
FUNCTION GetWkStn( o )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cWorkStation := NIL
   RETURN .T.
ENDIF

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter work station:"
cWorkStation := "    "
@ 5, 6  GET cWorkStation PICTURE "9999" ;
          SEND PreBlock  :=  {|o| PrePick( o, "c_wkstn")}          ;
          SEND PostBlock :=  {|o| PostPick(o,"c_wkstn","wkstn_id")};
          COLOR GETCOLORS
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

************************************
FUNCTION Get42WkStn( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cWorkStation := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cWrkStnFrom TO aParamList[1] , cWrkstnTo TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter work station:"
@ 4, 6  SAY "From:     To:"
@ 5, 6   GET cWrkStnFrom PICTURE "9999";
         SEND PreBlock  :=  {|o| PrePick( o, "c_wkstn")}          ;
         SEND postBlock := {|o| IF(o:VarGet()=="0000",ItsOk(), PostPick(o,"c_wkstn","wkstn_id")) };
         COLOR GETCOLORS
@ 5, 16  GET cWrkStnTo PICTURE "9999";
         SEND PreBlock  :=  {|o| PrePick( o, "c_wkstn")}          ;
         SEND postBlock := {|o| IF(o:VarGet()=="9999",ItsOk(), PostPick(o,"c_wkstn","wkstn_id")) };
         COLOR GETCOLORS
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
/*
STATIC FUNCTION postWork(o)
LOCAL lRetVal
LOCAL oTab
IF Empty( o:varGet() )
   RETURN .T.
ENDIF
oTab := TableTranslate():new( "c_wkstn" , {|| c_wkstn->wkstn_id == o:varGet() } )
IF oTab:xopen() .AND. oTab:search()
   lRetVal := .T.
ELSE
   Msg24( {"208"} , 3 , .T. )
   lRetVal := .F.
ENDIF
oTab:close()
RETURN lRetVal
*/
***********************
FUNCTION GetProcId( o )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cProcID := NIL
   RETURN .T.
ENDIF


scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter process:"
cProcId := "   "
@ 5, 6  GET cProcId PICTURE "999";
         SEND PreBlock  :=  {|o| PrePick( o, "c_proc")}          ;
         SEND postBlock := {|o|  PostPick(o,"c_proc","proc_id",,"iproc_id") };
         COLOR GETCOLORS
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
*************************************
FUNCTION Get42ProcId( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cProcID := NIL
   RETURN .T.
ENDIF


IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cProcIDFrom TO aParamList[1] , cProcIDTo TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter process :"
@ 4, 6  SAY "From:     To:"
@ 5, 6   GET cProcIdFrom PICTURE "999";
         SEND PreBlock  :=  {|o| PrePick( o, "c_proc")}          ;
         SEND postBlock := {|o| IF(o:VarGet()=="000",ItsOk(), PostPick(o,"c_proc","proc_id",,"iproc_id")) };
         COLOR GETCOLORS
@ 5, 16  GET cProcIdto PICTURE "999";
         SEND PreBlock  :=  {|o| PrePick( o, "c_proc")}          ;
         SEND postBlock := {|o| IF(o:VarGet()=="999",ItsOk(), PostPick(o,"c_proc","proc_id",,"iproc_id")) };
         COLOR GETCOLORS
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
*********************************
FUNCTION getpccode(o, aParamList)
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cProcID := NIL
   RETURN .T.
ENDIF


IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cPcCodeFrom TO aParamList[1] , cPcCodeTo TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter Pc_code :"
@ 4, 6  SAY "From:     To:"
@ 5, 6   GET cPcCodeFrom PICTURE "9999";
         COLOR GETCOLORS
@ 5, 16  GET cPcCodeTo PICTURE "9999";
         COLOR GETCOLORS
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*********************************
FUNCTION getpccode2(o, aParamList)
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cProcID := NIL
   RETURN .T.
ENDIF


IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cPcCodeFrom TO aParamList[1] , cPcCodeTo TO aParamList[2]

scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter Pc_code :"
@ 4, 6  SAY "From:     To:"
@ 5, 6   GET cPcCodeFrom PICTURE "9999";
         SEND PreBlock  := {|o| PrePick( o,"c_proc",{1,2,3,4,5},NIL,NIL,NIL,NIL,NIL,NIL,5)};
         SEND postBlock := {|o| PostPick(o,"c_proc","pcprocno")};
         COLOR GETCOLORS
@ 5, 16  GET cPcCodeTo PICTURE "9999";
         SEND PreBlock  := {|o| PrePick( o,"c_proc",{1,2,3,4,5},NIL,NIL,NIL,NIL,NIL,NIL,5)};
         SEND postBlock := {|o| PostPick(o,"c_proc","pcprocno")};
         COLOR GETCOLORS
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.

*************************************
FUNCTION GetPcAndPrior( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor
local bblock := {|| o:varPut( C_PROC->(FieldGet(6)) ),o:Changed := .T. }

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cProcID := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cPcCodeFrom TO aParamList[1] , cPcCodeTo TO aParamList[2]
DEFAULT cPriorFrom TO aParamList[3] , cPriorTo TO aParamList[4]

scrnPush()
DispBox( 1, 1 , 8 , 34 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter pccode and priority:"
@ 4, 6  SAY "From:  priority:"
@ 5, 6   GET cPcCodeFrom PICTURE "9999";
         SEND PreBlock  := {|o| PrePick( o,"c_proc",{1,2,3,4,5},NIL,NIL,NIL,NIL,NIL,NIL,5)};
         SEND postBlock := {|o| PostPick(o,"c_proc","pcprocno")};
         COLOR GETCOLORS
@ 5,14  GET cPriorFrom PICTURE "@!";
         send postblock := {|o| o:varGet() $ "HO_LT_ZZ" };
         COLOR GETCOLORS
@ 6, 6  SAY "To:    priority:"
@ 7, 6  GET cPcCodeto PICTURE "9999";
         SEND PreBlock  := {|o| PrePick( o,"c_proc",{1,2,3,4,5},,,,,,,5)};
         SEND postBlock := {|o| PostPick(o,"c_proc","pcprocno")};
         COLOR GETCOLORS
@ 7, 14 GET cPriorTo PICTURE "@!";
         send postblock := {|o| o:varGet() $ "HO_LT_ZZ" };
         COLOR GETCOLORS
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
*********************************

FUNCTION GetPcAndDate( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor,bblock := {|| o:varPut( C_PROC->(FieldGet(6)) ),o:Changed := .T. }

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   cProcID := NIL
   RETURN .T.
ENDIF

IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF

DEFAULT cPcCodeFrom TO aParamList[1] , cPcCodeTo TO aParamList[2]
DEFAULT dFinFrom TO aParamList[3] , dFinTo TO aParamList[4]

scrnPush()
DispBox( 1, 1 , 8 , 34 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter pccode and date:"
@ 4, 6  SAY "From:     date:"
@ 5, 6   GET cPcCodeFrom PICTURE "9999";
         SEND PreBlock  := {|o| PrePick( o,"c_proc",{1,2,3,4,5},NIL,NIL,NIL,NIL,NIL,NIL,5)};
         SEND postBlock := {|o| PostPick(o,"c_proc","pcprocno")};
         COLOR GETCOLORS
@ 5,14  GET dFinFrom Valid Notempty(dFinFrom)
@ 6, 6  SAY "To:       date:"
@ 7, 6  GET cPcCodeto PICTURE "9999";
         SEND PreBlock  := {|o| PrePick( o,"c_proc",{1,2,3,4,5},,,,,,,5)};
         SEND postBlock := {|o| PostPick(o,"c_proc","pcprocno")};
         COLOR GETCOLORS
@ 7, 14 GET dFinTo Valid Notempty(dFinTo)
XREAD

IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF

scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
*********************************
/*
STATIC FUNCTION postProcId(o)
LOCAL lRetVal
LOCAL oTab
IF Empty( o:varGet() )
   RETURN .T.
ENDIF
oTab := TableTranslate():new("c_proc",{||o:varGet()},{"iproc_id"})
IF oTab:xopen() .AND. oTab:search()
   lRetVal := .T.
ELSE
   Msg24( {"166","Process"} , 3 , .T. )
   lRetVal := .F.
ENDIF
oTab:close()
RETURN lRetVal
*/
/////////////////FROM VITALY Romanovsky/////////////////////////////////
*************************************
FUNCTION GetTop( o, aParamList )
LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor
IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF
IF !Eval(o:checkGsb)
   cTop := NIL
   RETURN .T.
ENDIF
IF ValType( aParamList ) != "A"
   Alert( "No params in criterion", {" OK "} )
   RETURN .F.
ENDIF
DEFAULT cTop TO aParamList[1]
scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter Top number:"
@ 5, 6  GET cTop PICTURE "99"  Valid Notempty(cTop)
XREAD
IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF
scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
****************************
//YG , Shalom 9/97
FUNCTION GetFinDateTime(o,aParamList)

LOCAL aGetList
LOCAL nRow := Row()
LOCAL nCol := Col()
LOCAL nCursor
LOCAL nLen, i

// 1st level error trapping
DEFAULT aParamList TO {NIL,NIL,NIL,NIL}
// 2nd level
IF ( nLen := LEN(aParamList) ) < 4
   FOR i := nLen+1 TO  4
      AADD( aParamList, NIL )
   NEXT
ENDIF

IF o:ExitState != GE_NOEXIT
   RETURN .T.
ENDIF

IF !Eval(o:checkGsb)
   dFrom:=tfrom:=dTo:=tTo:=NIL
   RETURN .T.
ENDIF
dFrom := IF( aParamList[1] == NIL, date() - 1, aParamList[1] )
dTo   := IF( aParamList[2] == NIL, date()    , aParamList[2])
tFrom := IF( aParamList[3] == NIL, "07:00"   , aParamList[3] )
tTo   := IF( aParamList[4] == NIL, "06:59"   , aParamlist[4] )
scrnPush()
DispBox( 1, 1 , 7 , 29 , B_SINGLE+" " )
nCursor := SetCursor( SC_NORMAL )
aGetList := GetList
GetList := {}
@ 3, 6  SAY "Enter date & time:"
@ 4, 6  SAY "From      To:"
//dFrom:= dTo := date() - 1
//tFrom := "00:00"
//tTo:= "23:59"
@ 5, 6  GET dFrom
@ 6, 6  GET tFrom PICTURE "99:99" SEND postBlock := {|o| postTime(o) }
@ 5,16  GET dTo   // SEND postBlock := {|o| postFinDate(o,dfrom) }
@ 6,16  GET tTo PICTURE "99:99" SEND postBlock := {|o| postTime(o) }
XREAD
IF LastKey() = K_ESC
   SETKEY( K_F2, NIL )
   Eval(o:checkGsb , .F. )
ELSE
SetCursor( nCursor )
ENDIF
scrnPop()
GetList := aGetList
SetPos( nRow , nCol )
RETURN .T.
*************************************
Static Function PostFindate(o,ddFrom)
LOCAL lRetVal
LOCAL nOrder
GenOpenFiles({"M_linemv"})
nOrder := m_linemv->(indexord())
m_linemv->( ORDSETFOCUS("ilnmvfin"))

IF m_linemv->(DBSEEK(dtos(ddFrom),.t.)) .and. m_linemv->cp_dfin<= o:VarGet()
    lRetVal := .T.
ELSE
   Msg24( "No Process on this date" , 3 , .T. )
   lRetVal := .F.
ENDIF

GenCloseFiles({"M_linemv"})
RETURN lRetVal
****************************
STATIC FUNCTION postTime(o)
LOCAL lRetVal
LOCAL nHH := Val(Left(o:varGet(),2))
LOCAL nMM := Val(Right(o:varGet(),2))

IF (nHH >= 0 .AND. nHH <= 24) .AND. (nMM >= 0 .AND. nMM <= 59)
   lRetVal := .T.
ELSE
   Msg24( {"210"} , 3 , .T. )
   lRetVal := .F.
ENDIF

RETURN lRetVal
******************************
FUNCTION InStr( cSearch, cSource )
RETURN (cSearch $ cSource)
****************************
FUNCTION Equal( u1, u2 )
RETURN ( u1 == u2 )
****************************
FUNCTION Great( u1, u2 )
RETURN ( u1 > u2 )
****************************
FUNCTION GrEq( u1, u2 )
RETURN ( u1 >= u2 )
****************************
FUNCTION Less( u1, u2 )
RETURN ( u1 < u2 )
****************************
FUNCTION LessEq( u1, u2 )
RETURN ( u1 <= u2 )
****************************
FUNCTION GrLeEq( u1, u2, u3 )
RETURN ( u1 >= u2 .AND. u1 <= u3)
****************************
// this function is a rewrite of the one in PrnCrit.prg
// added Max days on Line 01/01/98 SS for RPQC15V1
FUNCTION GetBuffer( cCode )

LOCAL aSource := ;
   {"c_ptype"           ,"c_pline"      ,"c_size",;
    "c_value"           ,"c_yesno","c_tol"        ,"c_volt",;
    "c_tc","omaesn"             ,"c_repunt", "c_cust",    "c_term"    ,"c_esnxx",;
    "c_esny"            ,"c_btype"      ,"c_bstat","Previous Status","Next Status","c_prornd",;
    "c_bpurp"           ,"c_prior"      ,"B/N from...to",;
    "Date from...to"    ,"Date into Warehouse","Scheduled finish date","No of days at current proc",;
    "Time from...to"    ,"c_stkmv"      ,"Ser# from...to",;
    "Ref# from...to"    ,"Date time finish","Specific Process",;
    "Specific workstation","Process"    , "Workstation",;
    "c_potype"          ,"Booked from...to","WH Entry from...to","c_sarea","c_warea","Non-non date" ,;
    "c_Cntry"           ,"c_agency"     ,"c_cuco" ,;
    "c_cust"            ,"P/O Status"   ,"d_ordrec",;
    "dadd_rec"          ,"c_wkstn"      ,"c_proc",;
    "DUE Date from...to","ESN from...to","Minimum sold",;
    "Sales period"      ,"Location","Warehouse","Source WH","Destination WH","Gap in days",;
    "Pc_code"           ,"Billed from...to","Close from...to",;
    "Allocation from...to","Max days on Line","No of days",;
    "Entry-allocation gap","Allocation-shipment gap","Revised due date from...to",;
    "Cancel date from...to","Invoice from...to","Worker ID",;
    "Rejection code"    ,"Due-shipment gap","Entry-alloc-ship gaps",;
    "Shipment-due gap"  ,"Arrive/finish dates","Months booked","No. of weeks for profile",;
    "Pc_code and priority","Remark"     ,"Dialecter",;
    "Months"            ,"Marketing due date from...to","Batch number","Top number",; //VR
    "Order number"      ,"Finish dates" ,"Process No.", "Process name",;
    "Request date from...to","c_tinit"  ,"Quantity greater than","Scrap cost greater than","Var from Std. greater than",;
    "Max days at stage" ,"Route filter" ,"Date into Warehouse",;
    "Order date from...to","Slack from...to","Min&Max Hours at Rep.Stage","Start date from...to",;
    "Pc_code and date"  ,"Process type"   ,"Caps","Stock valuation file",;
    "d_svfls"           ,"Stock movement valuations","d_svmfls",;
    "To date"           ,"c_svalmv", "Engineer","Sono from...to",;
    "Soordln from...to","Reporting Stage","NV Pcs to transfer","Difference in the ESN's","Pack instruction's","Actual transfer","Promised Batches","Slow Moving","Unexplained Gap","Ship As","c_doctyp" } //VR ID:111132334
LOCAL aTitle := ;
   {"Product type"      ,"Product line" ,"Size",;
    "Value"             ,"Split batch","Tolerance"    ,"Voltage",;
    "TC"                ,"OMARIM Parts","Rep.Unit","End Customer", "Termination"  ,"ESN(XX)",;
    "ESN(Y)"            ,"Batch type"   ,"Status","Previous Status","Next Status","Project",;
    "Purpose"           ,"Priority"     ,NIL,;
     NIL, NIL, NIL, NIL,;
     NIL                ,"Movement type", NIL,;
     NIL, NIL, NIL,;
     NIL, NIL, NIL,;
     "P/O Type"         ,NIL            ,"WH Entry from...to","Area","Production Area","Non-non date",;
     "Country"          ,"Billing Agency","Company",;
     "Customer"         ,"P/O Status"   ,"Booked from...to",;
     "Shipped from...to","Workstations" ,"Processes",;
     "DUE Date from...to","ESN from...to","Minimum sold",;
     "Sales period"     ,"Location","Warehouse","Source WH","Destination WH","Gap in days",;
     "Pc_code"          ,"Billed from...to","Close from...to",;
     "Allocation from...to","Max days on line", "No of days",;
     "Entry-allocation gap","Allocation-shipment gap","Revised due date from...to",;
     "Cancel date from...to","Invoice from...to","Worker ID",;
     "Rejection code"   ,"Due-shipment gap","Entry-alloc-ship gaps",;
     "Shipment-due gap" ,"Arrive/finish dates","Months booked","No. of weeks for profile",;
     "Pc_code and priority","Remark"    ,"Dialecter",;
     "Months"           ,"Marketing due date from...to","Batch number","Top number",;//VR
     "Order number"     ,"Finish dates" ,"Process No.", "Process name",;
     "Request date from...to","Initiator","Quantity greater than","Scrap cost greater than","Var from Std. greater than",;
     "Max days at stage" ,"Route filter","Date into Warehouse",;
     "Order date from...to","Slack from...to","Min&Max Hours at Rep.Stage","Start date from...to",;
     "Pc_code and date"    ,"Process type"   ,"Caps","Stock valuation file" ,;
     NIL                  ,"Stock movement valuations",NIL,;
     "To date"            ,"Move type", "Engineer","Sono from...to","Soord from...to","Reporting Stage","NV Pcs to transfer","Difference in the ESN's","Pack instruction's","Actual transfer","Promised Batches","Slow Moving","Unexplained Gap","Ship As","From-To"}
LOCAL aRet := ;
    { cProductType       ,cProductLine  ,cSize,;
      cValue             ,cYesNo,cTol          ,cVoltage,;
      cTC                ,cOma,cRepUnit,cEndCustomer,  cTermination  ,cESNXX,;
      cESNY              ,cBatchType    ,cBstat,cPBstat,cNBstat,cProj,;
      cPurpose           ,cPriority     ,{cBNFrom , cBNTo},;
      {dFrom , dTo }     ,{dFrom , dTo },dFrom, nDays,;
      {tFrom , tTo }     ,cMoveType     ,{ nSerFrom , nSerTo },;
      {cRefFrom,cRefTo}  ,{ dFrom , tFrom , dTo , tTo },cProcId,;
      cWorkStation       ,{cProcIdFrom,cProcIdTo},{cWrkStnFrom,cWrkStnTo},;
      cPOType            ,{dBookedFrom,dBookedTo},{dEFrom,dETo},cArea,cPArea,dNonNon,;
      cCountry           ,cAgency       ,cCompany,;
      cCustomer          ,cPoStat       ,{dBookedFrom,dBookedTo},;
      {dShippedFrom,dShippedTo} ,cWorkStation,cProcId,;
      {dFrom , dTo }     ,{cEsnFrom,cEsnTo},nMinSold,;
      {dFrom,dTo}        ,cLoc,cWarehouse,cWareh_sr,cWareh_ds,nDays,;
      {cPcCodeFrom,cPcCodeTo},{dFrom ,dTo},{dFrom ,dTo},;
      {dAllocFrom ,dAllocTo},nDays      ,nDays,;
      nGap1              ,nGap2         ,{dRevFrom , dRevTo },;
      {dFrom , dTo }     ,{cDOCfrom,cDOCTo},cWorkerId,;
      cRejectId          ,{nFromGap3,nToGap3}, {nGap1,nGap2},;
      {nFromGap3,nToGap3},{dFrom,dTo}   ,nMonths, nWeeks,;
      {cPcCodeFrom,cPcCodeTo,cPriorFrom,cPriorTo},cRemark,cDialecter,;
      nMonths            ,{dMrkDueFrom,dMrkDueTo },cBid,cTop,;
      nPolnid            ,{dFrom,dTo}   ,{cProcFrom,cProcTo},{cProcFrom,cProcTo},;
      {dReqFrom,dReqTo}  ,cBloke        ,nQty,nQty_A,nQty_B,;
      nDays              ,cRoute        ,{dFrom,dTo},;
      {dFrom,dTo}        ,{nFrom,nTo},{nHoursFrom,nHoursTo}   ,{dStartFrom,dStartTo},;
      {cPcCodeFrom,cPcCodeTo,dFinFrom,dFinTo},cProcType,cBncaps,cSValFile,;
      cSValFile          ,cSValmFile    ,cSValmFile,;
      {dTo,tTo}          ,cSVMoveType, cEngineer,{cSonofrom,cSonoTo},{cSoordlnfrom,cSoordlnTo},cStage,lProm_ANV,lProm_Sameesn,lProm_Pack,lProm_Trans,lPromised,lPromised,lM999,cShipAs,;
		cFrom_To} //VR ID 111132334

LOCAL n, uRet := ""

DO CASE
   CASE (n := ASCAN( aSource, cCode ) ) > 0
        uRet := aRet[n]
   CASE (n := ASCAN( aTitle, cCode ) ) > 0
        uRet := aRet[n]
   OTHERWISE
        Tone( 400,2) ; Tone( 500 , 2 ) ; Tone( 200 , 2 )
        Alert( "Criterion title not found in GetBuffer crit list ("+cCode+")",{"CALL MIS NOW!!"})
END

RETURN uRet

//This is a generic replacement for the STATIC FUNCTION Cond found in all
//of the report .prgs

FUNCTION ShaiCond( aTestBlocks,oScrl,cField,lReport )
// S.b The function that running on the CritArrays and return .t. or. .f.
// for this record.
LOCAL aIndic
LOCAL nLen
LOCAL nTrues := 0
LOCAL lRetVal
LOCAL cColor
LOCAL i,nKey

default lReport to TRUE
////e-mail query
IF GetRepFileName() == "QMRK31VE"
	aIndic := array(18)
	afill(aIndic,.F.)
	aIndic[12] := aIndic[13] := aIndic[14] := aIndic[18] := TRUE
ELSEIF GetRepFileName() == "QMRK33VE"
	aIndic := array(20)
	afill(aIndic,.F.)
	aIndic[14] := aIndic[15] := aIndic[19] := TRUE
ELSEIF GetRepFileName() == "QMRK30VE"
	aIndic := array(19)
	afill(aIndic,.F.)
	aIndic[12] := aIndic[13] := aIndic[19] := TRUE
ELSEIF GetRepFileName() == "QMRK32VE"
	aIndic := array(21)
	afill(aIndic,.F.)
	aIndic[13] := aIndic[18] := TRUE
ELSEIF GetRepFileName() == "QFGS32VE"
	aIndic := array(19)
	afill(aIndic,.F.)
	aIndic[12] := aIndic[10] := aIndic[14] := aIndic[17] := TRUE
ELSEIF GetRepFileName() == "QMRK39VE"
	aIndic := array(20)
	afill(aIndic,.F.)
	aIndic[2] := aIndic[12] := aIndic[13] := aIndic[14] := aIndic[18] := aIndic[19] := aIndic[20] := TRUE
ELSEIF GetRepFileName() == "RACC07VE"
	aIndic := array(15)
	afill(aIndic,.F.)
	aIndic[1] := TRUE
ELSEIF GetRepFileName() == "RPQC08VE"
	aIndic := array(15)
	afill(aIndic,.F.)
	aIndic[4] := TRUE
ELSEIF GetRepFileName() == "RPQC06VE"
	aIndic := array(15)
	afill(aIndic,.F.)
	aIndic[3]  := TRUE
	aIndic[4]  := TRUE
	aIndic[15] := TRUE
ELSEIF GetRepFileName() == "RACC07VU"
	aIndic := array(15)
	afill(aIndic,.F.)
	aIndic[1] := TRUE
	aIndic[2] := TRUE
ELSEIF GetRepFileName() == "RACC01VE"
	aIndic := array(11)
	afill(aIndic,.F.)
	aIndic[4] := aIndic[11] := TRUE
ELSEIF GetRepFileName() == "RPQC39VE"
	aIndic := array(5)
	afill(aIndic,.F.)
ELSEIF GetRepFileName() == "RPQC24VE"
	aIndic := array(17)
	afill(aIndic,.F.)
	aIndic[9] := aIndic[10] := aIndic[15] := TRUE
ELSEIF GetRepFileName() == "RPQC03VE" .OR. GetRepFileName() == "RPQC03VU"
	aIndic := array(20)
	afill(aIndic,.F.)
	aIndic[1] := aIndic[2] := aIndic[6] := aIndic[18] := TRUE
ELSE
	aIndic := GetIndicators()
ENDIF
nLen := LEN( aIndic )

FOR i := 1 TO nLen
   IF aIndic[i]
      IF EVAL(aTestBlocks[i])
         nTrues++
      ENDIF
   ELSE
      nTrues++
   ENDIF
NEXT

lRetVal := (Len( aIndic ) = nTrues)

IF lRetVal
   cColor := "W+/BG"
ELSE
   cColor := "G/BG"
END

IF oScrl <> NIL .AND. lReport
   oScrl:scrollUp( (cField) , cColor )
IF (lReport .AND. InKey() == K_ESC)
       nKey := Alert( "Do you really want to stop?" , {"No" , "Yes" }, ALERT_STD )
       IF nKey = 2
          (Alias())->( DbGoBottom() )
          Tone(300,1) ; Tone(900,2)
          lRetVal := FALSE
       ENDIF
ENDIF
ENDIF

RETURN lRetVal


FUNCTION IsThereCond()
LOCAL aIndic := GetIndicators()
LOCAL nLen
LOCAL lRetVal := FALSE
LOCAL i

IF aIndic <> NIL
   nLen   := LEN( aIndic )
   FOR i := 1 TO nLen
      IF aIndic[i]
            lRetVal := TRUE
      ENDIF
   NEXT
END


RETURN lRetVal

/*************************************************************
*
*  I need this to kiil the F2 SetKey in some GETs
*
**************/
FUNCTION ItsOk
   SetKey( K_F2 , NIL )
RETURN TRUE


function Searchit(cSearcher,cFile)
local lfound := .F.
local nBroRow := oBro:rowPos
local nRow, nCol, i, j , ctesting
local cSaveLine := SaveScreen( 2,0, 2,79 )

IF empty((cFile)->(ordname()))
     return cSearcher
ENDIF
IF empty(cSearcher)
     (cFile)->(dbgotop())
     oBro:refreshall()
     oBro:stabilize()
     return cSearcher
ENDIF
if (cFile)->(dbseek(cSearcher))
     oBro:refreshall()
     oBro:stabilize()
else
    Alert("!! �� ����� ����� ���",{"..."})
    cSearcher := SubStr(cSearcher,1,Len(cSearcher)-1)
    if (cFile)->(dbseek(cSearcher))
          oBro:refreshall()
          oBro:stabilize()
    else
          (cFile)->(dbgotop())
          oBro:refreshall()
          oBro:stabilize()
    endif
endif
nRow := row()
nCol := col()
RestScreen(2,0,2,79,cSaveLine)
@ 2,3 say space(13)
@ 2,3 say cSearcher
setpos(nRow,nCol)
return cSearcher