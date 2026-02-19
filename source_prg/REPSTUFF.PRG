/*
 Algorithm for RPQC03V1
 -----------------------------------------------------------
 Algorithm for RPQC04V1
 -----------------------------------------------------------
 Algorithm for RPQC06V1
 -----------------------------------------------------------
 Algorithm for RPQC07V1
 -----------------------------------------------------------
 Algorithm for RPQC20V1
 -----------------------------------------------------------
 Algorithm for RPQC13V1
*/

#include "avxdefs.ch"

// Constants for RPQC06V1
#define C06_TOW   4
#define C06_TOP   5
#define C06_OTP  11
#define C06_TARG 12
#define C06_GAP  13
#define C04_POTP   9
#define C04_PTOP  10
#define C04_COTP  16
#define C04_CTOP  17

#define C07_PERW 9
#define C07_POTW 10
#define C07_DAYS 11
#define C07_BATS 12
#define TTL_TOW   1
#define TTL_TOP   2

STATIC atpqc03
STATIC NCUMP_APP,NCUMP_AFP,NCUMP_ALP,NCUMP_AGP,NCUMP_CPP,NCUMP_OTP,NCUMP_TOP



/********************************************************************/


FUNCTION t_lineTtls( oRep )

LOCAL aTtlInfo[17]
LOCAL aTpqc04 := {;
                   {"PCPROCNO" ,"C", 4, 0 },;
                   {"PROC_ID"  ,"C", 5, 0 },;
                   {"PROC_NMH" ,"C",20, 0 },;
                   {"P_APP"    ,"N",12, 0 },;
                   {"P_AFP"    ,"N",12, 0 },;
                   {"P_ALP"    ,"N",12, 0 },;
                   {"P_AGP"    ,"N",12, 0 },;
                   {"P_CPP"    ,"N",12, 0 },;
                   {"P_OTP"    ,"N",12, 0 },;
                   {"P_TOP"    ,"N",12, 0 },;
                   {"CUMP_APP" ,"N",12, 0 },;
                   {"CUMP_AFP" ,"N",12, 0 },;
                   {"CUMP_ALP" ,"N",12, 0 },;
                   {"CUMP_AGP" ,"N",12, 0 },;
                   {"CUMP_CPP" ,"N",12, 0 },;
                   {"CUMP_OTP" ,"N",12, 0 },;
                   {"CUMP_TOP" ,"N",12, 0 } }
LOCAL aDbfs := {"T_pqc04","t_line","c_proc"  ,"c_thqty","c_tthqty","c_expqty","d_line"  }
LOCAL aNtxs := {""       , ""     ,"iproc_id","itsuom" ,"itsuom" ,"iuts"    ,"ib_idln"}
LOCAL aTtls
LOCAL i
LOCAL cTempDir := GetUserInfo():cTempDir

IF !FILE(cTempDir + "T_pqc04")
   DBCREATE( cTempDir+"T_pqc04",aTpqc04,   )
ENDIF
IF SELECT("T_pqc04") > 0; CLOSE T_pqc04 ; ENDIF
NetUse( "T_pqc04", STD_RETRY,  , USE_EXCLUSIVE, USE_NEW, cTempDir )
T_pqc04->( __DBZAP() )

NetUse( "T_line", STD_RETRY,  , USE_EXCLUSIVE, USE_NEW, cTempDir )

CheckIndex("t_line",cTempDir + "t_line","cpproc_id+pline_id",.T.)

FOR i := 3 TO LEN( aDbfs )
    NetUse( aDbfs[i]  , STD_RETRY, RDD_IN_USE, USE_SHARED, USE_NEW, NIL ) //vr
    (aDbfs[i])->( ordsetfocus(aNtxs[i]) )
	 (aDbfs[i])->( DBGOTOP() )
NEXT


WHILE !c_proc->( EOF() )
      ResetArray(aTtlInfo)
      IF t_line->( DBSEEK(c_proc->proc_id ) )
         WHILE t_line->cpproc_id == c_proc->proc_id
              aTtls := TtlQty("t_line")
              UpdateC04Array( aTtlInfo, aTtls, "t_line" )
              t_line->( DBSKIP() )
         END
         Tpqc04Update( aTtlInfo )
      ELSE
         Tpqc04Update( aTtlInfo )
      END
      c_proc->( DBSKIP() )
END

// now do the cumulative calculation of t_pqc04->cump_top

Tpqc04TtlUpdate()

FOR i := 1 TO LEN( aDbfs )
    (aDbfs[i])->( DBCLOSEAREA() )
NEXT

oRep:cPrepDbf := "T_pqc04"

RETURN NIL
*
FUNCTION C_procTtls( oRep )
LOCAL nDays   := 1
LOCAL aToFrom := oRep:aMyBuffer[6]  // YH 21-6-00
LOCAL aTtlInfo[13]
LOCAL aTpqc06 := {;
                   {"PCPROCNO" ,"C", 4, 0 },;
                   {"PROC_ID"  ,"C", 5, 0 },;
                   {"PROC_NMH" ,"C",20, 0 },;
                   {"P_TOW"    ,"N",12, 0 },;
                   {"P_TOP"    ,"N",12, 0 },;
                   {"P_APP"    ,"N",12, 0 },;
                   {"P_AFP"    ,"N",12, 0 },;
                   {"P_ALP"    ,"N",12, 0 },;
                   {"P_AGP"    ,"N",12, 0 },;
                   {"P_CPP"    ,"N",12, 0 },;
                   {"P_OTP"    ,"N",12, 0 },;
                   {"DAY_TARG" ,"N",12, 0 },;
                   {"GAP"      ,"C", 6, 0 };
                   }
LOCAL aTtls,aZmniTtls,nZefiPieces
LOCAL aDbfs    := {"T_pqc06","t_linemv","d_line" ,"c_proc"  ,"c_thqty","c_tthqty","c_expqty","d_line"  }
LOCAL aNtxs    := {""       , ""       ,"ib_idln","iproc_id","itsuom" ,"itsuom" ,"iuts"    ,"ib_idln"}
LOCAL i
LOCAL cTempDir := GetUserInfo():cTempDir
LOCAL cScr     := SAVESCREEN(2,0,2,79)
LOCAL nTrec    := 0
LOCAL nCrec    := 0,nSikum:=0
IF !FILE(cTempDir+"T_pqc06")
   DBCREATE( cTempDir+"T_pqc06",aTpqc06,   )
ENDIF
IF SELECT("T_pqc06") > 0; CLOSE T_pqc06 ; ENDIF
NetUse( "T_pqc06", STD_RETRY,  , USE_EXCLUSIVE, USE_NEW, cTempDir )
T_pqc06->( __DBZAP() )
NetUse( "t_linemv", STD_RETRY,  , USE_EXCLUSIVE, USE_NEW, cTempDir )
CheckIndex("t_linemv",cTempdir + "t_linemv","t_linemv->cpproc_id",.T.)
FOR i := 3 TO LEN( aDbfs ) -1
    IF Select(aDbfs[i])==0
        NetUse( aDbfs[i]  , STD_RETRY, RDD_IN_USE, USE_SHARED, USE_NEW, NIL )
    ENDIF
    (aDbfs[i])->( ordsetfocus(aNtxs[i]) )
	 (aDbfs[i])->( dbgotop() )
NEXT
IF aToFrom == NIL
   aToFrom := 1
ENDIF
nDays :=  aToFrom
@2, 40 SAY "Process:"      COLOR "w+/br"
WHILE !C_proc->( EOF() )
      @2,48 SAY C_proc->Proc_nmh COLOR "gr+/br"
      ResetArray(aTtlInfo)
      IF T_linemv->( DBSEEK(c_proc->proc_id) )
         nTrec := 0
         WHILE T_linemv->cpproc_id == c_proc->proc_id
               aZmniTtls:=GetMaxWSP(C_proc->UOM_id,,T_lineMv->Cp_totScr,"T_lineMv")
               nZefiPieces:=Exc06Qty("T_lineMv",T_lineMv->Cp_bQtyP-aZmniTtls[3],"D_line")
               aTtls:={T_lineMv->Cp_bQtyW,nZefiPieces} // Wafers š…Ž‹,Pieces ‰”–
               UpdateC06Array( aTtlInfo, aTtls, "t_linemv" )
               T_linemv->( DBSKIP() )
               @2,69 SAY ++nTrec  COLOR "gr+/br"
         END
         Tpqc06Update( aTtlInfo, nDays, c_proc->day_targ )
      ELSE
         T_pqc06->( AddRec(5) )  // „™ƒ‡ „Ž…™˜
         T_pqc06->PCPROCNO :=c_proc->pcprocNo
         T_pqc06->PROC_ID  :=c_proc->Proc_id
         T_pqc06->PROC_NMH := C_PROC->PROC_NME//IIF("CZ" $ GetUserInfo():cGroupID,C_PROC->PROC_NME,C_PROC->PROC_NMH)
         T_pqc06->DAY_TARG :=c_proc->DAY_TARG
         T_pqc06->( DBUNLOCK() )
      END
      C_proc->( DBSKIP() )
END
FOR i := 1 TO LEN( aDbfs )
    (aDbfs[i])->( DBCLOSEAREA() )
NEXT

oRep:cPrepDbf := "T_pqc06"
RESTSCREEN( 2,0,2,79,cScr)
RETURN NIL
*
FUNCTION WipWafer( oRep )

LOCAL aTtlInfo[12]
LOCAL aTpqc07 := {;
                   {"PCPROCNO" ,"C", 4, 0 },;
                   {"PROC_ID"  ,"C", 5, 0 },;
                   {"PROC_NMH" ,"C",20, 0 },;
                   {"P_APW"    ,"N", 8, 0 },;
                   {"P_AFW"    ,"N", 8, 0 },;
                   {"P_ALW"    ,"N", 8, 0 },;
                   {"P_AGW"    ,"N", 8, 0 },;
                   {"P_CPW"    ,"N", 8, 0 },;
                   {"P_ERW"    ,"N", 8, 0 },;
                   {"P_OTW"    ,"N", 8, 0 },;
                   {"P_TOW"    ,"N", 8, 0 },;
                   {"AVG_DAYS" ,"N", 8, 0 } }
LOCAL aDbfs := {"T_pqc07","t_line","c_proc" }
LOCAL aNtxs := {""       , ""     ,"iproc_id"}
LOCAL aTtls
LOCAL i
LOCAL cTempDir := GetUserInfo():cTempDir

IF !FILE(cTempDir + "T_pqc07")
   DBCREATE( cTempDir+"T_pqc07",aTpqc07,   )
ENDIF
IF SELECT("T_pqc07") > 0; CLOSE T_pqc07 ; ENDIF
NetUse( "T_pqc07", STD_RETRY,  , USE_EXCLUSIVE, USE_NEW, cTempDir )
T_pqc07->( __DBZAP() )

NetUse( "T_line", STD_RETRY,  , USE_EXCLUSIVE, USE_NEW, cTempDir )
CheckIndex("t_line","t_line","cpproc_id+pline_id",.T.)

FOR i := 3 TO LEN( aDbfs )
    NetUse( aDbfs[i]  , STD_RETRY, RDD_IN_USE, USE_SHARED, USE_NEW, NIL )
    (aDbfs[i])->( ordsetfocus(aNtxs[i]) )
NEXT


WHILE !C_proc->( EOF() )
      ResetArray(aTtlInfo)
      IF T_line->( DBSEEK(C_proc->Proc_id ) )
         WHILE t_line->cpproc_id == c_proc->proc_id
              TtlWafers(aTtlInfo)
              t_line->( DBSKIP() )
         END
         Tpqc07Update( aTtlInfo )
      ELSE
            t_pqc07->( AddRec(5) )
            t_pqc07->PCPROCNO :=c_proc->pcprocNo
            t_pqc07->PROC_ID  :=c_proc->Proc_id
            t_pqc07->PROC_NMH :=GetUniProc()//c_proc->Proc_nmh
            t_pqc07->( DBUNLOCK() )
      END
      c_proc->( DBSKIP() )
END

FOR i := 1 TO LEN( aDbfs )
    (aDbfs[i])->( DBCLOSEAREA() )
NEXT

oRep:cPrepDbf := "T_pqc07"
oRep:aDbs[1] := "T_pqc07"

RETURN NIL

/********************************************************/

FUNCTION ScrapRep(oRep)

LOCAL cTempDir := GetUserInfo():cTempDir
LOCAL aTpqc20 := { {"REPLINENO" ,"C", 10, 0 },;
                   {"B_PURP"    ,"C",  1, 0 },;
                   {"B_ID"      ,"C",  6, 0 },;
                   {"B_TYPE"    ,"C",  1, 0 },;
                   {"SIZE_ID"   ,"C",  4, 0 },;
                   {"VALUE_ID"  ,"C",  8, 0 },;
                   {"TOL_ID"    ,"C",  1, 0 },;
                   {"B_NCAPS"   ,"C",  2, 0 },;
                   {"DIEL_ID"   ,"C",  1, 0 },;
                   {"CP_QTYWS"  ,"N",  8, 0 },;
                   {"LP_QTYWS"  ,"L",  1, 0 },;
                   {"CP_QTYSS"  ,"N", 10, 0 },;
                   {"LP_QTYSS"  ,"L",  1, 0 },;
                   {"CP_QTYPS"  ,"N", 12, 0 },;
                   {"LP_QTYPS"  ,"L",  1, 0 },;
                   {"CP_DSTA"   ,"D",  8, 0 },;
                   {"CP_QTYWF"  ,"N",  8, 0 },;
                   {"LP_QTYWF"  ,"L",  1, 0 },;
                   {"CP_QTYSF"  ,"N", 10, 0 },;
                   {"LP_QTYSF"  ,"L",  1, 0 },;
                   {"CP_QTYPF"  ,"N", 12, 0 },;
                   {"LP_QTYPF"  ,"L",  1, 0 },;
                   {"CP_DFIN"   ,"D",  8, 0 },;
                   {"CP_TFIN"   ,"C",  5, 0 },;
                   {"CP_WIDFIN" ,"C",  5, 0 },;
                   {"CP_TOTSCRW","N",  8, 0 },;
                   {"LP_TOTSCRW","L",  1, 0 },;
                   {"CP_TOTSCRS","N", 10, 0 },;
                   {"LP_TOTSCRS","L",  1, 0 },;
                   {"CP_TOTSCRP","N", 12, 0 },;
                   {"LP_TOTSCRP","L",  1, 0 },;
                   {"SCR_COST"  ,"N", 12, 0 },;
                   {"B_LREJ"    ,"C",  4, 0 },;
                   {"DISC_REP"  ,"C",  6, 0 },;
                   {"ESN_ID"    ,"C", 16, 0 },;
                   {"ROUTE_ID"  ,"C",  4, 0 } }

IF !FILE(cTempDir+"T_pqc20")
   DBCREATE( cTempDir+"T_pqc20",aTpqc20,   )
ENDIF
IF SELECT("T_pqc20") > 0; CLOSE T_pqc20 ; ENDIF
NetUse( "T_pqc20", STD_RETRY,  , USE_EXCLUSIVE, USE_NEW, cTempDir )
T_pqc20->( __DBZAP() )

NetUse( "t_linemv", STD_RETRY,  , USE_EXCLUSIVE, USE_NEW, cTempDir )

oRep:cPrepDbf := "T_pqc20"



RETURN NIL
/********************************************************/

FUNCTION PmTtls(oRep) // for pqc03v1

LOCAL cTempDir := GetUserInfo():cTempDir
LOCAL aTpqc20 := { {"NO"        ,"C", 10, 0 },;
                   {"B_PURP"    ,"C",  1, 0 },;
                   {"B_ID"      ,"C",  6, 0 },;
                   {"B_TYPE"    ,"C",  1, 0 },;
                   {"SIZE_ID"   ,"C",  4, 0 },;
                   {"VALUE_ID"  ,"C",  8, 0 },;
                   {"TOL_ID"    ,"C",  1, 0 },;
                   {"B_NCAPS"   ,"C",  2, 0 },;
                   {"DIEL_ID"   ,"C",  1, 0 },;
                   {"CP_QTYWS"  ,"N",  8, 0 },;
                   {"LP_QTYWS"  ,"L",  1, 0 },;
                   {"CP_QTYSS"  ,"N", 10, 0 },;
                   {"LP_QTYSS"  ,"L",  1, 0 },;
                   {"CP_QTYPS"  ,"N", 12, 0 },;
                   {"LP_QTYPS"  ,"L",  1, 0 },;
                   {"CP_DFIN"   ,"D",  8, 0 },;
                   {"CP_TFIN"   ,"C",  5, 0 },;
                   {"CP_WIDFIN" ,"C",  5, 0 },;
                   {"ESN_ID"    ,"C", 16, 0 } }

IF !FILE(cTempDir+"T_pqc03")
   DBCREATE( cTempDir+"T_pqc03",aTpqc03,   )
ENDIF
IF SELECT("T_pqc20") > 0; CLOSE T_pqc03 ; ENDIF
NetUse( "T_pqc03", STD_RETRY,  , USE_EXCLUSIVE, USE_NEW, cTempDir )
T_pqc03->( __DBZAP() )

NetUse( "t_linemv", STD_RETRY,  , USE_EXCLUSIVE, USE_NEW, cTempDir )

oRep:cPrepDbf := "T_pqc03"

RETURN NIL

/*****************************************************************/
STATIC FUNCTION TtlWafers( aTtls )

LOCAL aLines := {"AP","AF","AL","AG","CP" }, n
n := ASCAN( aLines, LEFT( t_line->pline_id,2) )

IF t_line->b_purp $ "E"                         // "R" WASN'T NEEDED. D.Laor 29-09-97
   aTtls[C07_PERW] += t_line->cp_bqtyw
ELSEIF t_line->b_purp $ "9Q" .OR. n = 0         // r&d BATCHES WERE NOT
   aTtls[C07_POTW] += t_line->cp_bqtyw          // CALCULATED BY ITS PURPOSE
ELSE                                            // THE ELSEIF WAS --->   ELSEIF n == 0
   aTtls[n+3] += t_line->cp_bqtyw // the 1st 3 array elements are proc info
ENDIF
aTtls[C07_DAYS] += date() - t_line->pp_dfin
aTtls[C07_BATS] += 1

RETURN NIL

STATIC FUNCTION ResetArray( aTtlInfo )
LOCAL i
LOCAL nLen := LEN( aTtlInfo )

        aTtlInfo[ 1] := c_proc->PCPROCNO
        aTtlInfo[ 2] := c_proc->PROC_ID
        aTtlInfo[ 3] := SUBSTR(C_PROC->PROC_NME,1,19)
        FOR i :=  4 TO nLen
            aTtlInfo[i] := 0
        NEXT

RETURN NIL


STATIC FUNCTION UpdateC06Array( aTtlInfo, aTtls, cDbf)
LOCAL aLines := {"AP","AF","AL","AG","CP","CN","CL","CS","DB" }, n,Zmani
n := ASCAN( aLines, LEFT( (cDbf)->pline_id,2) )

IF n > 5
	n := 5
ENDIF

aTtlInfo[C06_TOW] += aTTls[TTL_TOW]
aTtlInfo[C06_TOP] += aTTls[TTL_TOP]
IF n == 0
   aTtlInfo[C06_OTP] += aTTls[TTL_TOP]
ELSE
   aTtlInfo[n+5] += aTTls[TTL_TOP]
ENDIF

RETURN NIL
*
STATIC FUNCTION UpdateC04Array( aTtlInfo, aTtls, cDbf )

LOCAL aLines := {"AP","AF","AL","AG","CP","CN","CL","CS","DB" }, n

n := ASCAN( aLines, LEFT( (cDbf)->pline_id,2) )


IF n > 5
	n := 5
ENDIF

aTtlInfo[C04_PTOP] += aTTls[TTL_TOP]
aTtlInfo[C04_CTOP] += aTTls[TTL_TOP]
aTtlInfo[5]        += aTTls[TTL_TOW]

IF n == 0
   aTtlInfo[C04_POTP] += aTTls[TTL_TOP]
   aTtlInfo[C04_COTP] += aTTls[TTL_TOP]
ELSEIF n<>2
   aTtlInfo[n+3] += aTTls[TTL_TOP]  // "P_APP" array element is at pos 4
   aTtlInfo[n+10] += aTTls[TTL_TOP]  // "C_APP" array element is at pos 10
ENDIF

RETURN NIL

/******************************************************************/
STATIC FUNCTION Tpqc07Update( aTtls )
LOCAL nGrandTotal := 0
LOCAL nDayAverage := 0
LOCAL i

FOR i := 4 TO 10
    nGrandTotal += aTtls[i]
NEXT
nDayAverage := IF( aTtls[C07_BATS] == 0,0,(aTtls[C07_DAYS] /  aTtls[C07_BATS]) )

T_pqc07->( AddRec(5) )
FOR i := 1 TO T_pqc07->( FCOUNT() ) - 2
    t_pqc07->( FIELDPUT( i, aTtls[i] ) )
NEXT
t_pqc07->( FIELDPUT( i++, nGrandTotal ) )
t_pqc07->( FIELDPUT( i  , nDayAverage ) )
t_pqc07->( DBUNLOCK() )

RETURN NIL
*
STATIC FUNCTION Tpqc06Update( aTtlInfo, nDays, nTarget )
LOCAL i
LOCAL nGap,zmani
aTtlInfo[C06_TARG] := nTarget
nGap := INT( ( (aTtlInfo[C06_TOW]/(nTarget*nDays)) - 1 ) * 100 )
aTtlInfo[C06_GAP] := IF( nGap >= -5, "    OK"," " + STR( nGap, 4, 0) +"%" )
t_pqc06->( AddRec(5) )
FOR i := 1 TO t_pqc06->( FCOUNT() )
    t_pqc06->( FIELDPUT( i, aTtlInfo[i] ) )
NEXT
t_pqc06->( DBUNLOCK() )

RETURN NIL
*
STATIC FUNCTION Tpqc04Update( aTtlInfo )

LOCAL i

t_pqc04->( AddRec(5) )
FOR i := 1 TO t_pqc04->( FCOUNT() )
    t_pqc04->( FIELDPUT( i, aTtlInfo[i] ) )
NEXT
t_pqc04->( DBUNLOCK() )

RETURN NIL

/******************************************************************/
STATIC FUNCTION Tpqc04TtlUpdate

DBSELECTAREA("t_pqc04")
CheckIndex("t_pqc04","t_pqc04","pcprocno",.T.)

t_pqc04->( DBGOBOTTOM() )
NCUMP_TOP  := t_pqc04->cump_top
NCUMP_APP  := t_pqc04->cump_APp
NCUMP_AFP  := t_pqc04->cump_AFp
NCUMP_ALP  := t_pqc04->cump_ALp
NCUMP_AGP  := t_pqc04->cump_AGp
NCUMP_CPP  := t_pqc04->cump_CPp
NCUMP_OTP  := t_pqc04->cump_OTp
t_pqc04->( DBSKIP(-1) )
WHILE !t_pqc04->( BOF() )
  t_pqc04->( RecLock(5,"",.F.) )

  t_pqc04->cump_top += NCUMP_TOP
  NCUMP_TOP  := t_pqc04->cump_top

  t_pqc04->cump_APP += NCUMP_APP
  NCUMP_APP  := t_pqc04->cump_APp

  t_pqc04->cump_AFp += NCUMP_AFP
  NCUMP_AFP  := t_pqc04->cump_AFp

  t_pqc04->cump_ALp += NCUMP_ALP
  NCUMP_ALP  := t_pqc04->cump_ALp

  t_pqc04->cump_AGp += NCUMP_AGP
  NCUMP_AGP  := t_pqc04->cump_AGp

  t_pqc04->cump_CPp += NCUMP_CPP
  NCUMP_CPP  := t_pqc04->cump_CPp

  t_pqc04->cump_oTp += NCUMP_OTP
  NCUMP_OTP  := t_pqc04->cump_oTp

  t_pqc04->( DBUNLOCK() )
  t_pqc04->( DBSKIP(-1) )
END
t_pqc04->( DBCOMMIT() )

NCUMP_APP:=NCUMP_AFP:=NCUMP_ALP:=NCUMP_AGP:=NCUMP_CPP:=NCUMP_OTP:=NCUMP_TOP:=0

RETURN NIL
/******************************************************************/
STATIC FUNCTION TtlQty( cDbf )

LOCAL nRetVal := 0
LOCAL nCurrentQty
LOCAL nTtlP := 0
LOCAL nTtlW := 0

   DO CASE
      CASE c_proc->uom_id = "W"
           nCurrentQty := (cDbf)->cp_bqtyw
      CASE c_proc->uom_id = "S"
           nCurrentQty := (cDbf)->cp_bqtys
      CASE c_proc->uom_id = "P"
           nCurrentQty := (cDbf)->cp_bqtyp
   ENDCASE

   d_line->( DBSEEK( (cDbf)->b_id ) )
   IF c_proc->calc_exp
      nRetVal := nCurrentQty * SuitCalc1( c_proc->uom_id, cDbf ) * d_line->expq_fctr
          IF c_proc->uom_id <> "W"
             nTtlW += qtyconvert(nRetVal,(cDbf)->b_type,(cDbf)->pline_id,(cDbf)->size_id,c_proc->uom_id,"W",,,(cDbf)->route_id,(cDbf)->value_id) //VR TID:2114110
          ELSE
             nTtlW += nRetVal
          ENDIF
          IF c_proc->uom_id <> "P"
               nTtlP += qtyconvert(nRetVal,(cDbf)->b_type,(cDbf)->pline_id,(cDbf)->size_id,c_proc->uom_id,"P",,,(cDbf)->route_id,(cDbf)->value_id) //VR TID:2114110
          ELSE
             nTtlP += nRetVal
       ENDIF
   ELSE
      nTtlW += (cDbf)->cp_bqtyw
      nTtlP += (cDbf)->cp_bqtyp
   ENDIF

RETURN { nTtlW, nTtlP }

FUNCTION Exc06Qty( cDbf,nQty,cLineDbf ) // Needed also for rpqc08v1.prg last parameter added.
LOCAL nRetVal := 0,cZmani
LOCAL nCurrentQty
LOCAL nTtlP := 0
LOCAL nTtlW := 0
DEFAULT cLineDbf to "d_line"

nCurrentQty := nQty
IF cLineDbf # cDbf
   (cLineDbf)->( DBSEEK( (cDbf)->b_id ) ) // D_line  „Ž ™…”‰‡
ENDIF
IF C_proc->Calc_exp
   nRetVal := nCurrentQty * SuitCalc1( C_proc->Uom_id, cDbf ) * (cLineDbf)->expq_fctr
ELSE
   nRetVal := nQty
ENDIF
*
*YH 27-11-00 ‡"…ƒ …‹‰‘ ˆ…˜‰”Œ
*IF C_proc->PcProcNo=='3400'
*   ?ALLTRIM(STR(nRetVal)) + ' :˜…–‰‰„ ‰”–' + ' :' + ALLTRIM(STR(C_expQty->Exp_yld)) + ' Œ…”‹ '+ T_lineMv->B_id + ' :„ŽŒ ‰‰‹˜ ' + ALLTRIM(STR(T_lineMv->Cp_bQtyP))
*ENDIF

RETURN nRetVal
*
FUNCTION Ttl06Qty( cDbf,cDbfLine,nSikum06)  // Needed also for RPQC08V1
LOCAL cUom:=0
LOCAL aRet
LOCAL nTtlP := 0
LOCAL nTtlW := 0
DEFAULT cDbfLine to "d_line"
aRet := GetMaxWSP(GetUomCp((cdbf)->cpproc_id) ,,(cDbf)->cp_totscr,cdbf )
nTtlW := (cDbf)->cp_bqtyw
nTtlp := Exc06Qty(cDbf,((cDbf)->cp_bqtyp - aRet[3]), cDbfLine)
RETURN { nTtlW, nTtlP }

/*
 * ÚÄ Function ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
 * ³         Name: SuitCalc1()           Docs:                                ³
 * ³  Description:                                                            ³
 * ³       Author: Danny Hazan                                                ³
 * ³ Date created: 06-25-96              Date updated: þ06-25-96              ³
 * ³ Time created: 02:28:17pm            Time updated: þ02:28:17pm            ³
 * ³    Copyright: AVX                                                        ³
 * ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´
 * ³    Arguments: cWSP                                                       ³
 * ³ Return Value: nRetVal                                                    ³
 * ³     See Also: I had to copy this because of the non-generic versions
 * ³               that danny made
 * ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
 */
STATIC FUNCTION SuitCalc1( cWSP, cDbf )
LOCAL lExit := .F. , nRetVal := 0
LOCAL xKey := cWSP + (cDbf)->b_type + (cDbf)->size_id
LOCAL bKey := COMPILE( c_expqty->(ordKey()) )
DBSELECTAREA("c_expqty") // UOM+B_type+Size ‰”Œ ‰…ŽŽ
IF C_expqty->( DbSeek( xKey ) )
   WHILE bKey:eval() == xKey
         IF EMPTY( (cDbf)->Tol_id )
            lExit := ( (cDbf)->Value_id >= C_expqty->lval_lim .AND.;
                       (cDbf)->Value_id <= C_expqty->hval_lim )
         ELSE
            lExit := ( (cDbf)->Value_id >= C_expqty->Lval_lim   .AND.;
                       (cDbf)->Value_id <= C_expqty->Hval_lim ) .AND. ;
                       (cDbf)->Tol_id   == C_expqty->Tol_id
         ENDIF
         IF lExit
            nRetVal := C_expqty->Exp_yld
            EXIT
         ENDIF

         C_expqty->( DBSKIP() )
   ENDDO
ENDIF
RETURN nRetVal
*
FUNCTION Prep13pqc(oRep)
// I don't think this function is used anymore - YG 12/11/97 see rpqc13v1.prg
LOCAL cTempDir := GetUserInfo():cTempDir
LOCAL aTpqc13 := { {"B_ID"      ,"C",  6, 0 },;
                   {"B_TYPE"    ,"C",  1, 0 },;
                   {"B_PURP"    ,"C",  1, 0 },;
                   {"B_stat"    ,"C",  1, 0 },;
                   {"B_PRWK"    ,"C",  7, 0 },;
                   {"PTYPE_ID"  ,"C",  1, 0 },;
                   {"PLINE_ID"  ,"C",  3, 0 },;
                   {"SIZE_ID"   ,"C",  4, 0 },;
                   {"VALUE_ID"  ,"N",  8, 3 },;
                   {"TOL_ID"    ,"C",  1, 0 },;
                   {"VOLT_ID"    ,"C",  1, 0 },;
                   {"TC_ID"    ,"C",  1, 0 },;
                   {"TERM_ID"    ,"C",  1, 0 },;
                   {"ESNY_ID"    ,"C",  1, 0 },;
                   {"ESNXX_ID"  ,"C",  2, 0 },;
                   {"QTY_BINIW" ,"N",  4, 0 },;
                   {"QTY_BINIP" ,"N",  8, 0 },;
                   {"PP_BQTYW"  ,"N",  4, 0 },;
	                   {"PP_BQTYP"  ,"N",  8, 0 },;
                   {"QTY_OKP"   ,"N",  8, 0 },;
                   {"QTY_OUTP"  ,"N",  8, 0 },;
                   {"QTY_RP"    ,"N",  8, 0 },;
                   {"Tech_yld"  ,"N",  8, 2 },;
                   {"DADD_REC"  ,"D",  8, 0 },;
                   {"PROD_DAYS" ,"N",  4, 0 } }

LOCAL aDbfs := {"T_pqc13","t_stock","c_proc"  ,"c_thqty","c_tthqty","c_expqty","d_line","H_FINQC" }
LOCAL aNtxs := {""       , ""     ,"iproc_id","itsuom" ,"itsuom" ,"iuts"    ,"ib_idln","hfinqcbn"}
LOCAL nQty_biniw,nqty_binip,nqty_okp,nqty_outp,nqty_rp,i
LOCAL dAddRec := CTOD("  /  /  ")
nQty_biniw := 0
nqty_binip := 0
nqty_okp    := 0
nqty_outp   := 0
nqty_rp     := 0

IF !FILE(cTempDir+"T_pqc13")
   DBCREATE( cTempDir+"T_pqc13",aTpqc13,   )
ENDIF
IF SELECT("T_pqc13") > 0; CLOSE T_pqc13 ; ENDIF
NetUse( "T_pqc13", STD_RETRY,  , USE_EXCLUSIVE, USE_NEW, cTempDir )
T_pqc13->( __DBZAP() )
CheckIndex("t_pqc13","t_pqc13","b_id",.T.)

NetUse( "t_STOCK", STD_RETRY,  , USE_EXCLUSIVE, USE_NEW, cTempDir )

CheckIndex("t_stock","t_stock","b_id",.T.)

FOR i := 3 TO LEN( aDbfs )
    IF Select(aDbfs[i])=0
       NetUse( aDbfs[i]  , STD_RETRY, RDD_IN_USE, USE_SHARED, USE_NEW, NIL )
    ENDIF
    (aDbfs[i])->( ORDSETFOCUS(aNtxs[i]) )
NEXT

t_stock->( DBGOTOP())

WHILE !t_stock->( EOF() )
   IF D_line->(DbSeek(t_Stock->b_id))
     IF D_line->UOM_INI="W"
          nQty_biniw := d_line->QTY_BINI
     ELSE
          nQty_biniw := qtyconvERT(D_LINE->QTY_BINI,D_LINE->B_TYPE,D_LINE->PLINE_ID,D_LINE->SIZE_ID,d_line->UOM_INI,"W",.T.,,d_line->route_id,d_line->value_id) //VR TID:2114110
     ENDIF

     IF D_line->UOM_INI="P"
          nQty_binip := d_line->QTY_BINI
     ELSE
          nQty_binip := qtyconvERT(D_LINE->QTY_BINI,D_LINE->B_TYPE,D_LINE->PLINE_ID,D_LINE->SIZE_ID,d_line->UOM_INI,"P",.T.,,d_line->route_id,d_line->value_id) //VR TID:2114110
     ENDIF

     H_FINQC->(DBSeek(T_STOCK->B_ID))

     nqty_okp    := 0
     nqty_outp   := 0
     nqty_rp     := 0

     DO WHILE T_STOCK->B_ID == H_FINQC->B_ID
          IF IsValSubst(oRep,"D_LINE","H_FINQC")
               nQTY_OKP  := nQTY_OKP + H_FINQC->SM_QTY
          ELSE
               If .not. H_FINQC->TOL_ID $ "R "  /* exclude "R" and " "! */
                    nQTY_OUTP := nQTY_OUTP + H_FINQC->SM_QTY
               Endif

               If H_FINQC->TOL_ID $ "R "  /* "R" only!  */
                    nQTY_RP := nQTY_RP + H_FINQC->SM_QTY
               Endif
          ENDIF
         dAddRec := H_finQc->b_dqc
         H_FINQC->(dbSkip())
     ENDDO
     IF ! t_pqc13->(DBSeek(T_STOCK->B_ID))
          t_pqc13->( AddRec(5) )

          t_pqc13->B_ID              := T_STOCK->B_ID
          t_pqc13->B_TYPE            := D_LINE->B_TYPE
          t_pqc13->B_PURP            := D_LINE->B_PURP
          t_pqc13->B_STAT            := D_LINE->B_STAT
          t_pqc13->B_PRWK            := D_LINE->B_PRIOR + D_LINE->B_WKPROM
          t_pqc13->PTYPE_ID          := D_LINE->PTYPE_ID
          t_pqc13->PLINE_ID          := D_LINE->PLINE_ID
          t_pqc13->SIZE_ID           := D_LINE->SIZE_ID
          t_pqc13->VALUE_ID          := D_LINE->VALUE_ID
          t_pqc13->TOL_ID            := D_LINE->TOL_ID
          t_pqc13->ESNXX_ID          := D_LINE->ESNXX_ID
          t_pqc13->VOLT_ID           := D_LINE->VOLT_ID
          t_pqc13->TC_ID             := D_LINE->TC_ID
          t_pqc13->TERM_ID           := D_LINE->TERM_ID
          t_pqc13->ESNY_ID           := D_LINE->ESNY_ID
          t_pqc13->QTY_BINIW         := nQty_biniw
          t_pqc13->QTY_BINIP         := nQty_binip
          t_pqc13->PP_BQTYW          := D_LINE->PP_BQTYW
          t_pqc13->PP_BQTYP          := nQTY_OKP+nQTY_OUTP+nQTY_RP   //D_LINE->PP_BQTYP
          t_pqc13->QTY_OKP           := nQTY_OKP
          t_pqc13->QTY_OUTP          := nQTY_OUTP
          t_pqc13->QTY_RP            := nQTY_RP
          t_pqc13->TECH_YLD          := round((nQTY_OKP/nQTY_BINIP) * 100,2 )
          t_pqc13->DADD_REC          := dAddRec
          t_pqc13->PROD_DAYS         := IF(EMPTY(dAddRec),date() - D_LINE->B_DOPEN,dAddRec - D_LINE->B_DOPEN)

     ENDIF
   ENDIF

   @ 2,50 Say t_stock->(recno())
   t_stock->( DBSKIP() )
END

FOR i := 1 TO LEN( aDbfs )
    (aDbfs[i])->( DBCLOSEAREA() )
NEXT

oRep:cPrepDbf := "T_pqc13"

RETURN NIL

Function GetGap(nTarg,nWaf)
LOCAL aToFrom := GetBuffer("Date from...to")
LOCAL nGapInDays := GetBuffer("Gap in days")
LOCAl nGap,nDays := 1
LOCAL cRet :=""

IF EMPTY( aToFrom[1] )
   aTofrom[1] := date()
   aTofrom[2] := date()
ENDIF
nDays := aToFrom[2] - aToFrom[1] + 1
*
* YH 14-11-00
* nwaf/Š˜’ .0 š…Ž‹ ‰„Œ „‹‰˜– "T" ‘…ˆˆ‘ š˜‰‡ …‰”Œ Task: 00111302 š…—’
*           !0 Ž „…™ nwaf ™ ‰€š ™…‡‰ ‰…ˆ‰„™ ‰€š ‰š”‘…„ ! „€‰‚™ ˜–…‰
*
IF nwaf#0
   nGap := INT( ( (nwaf/(nTarg*nGapInDays)) - 1 ) * 100 )
   cRet := IF( nGap >= -5, "    OK"," " + STR( nGap, 4, 0) +"%" )
ENDIF
RETURN cRet