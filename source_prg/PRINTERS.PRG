/*
 * ÚÄ Program ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
 * ³  Application: AVXBMS                                                     ³
 * ³    File Name: PRINTERS.PRG                                               ³
 * ³  Description: All functions to handle the printers.dbf for user-defined  ³
 * ³             : printing                                                   ³
 * ³             :                                                            ³
 * ³             :                                                            ³
 * ³       Author: Shalom LeVine         Tester:                              ³
 * ³ Date created: 11-24-96              Date updated: ş11-24-96              ³
 * ³ Time created: 10:02:10am            Time updated: ş10:02:10am            ³
 * ³    Make File: AVXBMS.RMK                                                 ³
 * ³    Exec File: AVXBMS.EXE            Docs By:                             ³
 * ³    DBFs/NTXs:                                                            ³
 * ³             :                                                            ³
 * ³             :                                                            ³
 * ³             :                                                            ³
 * ³             :                                                            ³
 * ³             :                                                            ³
 * ³    Copyright: (c) 1996 by AVX                                            ³
 * ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
 */

#include "avxdefs.ch"

#xcommand @ <row>, <col> ESCAPEGET <var>                                ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [SEND <msg>]                                    ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
           GetList,                                                     ;
           _GET_( <var>, <"var">, <pic>, <{valid}>, <{when}> ):display();
             )                                                          ;
       ; ATail(GetList):Reader := {|x| EscapeReader(x) }                ;
      [; ATail(GetList):<msg>]

#xcommand @ <row>, <col> BINARYGET <var>                                ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [SEND <msg>]                                    ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
           GetList,                                                     ;
           _GET_( <var>, <"var">, <pic>, <{valid}>, <{when}> ):display();
             )                                                          ;
       ; ATail(GetList):Reader := {|x| BinaryReader(x) }                ;
      [; ATail(GetList):<msg>]

STATIC oPc
STATIC bKey
STATIC cSearcher
STATIC cProgId
STATIC aColos := {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19}
STATIC aCurrentCols
STATIC bApplyKeys

/************************************************************************/
FUNCTION ThePrinters

LOCAL aOldList
LOCAL aKeyList :=  {{ "Ins             - Add a printer                 ",K_INS        },;
                    { "Enter           - Update printer                ",K_ENTER      },;
                    { "Control-Enter   - Copy this driver to new record",K_CTRL_ENTER } }

DEFAULT cProgId    TO "Printers" ,;
        bApplyKeys TO {|nKey , o , oForm | ApplyKeys(nKey,o,oForm ) } ,;
        aColos     TO {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19}

IF Empty( aCurrentCols )
   aCurrentCols := aColos
ENDIF

aOldList := SetHlpKeys( aKeyList )

oPc := BroCenter():new( "printers" , aCurrentCols, FALSE )

oPc:oForm :=  Form():new(,,,, "Printer Center",, ProcName() )

oPc:createBrowse():addDefaultCols()

oPc:ApplyBlock := bApplyKeys
oPc:process()

aCurrentCols := oPc:CurrentCols()

oPc:oForm:hide()

oPc := NIL

aKeyList := SetHlpKeys( aOldList )

RETURN NIL
/************************************************************************/

STATIC FUNCTION ApplyKeys( nKey , oBroCenter , oForm )

LOCAL lProcessed := .T.
LOCAL nSaveRec
LOCAL cClr

DO CASE
   CASE nKey == K_DEL
        IF ReadOnly()
           Msg24( {MSG_RO}, 3, TRUE )
        ELSE
           TONE(800,0)
           IF ALERT("Are you sure you want to;delete this printer?",{"Yes","No"}, ALERT_STD) == 1
              Printers->( RecLock(5,"",.F.) )
              Printers->( DBDELETE() )
              Printers->( DBUNLOCK() )
              Printers->( DBSKIP(0) )
           ENDIF
        END
   CASE nKey = K_INS .OR. nKey == K_ENTER
       IF ReadOnly()
         Msg24( {MSG_RO}, 3, TRUE )
       ELSE
         scrnPush(2,0,24,79, .T. )
         cClr := SETCOLOR( "gr+/b,gr+/r,b/b,b/b,w+/b")
         SayPrnCaptions()
         IF( nKey == K_INS, GetPrnValues(TRUE), GetPrnValues(FALSE) )
         SETCOLOR(cClr)
         scrnPop()
       END
   CASE nKey = K_CTRL_ENTER
        CopyThisRecord()
   OTHERWISE
        lProcessed := .F.
ENDCASE
oBroCenter:oBro:refreshAll()


RETURN lProcessed


/************************************************************************/
STATIC FUNCTION SayPrnCaptions()


@ 2, 0 SAY "Lan Server  :"
@ 3, 0 SAY "Queue       :"
@ 4, 0 SAY "Printer name:"
@ 5, 0 SAY "Groups      :"
@ 6, 0 SAY "Lanprinter  :"
@ 7, 0 SAY "Setup       :"
@ 8, 0 SAY "Reset       :"
@ 9, 0 SAY "6lpi        :"
@10, 0 SAY "8lpi        :"
@11, 0 SAY "10cpi       :"
@12, 0 SAY "12cpi       :"
@13, 0 SAY "Compr       :"
@14, 0 SAY "Portrait    :"
@15, 0 SAY "Landscape   :"
@16, 0 SAY "Bold On     :"
@17, 0 SAY "Bold Off    :"
@18, 0 SAY "Underline On:"
@19, 0 SAY "Underln Off :"
@20, 0 SAY "Italics On  :"
@21, 0 SAY "Italics Off :"
@ 24, 0 SAY SPACE(80) COLOR "W+/r"

RETURN NIL
/************************************************************************/
STATIC FUNCTION GetPrnValues( lNew )

LOCAL   cLan_server  := IF( lNew, SPACE(25) , Printers->lan_server        )
LOCAL   cQueue       := IF( lNew, SPACE(15) , Printers->Queue             )
LOCAL   cPr_name     := IF( lNew, SPACE(25) , Printers->Pr_name           )
LOCAL   cGroups      := IF( lNew, REPLICATE("0",20) , Printers->Groups    )
LOCAL   cLanprinte   := IF( lNew, REPLICATE("0",20) , Printers->Lanprinter        )
LOCAL   cPr_setup    := IF( lNew, SPACE(60) , LEFT(Printers->Pr_setup,60) )
LOCAL   cPr_reset    := IF( lNew, SPACE(60) , LEFT(Printers->Pr_reset,60) )
LOCAL   cPr_6lpi     := IF( lNew, SPACE(60) , LEFT(Printers->Pr_6lpi ,60) )
LOCAL   cPr_8lpi     := IF( lNew, SPACE(60) , LEFT(Printers->Pr_8lpi ,60) )
LOCAL   cPr_10cpi    := IF( lNew, SPACE(60) , LEFT(Printers->Pr_10cpi,60) )
LOCAL   cPr_12cpi    := IF( lNew, SPACE(60) , LEFT(Printers->Pr_12cpi,60) )
LOCAL   cPr_compr    := IF( lNew, SPACE(60) , LEFT(Printers->Pr_compr,60) )
LOCAL   cPr_port     := IF( lNew, SPACE(60) , LEFT(Printers->Pr_port ,60) )
LOCAL   cPr_land     := IF( lNew, SPACE(60) , LEFT(Printers->Pr_land ,60) )
LOCAL   cPr_bdon     := IF( lNew, SPACE(60) , LEFT(Printers->Pr_bdon ,60) )
LOCAL   cPr_bdoff    := IF( lNew, SPACE(60) , LEFT(Printers->Pr_bdoff,60) )
LOCAL   cPr_ulon     := IF( lNew, SPACE(60) , LEFT(Printers->Pr_ulon ,60) )
LOCAL   cPr_uloff    := IF( lNew, SPACE(60) , LEFT(Printers->Pr_uloff,60) )
LOCAL   cPr_iton     := IF( lNew, SPACE(60) , LEFT(Printers->Pr_iton ,60) )
LOCAL   cPr_itoff    := IF( lNew, SPACE(60) , LEFT(Printers->Pr_itoff,60) )

SETCURSOR(1)

@ 2,14 GET cLan_Server
@ 3,14 GET cQueue
@ 4,14 GET cPr_name
@ 5,14 BINARYGET cGroups
@ 6,14 BINARYGET cLanprinte
@ 7,14 ASCIIGET cPr_setup
@ 8,14 ASCIIGET cPr_reset
@ 9,14 ASCIIGET cPr_6lpi
@10,14 ASCIIGET cPr_8lpi
@11,14 ASCIIGET cPr_10cpi
@12,14 ASCIIGET cPr_12cpi
@13,14 ASCIIGET cPr_compr
@14,14 ASCIIGET cPr_port
@15,14 ASCIIGET cPr_land
@16,14 ASCIIGET cPr_bdon
@17,14 ASCIIGET cPr_bdoff
@18,14 ASCIIGET cPr_ulon
@19,14 ASCIIGET cPr_uloff
@20,14 ASCIIGET cPr_iton
@21,14 ASCIIGET cPr_itoff

READ
SETKEY( K_F2, NIL )

IF LASTKEY() <> K_ESC
   IF( lNew, printers->( AddRec(5,"printers") ), printers->( RecLock(5,"printers") ) )
   Printers->lan_server:= cLan_Server
   Printers->Queue     := cQueue
   Printers->Pr_name   := cPr_name
   Printers->Groups    := cGroups
   Printers->Lanprinter:= cLanprinte
   Printers->Pr_setup  := cPr_setup
   Printers->Pr_reset  := cPr_reset
   Printers->Pr_6lpi   := cPr_6lpi
   Printers->Pr_8lpi   := cPr_8lpi
   Printers->Pr_10cpi  := cPr_10cpi
   Printers->Pr_12cpi  := cPr_12cpi
   Printers->Pr_compr  := cPr_compr
   Printers->Pr_port   := cPr_port
   Printers->Pr_land   := cPr_land
   Printers->Pr_bdon   := cPr_bdon
   Printers->Pr_bdoff  := cPr_bdoff
   Printers->Pr_ulon   := cPr_ulon
   Printers->Pr_uloff  := cPr_uloff
   Printers->Pr_iton   := cPr_iton
   Printers->Pr_itoff  := cPr_itoff
   printers->( dbcommit() )
   printers->( dbunlock() )
END

SETCURSOR(0)
RETURN NIL

/*******************************************************************
*
* PROCEDURE EscapeReader( oGet )
* Custom reader for Printers, hotkey to imbed CHR(K_ESC) in a GET
*/
PROCEDURE EscapeReader( oGet )


// don't allow ever to use insert mode in this GET
SET( _SET_INSERT, .F. )

   // Read the GET if the WHEN condition is satisfied
   IF ( GetPreValidate( oGet ) )

      // Activate the GET for reading
      oGet:setFocus()

      WHILE ( oGet:exitState == GE_NOEXIT )

         // Check for initial typeout (no editable positions)
         IF ( oGet:typeOut )
            oGet:exitState := GE_ENTER
      ENDIF

      @ 24, 0 SAY PADC("Press F2 for an ESC character () ",80) COLOR "W+/r"
      DEVPOS( oGet:row, oGet:Col )
         // Apply keystrokes until exit
         WHILE ( oGet:exitState == GE_NOEXIT )
            EscapeApplyKey( oGet, inkey( 0 ) )
         ENDDO

         // Disallow exit if the VALID condition is not satisfied
         IF ( !GetPostValidate( oGet ) )
            oGet:exitState := GE_NOEXIT
         ENDIF
      ENDDO

      // De-activate the GET
      oGet:killFocus()

   ENDIF

   @ 24, 0 SAY SPACE(80) COLOR "W+/r"

   RETURN

/***
*
*  ESCAPEApplyKey()
*
*  Custom applykey for pline query
*/
PROCEDURE EscapeApplyKey( oGet, nKey )

   LOCAL cKey
   LOCAL bKeyBlock
  // LOCAL cTempBuffer

   // Check for SET KEY first
   IF !( ( bKeyBlock := setkey( nKey ) ) == NIL )
      GetDoSetKey( bKeyBlock, oGet )
      RETURN                           // NOTE
   ENDIF

   DO CASE
   CASE ( nKey == K_UP )
      oGet:exitState := GE_UP

   CASE ( nKey == K_SH_TAB )
      oGet:exitState := GE_UP

   CASE ( nKey == K_DOWN )
      oGet:exitState := GE_DOWN

   CASE ( nKey == K_TAB )
      oGet:exitState := GE_DOWN

   CASE ( nKey == K_ENTER )
      oGet:exitState := GE_ENTER

   CASE ( nKey == K_ESC )
      IF ( SET( _SET_ESCAPE ) )

         oGet:undo()
         oGet:exitState := GE_ESCAPE

      ENDIF

   CASE ( nKey == K_PGUP )
      oGet:exitState := GE_WRITE

   CASE ( nKey == K_PGDN )
      oGet:exitState := GE_WRITE

   CASE ( nKey == K_CTRL_HOME )
      oGet:exitState := GE_TOP


//#ifdef CTRL_END_SPECIAL

   // changed by Shalom 04-08-96 for multiple edits
   CASE ( nKey == K_CTRL_ENTER )
      oGet:exitState := GE_WRITE

//#else

   // Both ^W and ^End terminate the READ (the default)
   CASE ( nKey == K_CTRL_W )
      oGet:exitState := GE_WRITE

//#endif

   CASE ( nKey == K_HOME )
      oGet:home()

   CASE ( nKey == K_END )
      oGet:end()

   CASE ( nKey == K_RIGHT )
         oGet:right()

   CASE ( nKey == K_LEFT )
         oGet:left()

   CASE nKey == K_DEL
      oGet:delete()

   CASE nKey == K_BS
      oGet:backSpace()

   // special case for this GET, embeds a CHR(27) in the GET
   CASE nKey == K_F2
       oGet:Buffer :=  STUFF( oGet:Buffer, oGet:Pos, 1, "")
       oGet:right()
       oGet:Display()

   OTHERWISE

      IF ( nKey >= 32 .AND. nKey <= 255 )

         cKey := CHR( nKey )

            IF ( SET( _SET_INSERT ) )
               oGet:insert( cKey )
            ELSE
               oGet:overstrike( cKey )
            ENDIF

         IF ( oGet:typeOut )
            IF ( SET( _SET_BELL ) )
            ?? CHR(7)
            ENDIF

            IF ( !SET( _SET_CONFIRM ) )
            oGet:exitState := GE_ENTER
            ENDIF
         ENDIF

      END

   ENDCASE

  RETURN

/*******************************************************************
*
* PROCEDURE BinaryReader( oGet )
* Custom reader for Printers, Control Access to binary string for Groups
*/
PROCEDURE BinaryReader( oGet )

LOCAL nSelect := SELECT()
LOCAL aGroups[20]

AFILL( aGroups, "No group            ")
// the file MUST be open with the gpprnpos.ntx, to maintain correct order of records
NetUse("d_groups",5)
d_groups->( ordsetfocus("gpprnpos") )

WHILE !D_groups->( EOF() )
      aGroups[d_groups->prn_pos] := LEFT( d_groups->group_nm,20)
      d_groups->( DBSKIP() )
      IF d_groups->prn_pos > 20 // max length of GET
         EXIT
      ENDIF
END
d_groups->( DBCLOSEAREA() )
SELECT (nSelect)

// don't allow ever to use insert mode in this GET
SET( _SET_INSERT, .F. )

   // Read the GET if the WHEN condition is satisfied
   IF ( GetPreValidate( oGet ) )

      // Activate the GET for reading
      oGet:setFocus()

      WHILE ( oGet:exitState == GE_NOEXIT )

         // Check for initial typeout (no editable positions)
         IF ( oGet:typeOut )
            oGet:exitState := GE_ENTER
      ENDIF

      @ 24, 0  SAY aGroups[1] COLOR "bg+/r"
      @ 24, 23 SAY "SPACE - flag/unflag, CURSOR KEYS - move, ENTER - finish" COLOR "gr+/r"
      DEVPOS( oGet:Row,oGet:Col)
         // Apply keystrokes until exit
         WHILE ( oGet:exitState == GE_NOEXIT )
            BinaryApplyKey( oGet, inkey( 0 ), aGroups )
      ENDDO

         // Disallow exit if the VALID condition is not satisfied
         IF ( !GetPostValidate( oGet ) )
            oGet:exitState := GE_NOEXIT
         ENDIF
      ENDDO

      // De-activate the GET
      oGet:killFocus()

   ENDIF

   @ 24, 0 SAY SPACE(80) COLOR "W+/r"

   RETURN

/***
*
*  PlineApplyKey()
*
*  Custom applykey for pline query
*/
PROCEDURE BinaryApplyKey( oGet, nKey, aGroups )

   LOCAL cKey
   LOCAL bKeyBlock

   // Check for SET KEY first
   IF !( ( bKeyBlock := setkey( nKey ) ) == NIL )
      GetDoSetKey( bKeyBlock, oGet )
      RETURN                           // NOTE
   ENDIF

   DO CASE
   CASE ( nKey == K_UP )
      oGet:exitState := GE_UP

   CASE ( nKey == K_SH_TAB )
      oGet:exitState := GE_UP

   CASE ( nKey == K_DOWN )
      oGet:exitState := GE_DOWN

   CASE ( nKey == K_TAB )
      oGet:exitState := GE_DOWN

   CASE ( nKey == K_ENTER )
      oGet:exitState := GE_ENTER

   CASE ( nKey == K_ESC )
      IF ( SET( _SET_ESCAPE ) )

         oGet:undo()
         oGet:exitState := GE_ESCAPE

      ENDIF

   CASE ( nKey == K_PGUP )
      oGet:exitState := GE_WRITE

   CASE ( nKey == K_PGDN )
      oGet:exitState := GE_WRITE

   CASE ( nKey == K_CTRL_HOME )
      oGet:exitState := GE_TOP


//#ifdef CTRL_END_SPECIAL

   // changed by Shalom 04-08-96 for multiple edits
   CASE ( nKey == K_CTRL_ENTER )
      oGet:exitState := GE_WRITE

//#else

   // Both ^W and ^End terminate the READ (the default)
   CASE ( nKey == K_CTRL_W )
      oGet:exitState := GE_WRITE

//#endif

   CASE ( nKey == K_HOME )
        oGet:home()
        RefreshLine24(oGet, aGroups)

   CASE ( nKey == K_END )
        oGet:end()
        RefreshLine24(oGet, aGroups)

   CASE ( nKey == K_RIGHT )
        oGet:right()
        RefreshLine24(oGet, aGroups)

   CASE ( nKey == K_LEFT )
        oGet:left()
        RefreshLine24(oGet, aGroups)

   CASE  nKey == K_SPACE

            cKey := IF( SUBSTR( oGet:Buffer, oGet:Pos, 1 ) == "0", "1", "0" )

            //oGet:overstrike( cKey )
            oGet:Buffer := STUFF( oGet:Buffer, oGet:pos,1, cKey)
            oGet:VarPut(oGet:Buffer )
            oGet:Display()

         IF ( oGet:typeOut )
            IF ( SET( _SET_BELL ) )
            ?? CHR(7)
            ENDIF

            IF ( !SET( _SET_CONFIRM ) )
            oGet:exitState := GE_ENTER
            ENDIF
         ENDIF

   ENDCASE

  RETURN

STATIC FUNCTION RefreshLine24(oGet, aGroups)

LOCAL nCol := oGet:Col+oGet:Pos-1

@ 24, 0  SAY aGroups[oGet:pos] COLOR  "bg+/r"
DEVPOS(oGet:Row,nCol)

RETURN (NIL)

/***************************************************************/
STATIC FUNCTION CopyThisRecord()

LOCAL aInfo := {}, i

FOR i := 1 TO printers->( FCOUNT() )
    AADD( aInfo, printers->( FIELDGET(i) ) )
NEXT

Printers->( AddRec(5) )
FOR i := 1 TO printers->( FCOUNT() )
    printers->( FIELDPUT(i, aInfo[i]) )
NEXT

RETURN NIL

