/*
 * RPQC01V4.PRG — Production Line Report (Harbour version)
 * Replaces: RPQC01V1.PRG + RPQC01V4.RH2
 * Output: HTML file opened in default browser
 *
 * Port of AVX Israel Clipper report to Harbour
 * Original authors: Vitaly R., Shai B., Yossi G., David L.
 * Harbour port: 2026
 */

#include "inkey.ch"
#include "dbstruct.ch"
#include "setcurs.ch"

// Forward declarations for UDFs at bottom of file

STATIC aBuffer     // user criteria selections
STATIC aSortList   // sort display names
STATIC aSortExpr   // sort index expressions
STATIC aCritList   // criteria display names
STATIC aCritCheck  // criteria checkbox states (.T./.F.)
STATIC aCritValues // criteria selected values
STATIC cSortType   // selected sort
STATIC cDestType   // selected destination
STATIC cSubTitle   // user subtitle
STATIC cRepDbf     // working temp dbf alias

// ============================================================
FUNCTION Main()
// ============================================================
LOCAL cDataPath

   SET DATE FORMAT "dd/mm/yyyy"
   SET CENTURY ON
   SET DECIMALS TO 2
   SET DELETED ON

   // ADS or local DBF path — adjust to your environment
   cDataPath := GetDataPath()

   CLS
   SetColor( "W+/B" )
   @ 0, 0 SAY PadC( "AVX Israel — Production Line Report System", 80 ) COLOR "GR+/RB"

   MyRpqc01v4( cDataPath )

RETURN NIL


// ============================================================
FUNCTION MyRpqc01v4( cDataPath )
// ============================================================
LOCAL nSort, nDest
LOCAL cHtmlFile

   // ---- Setup sort options (same as original "new" variant) ----
   aSortList := { ;
      "Proc(D)+Rep.Stage+Hours At"            , ;
      "Type+Line+Size+Proc(D)+R.St.+Ho.At"    , ;
      "Type+Line+Size+Val+Proc(D)+R.St+Ho.At" , ;
      "Wrkst+Proc(D)+Rep.Stage+Hours At"      , ;
      "Proj+Line+Size+Val+Prior+Slack"           ;
   }

   aSortExpr := { ;
      "Str(10000-Val(cp_pccode),5)+RepStage(Alias())+Str(100000-HoursAt(Alias()))"                                                              , ;
      "ptype_id+pline_id+size_id+Str(10000-Val(cp_pccode),5)+RepStage(Alias())+Str(100000-HoursAt(Alias()))"                  , ;
      "ptype_id+pline_id+size_id+Str(value_id,8,3)+Str(10000-Val(cp_pccode),5)+RepStage(Alias())+Str(100000-HoursAt(Alias()))", ;
      "cpwkstn_id+Str(10000-Val(cp_pccode),5)+RepStage(Alias())+Str(100000-HoursAt(Alias()))"                                 , ;
      "esn_p2+pline_id+size_id+Str(value_id,8,3)+b_prior+Str(Slack_01(Alias()))"                                                ;
   }

   // ---- Setup criteria ----
   aCritList := { ;
      "Scheduled finish date"      , ;
      "Min&Max Hours at Rep.Stage" , ;
      "Status"                     , ;
      "Purpose"                    , ;
      "Priority"                   , ;
      "Product type"               , ;
      "Product line"               , ;
      "Size"                       , ;
      "Value"                      , ;
      "Promised Batches"           , ;
      "Project"                    , ;
      "Workstation"                , ;
      "Processes"                  , ;
      "Pc_code"                    , ;
      "Remark"                     , ;
      "Dialecter"                  , ;
      "Route filter"               , ;
      "Slack from...to"            , ;
      "Reporting Stage"            , ;
      "ESN(XX)"                    , ;
      "ESN(Y)"                       ;
   }

   aCritCheck  := Array( Len(aCritList) )
   aCritValues := Array( Len(aCritList) )
   AFill( aCritCheck, .F. )
   AFill( aCritValues, NIL )

   // ---- Show TUI selection screen ----
   IF ! ReportFace( "RPQC01V4", "Production line report" )
      RETURN NIL
   ENDIF

   // ---- Open databases ----
   IF ! OpenDatabases( cDataPath )
      Alert( "Cannot open database files!" )
      RETURN NIL
   ENDIF

   // ---- Filter and create temp DBF ----
   @ 12, 25 SAY "  Filtering data...  " COLOR "W+/R"
   CreateFilteredDbf( cDataPath )

   // ---- Index temp file ----
   nSort := AScan( aSortList, cSortType )
   IF nSort == 0
      nSort := 1
   ENDIF
   IndexTempFile( nSort )

   // ---- Generate output ----
   DO CASE
   CASE cDestType == "HTML (Browser)"
      cHtmlFile := GenHtmlReport( nSort )
      @ 24, 0 SAY PadR( "Report saved to: " + cHtmlFile, 80 ) COLOR "GR+/B"
      // Open in default browser
      hb_run( 'start "" "' + cHtmlFile + '"' )

   CASE cDestType == "Screen"
      BrowseReport()

   CASE cDestType == "CSV (Excel)"
      GenCsvExport()

   ENDCASE

   // ---- Cleanup ----
   CloseAllFiles()

RETURN NIL


// ============================================================
// TUI INTERFACE — prnFace equivalent
// ============================================================
FUNCTION ReportFace( cRepName, cRepTitle )
LOCAL i, nLen, nKey
LOCAL aDestList
LOCAL GetList := {}
LOCAL nSortSel := 1
LOCAL nDestSel := 1
LOCAL lRunReport := .F.

   cSubTitle := Space( 35 )
   cSortType := aSortList[1]
   aDestList := { "HTML (Browser)", "Screen", "CSV (Excel)" }
   cDestType := aDestList[1]
   nLen := Len( aCritList )

   SetColor( "W+/B,GR+/R,,,W+/B" )
   CLS

   // ---- Header / Footer ----
   @ 0, 0 SAY PadC( cRepTitle + " (" + cRepName + ")", 80 ) COLOR "GR+/RB"
   @ 24, 0 SAY PadR( " [Enter]=Select  [Space]=Toggle  [PgDn]=Run Report  [Esc]=Cancel", 80 ) COLOR "GR+/RB"

   // ---- Subtitle ----
   @ 2, 2 SAY "Optional subtitle:" COLOR "W+/B"

   // ---- Sort Rules (left column) ----
   @ 4, 2 SAY "Select sort rule:" COLOR "W+/B"
   FOR i := 1 TO Len( aSortList )
      @ 5 + i - 1, 4 SAY iif( i == nSortSel, "(o) ", "( ) " ) + aSortList[i] COLOR iif( i == nSortSel, "GR+/B", "W/B" )
   NEXT

   // ---- Destination (left column, below sort) ----
   @ 5 + Len(aSortList) + 1, 2 SAY "Select destination:" COLOR "W+/B"
   FOR i := 1 TO Len( aDestList )
      @ 5 + Len(aSortList) + 1 + i, 4 SAY iif( i == nDestSel, "(o) ", "( ) " ) + aDestList[i] COLOR iif( i == nDestSel, "GR+/B", "W/B" )
   NEXT

   // ---- Criteria checkboxes (right column) ----
   @ 2, 45 SAY "Select criteria:" COLOR "W+/B"
   FOR i := 1 TO nLen
      @ 2 + i, 45 SAY iif( aCritCheck[i], "[X] ", "[ ] " ) + aCritList[i] COLOR "W/B"
   NEXT

   // ---- Interactive loop ----
   // Section: 1=subtitle, 2=sort, 3=dest, 4=criteria
   LOCAL nSection := 2
   LOCAL nCritPos := 1
   LOCAL nRow

   SetCursor( SC_NONE )

   DO WHILE .T.

      // Highlight current position
      DO CASE
      CASE nSection == 1
         SetCursor( SC_NORMAL )
         @ 3, 2 GET cSubTitle PICTURE "@S35"
         READ
         SetCursor( SC_NONE )
         nSection := 2

      CASE nSection == 2
         // Draw sort with highlight
         FOR i := 1 TO Len( aSortList )
            IF i == nSortSel
               @ 5 + i - 1, 4 SAY "(o) " + PadR( aSortList[i], 38 ) COLOR "N/W"
            ELSE
               @ 5 + i - 1, 4 SAY "( ) " + PadR( aSortList[i], 38 ) COLOR "W/B"
            ENDIF
         NEXT

         nKey := Inkey( 0 )
         DO CASE
         CASE nKey == K_UP
            IF nSortSel > 1 ; nSortSel-- ; ENDIF
         CASE nKey == K_DOWN
            IF nSortSel < Len( aSortList ) ; nSortSel++ ; ELSE ; nSection := 3 ; ENDIF
         CASE nKey == K_ENTER .OR. nKey == K_SPACE
            cSortType := aSortList[ nSortSel ]
         CASE nKey == K_TAB .OR. nKey == K_RIGHT
            nSection := 4
         CASE nKey == K_PGDN
            lRunReport := .T. ; EXIT
         CASE nKey == K_ESC
            EXIT
         ENDCASE

      CASE nSection == 3
         // Draw destination with highlight
         FOR i := 1 TO Len( aDestList )
            nRow := 5 + Len(aSortList) + 1 + i
            IF i == nDestSel
               @ nRow, 4 SAY "(o) " + PadR( aDestList[i], 20 ) COLOR "N/W"
            ELSE
               @ nRow, 4 SAY "( ) " + PadR( aDestList[i], 20 ) COLOR "W/B"
            ENDIF
         NEXT

         nKey := Inkey( 0 )
         DO CASE
         CASE nKey == K_UP
            IF nDestSel > 1
               nDestSel--
            ELSE
               nSection := 2
            ENDIF
         CASE nKey == K_DOWN
            IF nDestSel < Len( aDestList ) ; nDestSel++ ; ENDIF
         CASE nKey == K_ENTER .OR. nKey == K_SPACE
            cDestType := aDestList[ nDestSel ]
         CASE nKey == K_TAB .OR. nKey == K_RIGHT
            nSection := 4
         CASE nKey == K_PGDN
            lRunReport := .T. ; EXIT
         CASE nKey == K_ESC
            EXIT
         ENDCASE

      CASE nSection == 4
         // Draw criteria with highlight
         FOR i := 1 TO nLen
            IF i == nCritPos
               @ 2 + i, 45 SAY iif( aCritCheck[i], "[X] ", "[ ] " ) + PadR( aCritList[i], 30 ) COLOR "N/W"
            ELSE
               @ 2 + i, 45 SAY iif( aCritCheck[i], "[X] ", "[ ] " ) + PadR( aCritList[i], 30 ) COLOR iif( aCritCheck[i], "GR+/B", "W/B" )
            ENDIF
         NEXT

         nKey := Inkey( 0 )
         DO CASE
         CASE nKey == K_UP
            IF nCritPos > 1 ; nCritPos-- ; ENDIF
         CASE nKey == K_DOWN
            IF nCritPos < nLen ; nCritPos++ ; ENDIF
         CASE nKey == K_SPACE
            aCritCheck[ nCritPos ] := ! aCritCheck[ nCritPos ]
         CASE nKey == K_ENTER
            // toggle + open criteria input dialog
            aCritCheck[ nCritPos ] := ! aCritCheck[ nCritPos ]
            IF aCritCheck[ nCritPos ]
               EditCriteria( nCritPos )
            ENDIF
         CASE nKey == K_TAB .OR. nKey == K_LEFT
            nSection := 2
         CASE nKey == K_PGDN
            lRunReport := .T. ; EXIT
         CASE nKey == K_ESC
            EXIT
         ENDCASE

      ENDCASE
   ENDDO

   // Save selections
   cSortType := aSortList[ nSortSel ]
   cDestType := aDestList[ nDestSel ]

RETURN lRunReport


// ============================================================
// CRITERIA INPUT DIALOGS
// ============================================================
FUNCTION EditCriteria( nCrit )
LOCAL cScr, GetList := {}
LOCAL dFrom, dTo, nFrom, nTo, cVal

   cScr := SaveScreen( 10, 20, 16, 60 )
   SetColor( "W+/BG,GR+/R" )
   @ 10, 20 CLEAR TO 16, 60
   DispBox( 10, 20, 16, 60, HB_B_SINGLE_UNI + " " )
   @ 10, 22 SAY " " + aCritList[nCrit] + " " COLOR "N/BG"

   DO CASE
   CASE nCrit == 1  // Scheduled finish date
      dFrom := CToD( "" )
      dTo   := CToD( "" )
      @ 12, 22 SAY "From:" ; @ 12, 28 GET dFrom
      @ 13, 22 SAY "To:  " ; @ 13, 28 GET dTo
      READ
      IF LastKey() != K_ESC
         aCritValues[ nCrit ] := { dFrom, dTo }
      ENDIF

   CASE nCrit == 2  // Min&Max Hours
      nFrom := 0 ; nTo := 24
      @ 12, 22 SAY "Min hours:" ; @ 12, 33 GET nFrom PICTURE "9999"
      @ 13, 22 SAY "Max hours:" ; @ 13, 33 GET nTo   PICTURE "9999"
      READ
      IF LastKey() != K_ESC
         aCritValues[ nCrit ] := { nFrom, nTo }
      ENDIF

   CASE nCrit >= 3 .AND. nCrit <= 9  // Status,Purpose,Priority,PType,PLine,Size,Value
      cVal := Space( 50 )
      @ 12, 22 SAY "Enter value(s):"
      @ 13, 22 GET cVal PICTURE "@S30"
      READ
      IF LastKey() != K_ESC
         aCritValues[ nCrit ] := AllTrim( cVal )
      ENDIF

   CASE nCrit == 12 // Workstation from..to
      cVal := "0000" ; LOCAL cVal2 := "9999"
      @ 12, 22 SAY "From:" ; @ 12, 28 GET cVal  PICTURE "XXXX"
      @ 13, 22 SAY "To:  " ; @ 13, 28 GET cVal2 PICTURE "XXXX"
      READ
      IF LastKey() != K_ESC
         aCritValues[ nCrit ] := { cVal, cVal2 }
      ENDIF

   CASE nCrit == 14 // Pc_code from..to
      cVal := "0000" ; LOCAL cVal3 := "9999"
      @ 12, 22 SAY "From:" ; @ 12, 28 GET cVal  PICTURE "XXXX"
      @ 13, 22 SAY "To:  " ; @ 13, 28 GET cVal3 PICTURE "XXXX"
      READ
      IF LastKey() != K_ESC
         aCritValues[ nCrit ] := { cVal, cVal3 }
      ENDIF

   CASE nCrit == 18 // Slack from..to
      nFrom := 0 ; nTo := 50
      @ 12, 22 SAY "From:" ; @ 12, 28 GET nFrom PICTURE "999"
      @ 13, 22 SAY "To:  " ; @ 13, 28 GET nTo   PICTURE "999"
      READ
      IF LastKey() != K_ESC
         aCritValues[ nCrit ] := { nFrom, nTo }
      ENDIF

   OTHERWISE
      // generic text input for other criteria
      cVal := Space( 50 )
      @ 12, 22 SAY "Value:"
      @ 13, 22 GET cVal PICTURE "@S30"
      READ
      IF LastKey() != K_ESC
         aCritValues[ nCrit ] := AllTrim( cVal )
      ENDIF

   ENDCASE

   RestScreen( 10, 20, 16, 60, cScr )
   SetColor( "W+/B,GR+/R,,,W+/B" )

RETURN NIL


// ============================================================
// DATABASE OPERATIONS
// ============================================================
FUNCTION GetDataPath()
   // Adjust this to your actual data location
   // For ADS: "\\server\share\data\"
   // For local: "G:\TEST\"
RETURN "G:\TEST\"


FUNCTION OpenDatabases( cPath )

   USE ( cPath + "D_LINE" ) ALIAS D_LINE NEW SHARED VIA "DBFCDX"
   IF NetErr()
      RETURN .F.
   ENDIF

   USE ( cPath + "C_PROC" ) ALIAS C_PROC NEW SHARED VIA "DBFCDX"
   IF NetErr()
      RETURN .F.
   ENDIF

   USE ( cPath + "C_PROC" ) ALIAS C_PROC1 NEW SHARED VIA "DBFCDX"
   IF NetErr()
      RETURN .F.
   ENDIF

   USE ( cPath + "M_LINEMV" ) ALIAS M_LINEMV NEW SHARED VIA "DBFCDX"
   IF NetErr()
      RETURN .F.
   ENDIF

RETURN .T.


FUNCTION CloseAllFiles()

   IF Select( "T_LINE01" ) > 0
      SELECT T_LINE01
      dbCloseArea()
   ENDIF
   IF Select( "D_LINE" ) > 0
      SELECT D_LINE
      dbCloseArea()
   ENDIF
   IF Select( "C_PROC" ) > 0
      SELECT C_PROC
      dbCloseArea()
   ENDIF
   IF Select( "C_PROC1" ) > 0
      SELECT C_PROC1
      dbCloseArea()
   ENDIF
   IF Select( "M_LINEMV" ) > 0
      SELECT M_LINEMV
      dbCloseArea()
   ENDIF

RETURN NIL


// ============================================================
// FILTER + CREATE TEMP DBF (same logic as TheReport:CreateCondDb)
// ============================================================
FUNCTION CreateFilteredDbf( cPath )
LOCAL cTempFile := hb_DirTemp() + "t_line01"

   SELECT D_LINE

   // Set relations
   SET RELATION TO FIELD->CPPROC_ID INTO C_PROC
   SET RELATION TO FIELD->PPPROC_ID INTO C_PROC1 ADDITIVE

   dbGoTop()

   // Copy with filter (ShaiCond equivalent)
   COPY TO ( cTempFile ) FOR EvalAllCriteria()

   // Open temp file
   USE ( cTempFile ) ALIAS T_LINE01 NEW EXCLUSIVE VIA "DBFCDX"
   cRepDbf := "T_LINE01"

RETURN NIL


FUNCTION EvalAllCriteria()
// Equivalent of ShaiCond(aTestBlocks) — evaluates all active criteria
LOCAL i

   FOR i := 1 TO Len( aCritCheck )
      IF aCritCheck[i] .AND. aCritValues[i] != NIL
         IF ! EvalOneCriterion( i )
            RETURN .F.
         ENDIF
      ENDIF
   NEXT

RETURN .T.


FUNCTION EvalOneCriterion( nCrit )
// Port of SetQueryBlocks from RPQC01V1.PRG (new variant)
LOCAL xVal := aCritValues[ nCrit ]

   DO CASE
   CASE nCrit == 1  // Scheduled finish date
      IF ! Empty( xVal[2] )
         RETURN D_LINE->B_DPROM <= xVal[2]
      ENDIF

   CASE nCrit == 2  // Hours at rep.stage
      RETURN HoursAt( "D_LINE" ) >= xVal[1] .AND. HoursAt( "D_LINE" ) <= xVal[2]

   CASE nCrit == 3  // Status
      RETURN D_LINE->B_STAT $ xVal

   CASE nCrit == 4  // Purpose
      RETURN D_LINE->B_PURP $ xVal

   CASE nCrit == 5  // Priority
      RETURN D_LINE->B_PRIOR $ xVal

   CASE nCrit == 6  // Product type
      RETURN D_LINE->PTYPE_ID $ xVal

   CASE nCrit == 7  // Product line
      RETURN D_LINE->PLINE_ID $ xVal

   CASE nCrit == 8  // Size
      RETURN D_LINE->SIZE_ID $ xVal

   CASE nCrit == 9  // Value
      RETURN Str( D_LINE->VALUE_ID, 8, 3 ) $ xVal

   CASE nCrit == 12 // Workstation
      RETURN D_LINE->CPWKSTN_ID >= xVal[1] .AND. D_LINE->CPWKSTN_ID <= xVal[2]

   CASE nCrit == 13 // Processes
      RETURN D_LINE->CPPROC_ID $ xVal

   CASE nCrit == 14 // Pc_code
      RETURN D_LINE->CP_PCCODE >= xVal[1] .AND. D_LINE->CP_PCCODE <= xVal[2]

   CASE nCrit == 15 // Remark
      RETURN xVal $ D_LINE->B_REMARK

   CASE nCrit == 16 // Dialecter
      RETURN D_LINE->DIEL_ID $ xVal

   CASE nCrit == 17 // Route filter
      RETURN AllTrim( xVal ) $ D_LINE->ROUTE_ID

   CASE nCrit == 18 // Slack from..to
      RETURN Slack_01( "D_LINE" ) >= xVal[1] .AND. Slack_01( "D_LINE" ) <= xVal[2]

   CASE nCrit == 19 // Reporting Stage
      RETURN RepStage( "D_LINE" ) $ xVal

   CASE nCrit == 20 // ESN(XX)
      RETURN D_LINE->ESNXX_ID $ xVal

   CASE nCrit == 21 // ESN(Y)
      RETURN D_LINE->ESNY_ID $ xVal

   ENDCASE

RETURN .T.


// ============================================================
// INDEX TEMP FILE
// ============================================================
FUNCTION IndexTempFile( nSort )
LOCAL cExpr

   SELECT T_LINE01
   cExpr := aSortExpr[ nSort ]

   IF ! Empty( cExpr )
      INDEX ON &cExpr TAG T_LINE01
   ENDIF

   dbGoTop()

RETURN NIL


// ============================================================
// HTML REPORT GENERATOR — replaces rpGenReport(RH2)
// ============================================================
FUNCTION GenHtmlReport( nSort )
LOCAL cFile := hb_DirTemp() + "rpqc01v4_report.html"
LOCAL hFile
LOCAL nLine := 0
LOCAL nTotalProm := 0, nTotalExp := 0
LOCAL cStage, nHours, nSlack, nProm, nExp

   hFile := FCreate( cFile )
   IF hFile == -1
      Alert( "Cannot create HTML file!" )
      RETURN ""
   ENDIF

   // ---- HTML Header ----
   FWrite( hFile, '<!DOCTYPE html>' + hb_eol() )
   FWrite( hFile, '<html><head><meta charset="utf-8">' + hb_eol() )
   FWrite( hFile, '<title>Production Line Report (PQC01V4)</title>' + hb_eol() )
   FWrite( hFile, '<style>' + hb_eol() )
   FWrite( hFile, 'body { font-family: Consolas, "Courier New", monospace; font-size: 11px; margin: 10px; }' + hb_eol() )
   FWrite( hFile, 'table { border-collapse: collapse; width: 100%; }' + hb_eol() )
   FWrite( hFile, 'th, td { border: 1px solid #333; padding: 2px 4px; text-align: left; white-space: nowrap; }' + hb_eol() )
   FWrite( hFile, 'th { background: #003366; color: white; position: sticky; top: 0; }' + hb_eol() )
   FWrite( hFile, 'tr:nth-child(even) { background: #f0f0f0; }' + hb_eol() )
   FWrite( hFile, 'tr:hover { background: #ffffcc; }' + hb_eol() )
   FWrite( hFile, '.header { background: #003366; color: white; padding: 8px; margin-bottom: 10px; }' + hb_eol() )
   FWrite( hFile, '.header h2 { margin: 0; }' + hb_eol() )
   FWrite( hFile, '.summary { margin-top: 10px; font-weight: bold; background: #e0e0e0; padding: 5px; }' + hb_eol() )
   FWrite( hFile, '.right { text-align: right; }' + hb_eol() )
   FWrite( hFile, '.center { text-align: center; }' + hb_eol() )
   FWrite( hFile, '@media print { th { background: #003366 !important; -webkit-print-color-adjust: exact; } }' + hb_eol() )
   FWrite( hFile, '</style></head><body>' + hb_eol() )

   // ---- Report Header ----
   FWrite( hFile, '<div class="header">' + hb_eol() )
   FWrite( hFile, '<h2>AVX Israel &mdash; Production Line Report</h2>' + hb_eol() )
   FWrite( hFile, '<div>Report: RPQC01V4 &nbsp;|&nbsp; Sort: ' + cSortType + ;
           ' &nbsp;|&nbsp; Date: ' + DToC( Date() ) + ' ' + Time() + ;
           ' &nbsp;|&nbsp; User: ' + hb_UserName() + '</div>' + hb_eol() )
   IF ! Empty( AllTrim( cSubTitle ) )
      FWrite( hFile, '<div>' + AllTrim( cSubTitle ) + '</div>' + hb_eol() )
   ENDIF
   FWrite( hFile, '</div>' + hb_eol() )

   // ---- Table Header ----
   FWrite( hFile, '<table>' + hb_eol() )
   FWrite( hFile, '<thead><tr>' + hb_eol() )
   FWrite( hFile, '<th>No.</th>' )
   FWrite( hFile, '<th>P</th>' )
   FWrite( hFile, '<th>B/N</th>' )
   FWrite( hFile, '<th>Size</th>' )
   FWrite( hFile, '<th>Value</th>' )
   FWrite( hFile, '<th>W</th>' )
   FWrite( hFile, '<th>S</th>' )
   FWrite( hFile, '<th>P</th>' )
   FWrite( hFile, '<th>Current Process</th>' )
   FWrite( hFile, '<th>Pr</th>' )
   FWrite( hFile, '<th>Fin.Date</th>' )
   FWrite( hFile, '<th>Rep.Stage</th>' )
   FWrite( hFile, '<th>Date</th>' )
   FWrite( hFile, '<th>Time</th>' )
   FWrite( hFile, '<th>Hours At</th>' )
   FWrite( hFile, '<th>Slack</th>' )
   FWrite( hFile, '<th>Promis.Qty</th>' )
   FWrite( hFile, '<th>Expect.Qty</th>' )
   FWrite( hFile, '<th>S</th>' )
   FWrite( hFile, '<th>Remark</th>' )
   FWrite( hFile, '<th>ESN</th>' )
   FWrite( hFile, '</tr></thead>' + hb_eol() )
   FWrite( hFile, '<tbody>' + hb_eol() )

   // ---- Data rows ----
   SELECT T_LINE01
   // Set relation from temp to C_PROC for current process name
   IF Select( "C_PROC" ) > 0
      SET RELATION TO FIELD->CPPROC_ID INTO C_PROC
   ENDIF
   IF Select( "C_PROC1" ) > 0
      SET RELATION TO FIELD->PPPROC_ID INTO C_PROC1 ADDITIVE
   ENDIF

   dbGoTop()
   DO WHILE ! Eof()
      nLine++

      cStage := RepStage( "T_LINE01" )
      nHours := HoursAt( "T_LINE01" )
      nSlack := Slack_01( "T_LINE01" )
      nProm  := QtyPromised( "T_LINE01" )
      nExp   := QtyExpected( "T_LINE01" )

      nTotalProm += nProm
      nTotalExp  += nExp

      FWrite( hFile, '<tr>' )
      FWrite( hFile, '<td class="right">' + hb_NToS( nLine ) + '</td>' )
      FWrite( hFile, '<td class="center">' + AllTrim( T_LINE01->B_PURP ) + '</td>' )
      FWrite( hFile, '<td>' + AllTrim( T_LINE01->B_ID ) + '</td>' )
      FWrite( hFile, '<td>' + AllTrim( T_LINE01->SIZE_ID ) + '</td>' )
      FWrite( hFile, '<td class="right">' + Str( T_LINE01->VALUE_ID, 8, 3 ) + '</td>' )
      FWrite( hFile, '<td class="right">' + Transform( T_LINE01->CP_BQTYW, "999" ) + StarMark( "W" ) + '</td>' )
      FWrite( hFile, '<td class="right">' + Transform( T_LINE01->CP_BQTYS, "99,999" ) + StarMark( "S" ) + '</td>' )
      FWrite( hFile, '<td class="right">' + Transform( T_LINE01->CP_BQTYP, "9,999,999" ) + StarMark( "P" ) + '</td>' )
      FWrite( hFile, '<td>' + CurrentProcNM( "T_LINE01" ) + '</td>' )
      FWrite( hFile, '<td class="center">' + AllTrim( T_LINE01->B_PRIOR ) + '</td>' )
      FWrite( hFile, '<td>' + DToC( T_LINE01->B_DPROM ) + '</td>' )
      FWrite( hFile, '<td class="center">' + cStage + '</td>' )
      FWrite( hFile, '<td>' + DToC( GetDate01( "T_LINE01" ) ) + '</td>' )
      FWrite( hFile, '<td>' + GetTime01( "T_LINE01" ) + '</td>' )
      FWrite( hFile, '<td class="right">' + Str( nHours, 8, 2 ) + '</td>' )
      FWrite( hFile, '<td class="right">' + Str( nSlack, 5 ) + '</td>' )
      FWrite( hFile, '<td class="right">' + Transform( nProm, "999,999" ) + '</td>' )
      FWrite( hFile, '<td class="right">' + Transform( nExp, "999,999" ) + '</td>' )
      FWrite( hFile, '<td class="center">' + AllTrim( T_LINE01->B_STAT ) + '</td>' )
      FWrite( hFile, '<td>' + AllTrim( T_LINE01->B_REMARK ) + '</td>' )
      FWrite( hFile, '<td>' + AllTrim( T_LINE01->ESN_ID ) + '</td>' )
      FWrite( hFile, '</tr>' + hb_eol() )

      dbSkip()
   ENDDO

   // ---- Summary ----
   FWrite( hFile, '</tbody></table>' + hb_eol() )
   FWrite( hFile, '<div class="summary">' + hb_eol() )
   FWrite( hFile, 'Total records: ' + hb_NToS( nLine ) )
   FWrite( hFile, ' &nbsp;|&nbsp; Total Promised Qty: ' + Transform( nTotalProm, "9,999,999" ) )
   FWrite( hFile, ' &nbsp;|&nbsp; Total Expected Qty: ' + Transform( nTotalExp, "9,999,999" ) )
   FWrite( hFile, '</div>' + hb_eol() )
   FWrite( hFile, '<div style="margin-top:5px;color:#666;">End Of Report: Production line report (PQC01V4)</div>' + hb_eol() )
   FWrite( hFile, '</body></html>' + hb_eol() )

   FClose( hFile )

RETURN cFile


// ============================================================
// SCREEN BROWSE (Query mode)
// ============================================================
FUNCTION BrowseReport()
LOCAL oBrw, nKey

   SELECT T_LINE01
   dbGoTop()

   SetColor( "W+/B,N/W,,,W+/B" )
   CLS
   @ 0, 0 SAY PadC( "Production Line Report — Browse (Esc to exit)", 80 ) COLOR "GR+/RB"

   oBrw := TBrowseDB( 1, 0, 23, 79 )
   oBrw:colorSpec := "W+/B,N/W,W+/R"

   oBrw:addColumn( TBColumnNew( "B/N",      {|| T_LINE01->B_ID } ) )
   oBrw:addColumn( TBColumnNew( "P",        {|| T_LINE01->B_PURP } ) )
   oBrw:addColumn( TBColumnNew( "Size",     {|| T_LINE01->SIZE_ID } ) )
   oBrw:addColumn( TBColumnNew( "Value",    {|| Str(T_LINE01->VALUE_ID,8,3) } ) )
   oBrw:addColumn( TBColumnNew( "Process",  {|| CurrentProcNM("T_LINE01") } ) )
   oBrw:addColumn( TBColumnNew( "Stage",    {|| RepStage("T_LINE01") } ) )
   oBrw:addColumn( TBColumnNew( "Hours",    {|| Str(HoursAt("T_LINE01"),7,2) } ) )
   oBrw:addColumn( TBColumnNew( "Slack",    {|| Str(Slack_01("T_LINE01"),5) } ) )
   oBrw:addColumn( TBColumnNew( "Fin.Date", {|| DToC(T_LINE01->B_DPROM) } ) )
   oBrw:addColumn( TBColumnNew( "Status",   {|| T_LINE01->B_STAT } ) )
   oBrw:addColumn( TBColumnNew( "ESN",      {|| T_LINE01->ESN_ID } ) )

   DO WHILE .T.
      DO WHILE ! oBrw:stabilize() ; ENDDO
      nKey := Inkey( 0 )
      DO CASE
      CASE nKey == K_ESC    ; EXIT
      CASE nKey == K_UP     ; oBrw:up()
      CASE nKey == K_DOWN   ; oBrw:down()
      CASE nKey == K_PGUP   ; oBrw:pageUp()
      CASE nKey == K_PGDN   ; oBrw:pageDown()
      CASE nKey == K_HOME   ; oBrw:goTop()
      CASE nKey == K_END    ; oBrw:goBottom()
      CASE nKey == K_LEFT   ; oBrw:left()
      CASE nKey == K_RIGHT  ; oBrw:right()
      ENDCASE
   ENDDO

RETURN NIL


// ============================================================
// CSV EXPORT
// ============================================================
FUNCTION GenCsvExport()
LOCAL cFile := hb_DirTemp() + "rpqc01v4.csv"

   SELECT T_LINE01
   COPY TO ( cFile ) DELIMITED

   Alert( "CSV exported to: " + cFile )
   hb_run( 'start "" "' + cFile + '"' )

RETURN NIL


// ============================================================
// UDF FUNCTIONS — ported from RPQC01V1.PRG
// ============================================================

FUNCTION RepStage( cAlias )
// F=Finished arriving at proc, A=Arrived, S=Started, ?=unknown
LOCAL cRetVal

   IF Empty( (cAlias)->CP_DARR ) .AND. Empty( (cAlias)->CP_DSTA ) .AND. ! Empty( (cAlias)->PP_DFIN )
      cRetVal := "F"
   ELSEIF ! Empty( (cAlias)->CP_DARR ) .AND. Empty( (cAlias)->CP_DSTA ) .AND. ! Empty( (cAlias)->PP_DFIN )
      cRetVal := "A"
   ELSEIF ! Empty( (cAlias)->CP_DARR ) .AND. ! Empty( (cAlias)->CP_DSTA ) .AND. ! Empty( (cAlias)->PP_DFIN )
      cRetVal := "S"
   ELSE
      cRetVal := "?"
   ENDIF

RETURN cRetVal


FUNCTION HoursAt( cAlias )
// Hours at current reporting stage
LOCAL nRetVal
LOCAL cStage := RepStage( cAlias )

   DO CASE
   CASE cStage == "F"
      nRetVal := CalcHours( (cAlias)->PP_DFIN, (cAlias)->PP_TFIN, Date(), Time() )
   CASE cStage == "A"
      nRetVal := CalcHours( (cAlias)->CP_DARR, (cAlias)->CP_TARR, Date(), Time() )
   CASE cStage == "S"
      nRetVal := CalcHours( (cAlias)->CP_DSTA, (cAlias)->CP_TSTA, Date(), Time() )
   OTHERWISE
      nRetVal := 999
   ENDCASE

RETURN nRetVal


FUNCTION CalcHours( dFirst, cTimeFirst, dSecond, cTimeSecond )
// Port of Hours() static function from original
LOCAL nHours, nMin1, nMin2, nMinDiff

   nHours := ( dSecond - dFirst ) * 24

   nMin1 := Val( SubStr( cTimeSecond, 1, 2 ) ) * 60 + Val( SubStr( cTimeSecond, 4, 2 ) )
   nMin2 := Val( SubStr( cTimeFirst, 1, 2 ) )  * 60 + Val( SubStr( cTimeFirst, 4, 2 ) )

   nMinDiff := nMin1 - nMin2
   nHours := nHours + Int( nMinDiff / 60 ) + ( ( nMinDiff % 60 ) / 60 / 1.66 )

RETURN nHours


FUNCTION Slack_01( cAlias )
RETURN (cAlias)->B_DPROM - (cAlias)->EXPFINDATE


FUNCTION GetDate01( cAlias )
LOCAL cStage := RepStage( cAlias )

   DO CASE
   CASE cStage == "F" ; RETURN (cAlias)->PP_DFIN
   CASE cStage == "A" ; RETURN (cAlias)->CP_DARR
   CASE cStage == "S" ; RETURN (cAlias)->CP_DSTA
   ENDCASE

RETURN CToD( "" )


FUNCTION GetTime01( cAlias )
LOCAL cStage := RepStage( cAlias )

   DO CASE
   CASE cStage == "F" ; RETURN (cAlias)->PP_TFIN
   CASE cStage == "A" ; RETURN (cAlias)->CP_TARR
   CASE cStage == "S" ; RETURN (cAlias)->CP_TSTA
   ENDCASE

RETURN "  :  "


FUNCTION CurrentProcNM( cAlias )
// Current process name: proc_id + ":" + process name
   C_PROC->( dbSeek( (cAlias)->CPPROC_ID ) )
RETURN AllTrim( (cAlias)->CPPROC_ID ) + ":" + AllTrim( C_PROC->PROC_NME )


FUNCTION StarMark( cType )
// Returns "*" if the process does NOT require this qty type
   DO CASE
   CASE cType == "W" ; RETURN iif( C_PROC->MUST_W, " ", "*" )
   CASE cType == "S" ; RETURN iif( C_PROC->MUST_S, " ", "*" )
   CASE cType == "P" ; RETURN iif( C_PROC->MUST_P, " ", "*" )
   ENDCASE
RETURN " "


FUNCTION QtyPromised( cAlias )
// Sum of promised quantities from D_PROM for this batch
// Simplified version — in full port needs D_PROM table
LOCAL nPromQty := 0
   // TODO: port full logic with D_PROM, D_ORD, D_ORDREQ tables
   // For now return 0 — needs D_PROM table access
RETURN nPromQty


FUNCTION QtyExpected( cAlias )
// Expected quantity calculation
// Simplified version — in full port needs c_expqty, c_thqty tables
LOCAL nQty := 0
   // TODO: port AvailbleForAnyOne() logic
   // For now return 0 — needs additional tables
RETURN nQty

// ============================================================
// END OF FILE
// ============================================================
