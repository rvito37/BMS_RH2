STATIC cFile
STATIC KyocerHome := '!R! UNIT C;RGST 0.5,0.3; EXIT; '

#include "avxdefs.ch"

Function BroPrintOut()

LOCAL oBro
LOCAL nKey,i,lFilter := FALSE
LOCAL cPrnDir  := GetUserInfo():cPrnDir
LOCAL aFiles   := {},cSaveLine
LOCAL cSearcher := "",bKey
LOCAL oForm    := Form():new(,,,,"Stock Printouts" ,, ProcName() )
LOCAL aKeyList := {{ "F3  - Print                       ",K_F3   } ,;
	                { "F2  - Set/Clear filter            ",K_F2   } ,;
	                { "TAB - Reprint/New Printout Window ",K_TAB  } ,;
	                { "DEL - Move to reprints file       ",K_DEL  }  }

SET DELETE ON
cFile := "d_prnout"
aFiles := GenOpenFiles({"h_prnout",cFile})
h_prnout->(ordsetfocus(1))
(cFile)->(ordsetfocus(1))
(cFile)->(dbgotop())
IF d_prnout->(lastRec()) == 0
	Alert("No records found",{"Ok"})
   GenCloseFiles(aFiles)
   oForm:hide()
	Return TRUE
ENDIF

@ 2, 65 SAY "New Printouts"
@ 2, 30 SAY "MVT. Type+Date+Time"
cSaveLine := SaveScreen( 2,0,2,79 )
@ 2,0 SAY cSearcher COLOR "n/w"

oBro := (cFile)->( TBrowseDB( 3, 0, 23 , 79 ) )
oBro:headSep   := "ÍÑÍ"
oBro:colSep    := " ³ "
oBro:footSep   := "ÍÏÍ"
oBro:colorSpec := Colors()


oBro:addColumn( TBColumnNEW( " "         , {|| IIF((cFile)->PRN_Y_N ,"û"," ")} ) )
oBro:addColumn( TBColumnNEW( "Mvt.;Type" , {|| (cFile)->SMTYPE_ID           } ) )
oBro:addColumn( TBColumnNEW( "Ref #"     , {|| IIF(!Empty((cFile)->SMTYPE_ID),SubStr((cFile)->file_name,3) ,(cFile)->file_name ) } ) )
oBro:addColumn( TBColumnNEW( "Date"      , {|| DTOC((cFile)->DADD_REC)      } ) )
oBro:addColumn( TBColumnNEW( "Time"      , {|| (cFile)->TADD_REC            } ) )
oBro:addColumn( TBColumnNEW( "File Name" , {|| (cFile)->file_name           } ) )

oBro:gotopBlock := {|| (cFile)->(Top(cSearcher))}
oBro:gobottomBlock := {|| (cFile)->(Bot(cSearcher))}
oBro:skipBlock := {|n| (cFile)->(SkipIt(n,cSearcher,bKey))}
bKey := COMPILE( (cFile)->( ordKey() ) )

WHILE .T.


  WHILE !oBro:stabilize() ; END

  nKey = InKey(1)

  ASK FOR HELP

  IF nKey = K_ESC
     EXIT
  ELSEIF nKey = K_TAB
     if cFile == "h_prnout"
		  cFile := "d_prnout"
		  @ 2, 65 SAY "New Printouts"
	  else
        cFile := "h_prnout"
		  @ 2, 65 SAY "  Reprints   "
	  endif
     oBro:gotopBlock := {|| (cFile)->(Top(cSearcher))}
     oBro:gobottomBlock := {|| (cFile)->(Bot(cSearcher))}
     oBro:skipBlock := {|n| (cFile)->(SkipIt(n,cSearcher,bKey))}
     bKey := COMPILE( (cFile)->( ordKey() ) )
	  dbselectarea(cFile)
	  oBro:RefreshAll()
  ELSEIF nKey = K_F3 .AND. (cFile)->( RecLock(5,"",.F.) )
	  (cFile)->PRN_Y_N := TRUE
	  (cFile)->( DBUNLOCK() )
	  TONE(800,0)
	  @ 24, 0 SAY " Printing, please wait...          " COLOR "gr+/rb"
	  run("echo "+KyocerHome+" >LPT1")
	  COPY FILE &(cPrnDir+AllTrim((cFile)->file_name)+".txt") TO LPT1
	  @ 24, 0 SAY "                                   "
	  oBro:RefreshAll()
  ELSEIF nKey = K_F2
	  IF lFilter
        (cFile)->(dbclearfilter())
		  @ 24, 0 SAY "                                   "
	  ELSE
        (cFile)->(dbsetfilter({|| !(cFile)->PRN_Y_N}))
		  @ 24, 0 SAY " Filtered,only not printed records " COLOR "gr+/rb"
	  ENDIF
	  lFilter := !lFilter
	  oBro:gotop()
	  oBro:RefreshAll()
  ELSEIF nKey = K_DEL
     if cFile == "h_prnout" .AND. !(cFile)->(deleted()) .AND. h_prnout->( RecLock(5,"",.F.) )
		  if Alert("This printout will permanently be deleted",{"Yes","No"}) == 1
            ferase(cPrnDir+AllTrim((cFile)->file_name)+".txt")
			   (cFile)->( DBDELETE() )
				(cFile)->( DBCOMMIT())
			   (cFile)->( DBUNLOCK() )
			   (cFile)->( DBSKIP(-1) )
		  else
			   (cFile)->( DBUNLOCK() )
		  endif
        oBro:RefreshAll()
	  elseif cFile == "d_prnout" .AND. !(cFile)->(deleted()) .AND. h_prnout->( AddRec(5) )
         FOR i := 1 TO h_prnout->( FCOUNT() )
             h_prnout->( FIELDPUT( i, (cFile)->( FIELDGET(i) ) ) )
         NEXT
			h_prnout->( DBUNLOCK() )
			IF (cFile)->( RecLock(5,"",.F.) )
			   (cFile)->( DBDELETE() )
				(cFile)->( DBCOMMIT())
			   (cFile)->( DBUNLOCK() )
			   (cFile)->( DBSKIP(-1) )
			END
         oBro:RefreshAll()
	  endif
  ELSEIF nKey >= 32 .AND. nKey <= 255
       cSearcher += Upper( Chr( nKey ) )
       IF (cFile)->( DbSeek(cSearcher,.T.) )
       ELSE
            Alert("!„€‰‚™;;!! „† …‡š š…˜…™ ‰€",{"˜…™‰€"})
            cSearcher := SubStr(cSearcher,1,Len(cSearcher)-1)
            (cFile)->(DbSeek(cSearcher,.T.))
       END
       RestScreen(2,0,2,79,cSaveLine)
       @ 2,0 SAY cSearcher COLOR "n/w"
       oBro:RefreshAll()
  ELSEIF nKey = K_BS
       IF !Empty(Len(cSearcher))
            cSearcher := SubStr(cSearcher,1,Len(cSearcher)-1)
            (cFile)->( DbSeek(cSearcher,.T.) )
            oBro:RefreshAll()
       END
       RestScreen(2,0,2,79,cSaveLine)
       @ 2,0 SAY cSearcher COLOR "n/w"
  ELSEIF StdKeys( nKey , oBro )
  ENDIF
END

GenCloseFiles(aFiles)
@ 24, 0 SAY "                                   "
oForm:hide()
Return TRUE