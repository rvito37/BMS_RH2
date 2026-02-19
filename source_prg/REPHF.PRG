// rephf.prg
// Function/Procedure Prototype Table  -  Last Update: 24-06-96 @ 15:01:09
// ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ
// Return Value         Function/Arguments
// ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ  ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ
// self                 METHOD footer
// self                 METHOD header
// self                 METHOD init( cTitle , cRepCode ,nLen )
// nLen                 METHOD print
// self                 METHOD setPage( nPage )

//G:\BMS\SOURCE\REPHF.PRG
// creates headers footers for printing


#include "avxdefs.ch"

CREATE CLASS RepHF
VISIBLE:
       VAR nLength
       VAR nPadLen
       VAR cTitle
       VAR aPrintLines
       VAR cUserInfoLine
       VAR cRepCode
       VAR nPageNoPos

       METHOD init
       METHOD header , setPage
       METHOD footer
       METHOD print

END CLASS

METHOD init( cTitle , cRepCode ,nLen )
       ::cTitle        := cTitle
       ::cRepCode      := cRepCode
       ::nLength       := nLen
       ::cUserInfoLine := SetRepTitles( ::nLength )
       ::aPrintLines   := {}

RETURN self

METHOD header
LOCAL nPageCodeLen := 10 + Len( ::cRepCode ) + 11

::nPadLen := ::nLength - nPageCodeLen - 1

Aadd( ::aPrintLines , "Ú" + Replicate("Ä",::nLength ) + "¿" )
Aadd( ::aPrintLines , "³" + ::cUserInfoLine           + "³" )
Aadd( ::aPrintLines , "³" + "Report No. "+::cRepCode + Padc( ::cTitle ,::nPadLen) + "Page #:    " + "³" )
Aadd( ::aPrintLines , "³" + Padc(Alltrim(prnGetUserMsg()), ::nLength )+"³" )
Aadd( ::aPrintLines , "À" + Replicate("Ä",::nLength)  + "Ù" )
::nPageNoPos := At( "#" , ::aPrintLines[3] ) + 1
RETURN self

METHOD setPage( nPage )
::aPrintLines[3] := Stuff(::aPrintLines[3], ::nPageNoPos , 4 , Str( nPage,4) )
RETURN self

METHOD footer
LOCAL cLine := "³" + "Report No. "+::cRepCode + Padc( "END OF " +::cTitle ,::nPadLen) + Space( 11 )+ "³"
::aPrintLines[3] := cLine

Adel( ::aPrintLines , Len( ::aPrintLines ) - 1 )
Asize( ::aPrintLines , Len( ::aPrintLines ) - 1 )

RETURN self

METHOD print

LOCAL nLen := Len( ::aPrintLine ) , i

FOR i := 1 TO nLen
    ? ::aPrintLines[i]
NEXT

RETURN nLen
