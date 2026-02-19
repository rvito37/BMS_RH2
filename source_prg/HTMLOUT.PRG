// htmlout.prg
// Convert TVR text file output to HTML
//
// After rpGenReport() writes a text file via rpPRINTER_FILE,
// TextToHtml() reads it and creates an HTML file with the same
// content wrapped in <pre> tags.  Form-feed characters (chr 12)
// become visual page breaks.
//
// No changes to PAGEOUT.PRG or any TVR library files required.

#include "avxdefs.ch"

/*========================================================================*/
FUNCTION TextToHtml( cTextFile, cHtmlFile )
/*
 *  Read the text file produced by rpPRINTER_FILE and write an
 *  HTML file with identical visual layout.
 *
 *  cTextFile - full path to the text file (no extension)
 *  cHtmlFile - full path for the output .html file
 *
 *  Returns .T. on success, .F. on failure
 */
LOCAL nIn, nOut
LOCAL cText, cLine
LOCAL cHeader, cFooter
LOCAL nPos, nLen
LOCAL nCR, nFF, cOut

nIn := FOPEN( cTextFile, 0 )     // read-only
IF nIn <= 0
   RETURN .F.
ENDIF

nOut := FCREATE( cHtmlFile )
IF nOut <= 0
   FCLOSE( nIn )
   RETURN .F.
ENDIF

cHeader := '<!DOCTYPE html>'                                      + Chr(13)+Chr(10) +;
           '<html>'                                               + Chr(13)+Chr(10) +;
           '<head>'                                               + Chr(13)+Chr(10) +;
           '<meta charset="windows-1255">'                        + Chr(13)+Chr(10) +;
           '<style>'                                              + Chr(13)+Chr(10) +;
           'body { margin: 10px; background: #fff; }'             + Chr(13)+Chr(10) +;
           'pre  { font-family: "Courier New", monospace;'        + Chr(13)+Chr(10) +;
           '       font-size: 10pt; line-height: 1.2;'           + Chr(13)+Chr(10) +;
           '       white-space: pre; margin: 0; }'                + Chr(13)+Chr(10) +;
           '.page-break { border-top: 1px dashed #999;'          + Chr(13)+Chr(10) +;
           '              margin: 8px 0; }'                       + Chr(13)+Chr(10) +;
           '@media print {'                                       + Chr(13)+Chr(10) +;
           '  .page-break { page-break-before: always;'          + Chr(13)+Chr(10) +;
           '                border: none; margin: 0; }'          + Chr(13)+Chr(10) +;
           '}'                                                    + Chr(13)+Chr(10) +;
           '</style>'                                             + Chr(13)+Chr(10) +;
           '</head>'                                              + Chr(13)+Chr(10) +;
           '<body>'                                               + Chr(13)+Chr(10) +;
           '<pre>'                                                + Chr(13)+Chr(10)

FWRITE( nOut, cHeader )

* read entire text file
cText := space( 65535 )
nLen  := 0
cLine := ""
DO WHILE .T.
   nLen := FREAD( nIn, @cText, 65535 )
   IF nLen <= 0
      EXIT
   ENDIF
   cLine := cLine + Left( cText, nLen )
ENDDO
FCLOSE( nIn )

* process line by line
nPos := 1
DO WHILE nPos <= Len( cLine )

   * check for form-feed (page break)
   nFF := At( Chr(12), SubStr( cLine, nPos ) )
   * check for line end
   nCR := At( Chr(13), SubStr( cLine, nPos ) )

   IF nFF > 0 .AND. ( nCR == 0 .OR. nFF < nCR )
      * form-feed before next CR - output page break
      IF nFF > 1
         cOut := SubStr( cLine, nPos, nFF - 1 )
         FWRITE( nOut, HtmlEsc( cOut ) + Chr(13)+Chr(10) )
      ENDIF
      FWRITE( nOut, '</pre>' + Chr(13)+Chr(10) +;
                     '<div class="page-break"></div>' + Chr(13)+Chr(10) +;
                     '<pre>' + Chr(13)+Chr(10) )
      nPos := nPos + nFF
   ELSEIF nCR > 0
      * normal line ending with CR/LF
      cOut := SubStr( cLine, nPos, nCR - 1 )
      FWRITE( nOut, HtmlEsc( cOut ) + Chr(13)+Chr(10) )
      nPos := nPos + nCR
      * skip LF after CR
      IF nPos <= Len( cLine ) .AND. SubStr( cLine, nPos, 1 ) == Chr(10)
         nPos := nPos + 1
      ENDIF
   ELSE
      * last line, no CR
      cOut := SubStr( cLine, nPos )
      IF !Empty( cOut )
         FWRITE( nOut, HtmlEsc( cOut ) + Chr(13)+Chr(10) )
      ENDIF
      EXIT
   ENDIF
ENDDO

cFooter := '</pre>'    + Chr(13)+Chr(10) +;
           '</body>'   + Chr(13)+Chr(10) +;
           '</html>'   + Chr(13)+Chr(10)

FWRITE( nOut, cFooter )
FCLOSE( nOut )

RETURN .T.
/*========================================================================*/

/*========================================================================*/
STATIC FUNCTION HtmlEsc( cStr )
/*
 *  Escape HTML special characters for display inside <pre> block
 */
cStr := STRTRAN( cStr, "&" , "&amp;"  )
cStr := STRTRAN( cStr, "<" , "&lt;"   )
cStr := STRTRAN( cStr, ">" , "&gt;"   )
cStr := STRTRAN( cStr, '"' , "&quot;" )
RETURN cStr
/*========================================================================*/
