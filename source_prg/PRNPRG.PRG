// prnprg.prg
// Function/Procedure Prototype Table  -  Last Update: 24-06-96 @ 14:59:42
// 컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴
// Return Value         Function/Arguments
// 컴컴컴컴컴컴컴컴컴�  컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴�

*---- PRNPRG.PRG : Temporary & Partial Printing Driver for AVXPR
*---- AVX_SERVER-MARKETING queue, RPRINTER 3 (Carlene's WS)
*---- avx_server-PROD_FLOOR queue, RPRINTER 4 (Shay's WS)
*---- Author: Dmitri Dain

*---- LAST UPDATE: 21/7/94


para xprinter,xfont

*---- HP DeskJet 500 ------------------------------------------------
DO CASE
  CASE upper(XPRINTER)="HPD500"
  _plength=63
  ! capture>nul
  do case
     case xfont=1        && Fix Narkis 16.6/8 (Cartridge ISR1/ISR2)
          ??? chr(27)+"(15Q"+chr(27)+"(s0p8v16.6h0s0b7t"
     case xfont=2        && Letter Gothic 12/12
          ??? chr(27)+"(10U"+chr(27)+"(s0p12h12v0s0b6t"
     case xfont=3        && Portrait Courier 20/12 (Should be a DEFAULT)
          ??? chr(27)+"&l0O"+chr(27)+"(10U"+chr(27)+"(s0p20h12v0s0b3t"
     case xfont=4        && For Foreign Invoices/Proformas
          ??? chr(27)+"(10U"+chr(27)+"(s0p12h12v0s0b(s6T"
     case xfont=5        && For Israeli Invoices (Cartridge ISR1/ISR2)
          ??? chr(27)+"(15Q"+chr(27)+"&l26A"+chr(27)+"(s0p10v12h0s0b7T"
     case xfont=6        && For Condensed LandScape (Cartridge ISR1/ISR2)
          HPLAND=chr(27)+"&l1O"
          HPCONDENSED=chr(27)+"(15Q"+chr(27)+"(s0p8v16.6h0s0b7T"
          _plength=43
          ??? HPLAND
          ??? HPCONDENSED
  endcase

*---- EPSON FX (Assumption: local) ---------------------------------
  CASE upper(XPRINTER)="EPS850"
  ! endcap
  _plength=66
  do case
     case xfont=3        && EPSON FX-850 Condensed Mode
          ??? chr(27)+chr(15)
  endcase


*---- HP LaserJet III -----------------------------------------------
  CASE upper(XPRINTER)="HP3"

       _plength=63
       ! capture s=avx_server q=hplaser>NUL
       do case

          case xfont=0   && Reset HP III, Go back to the native queue
               ??? chr(27)+"E"
               _pdriver="generic.pr2"
               ! capture>NUL

          case xfont=1   && Very small font with Hebrew
               && used for Net Reqs,Productivity
               ??? chr(27)+"(15Q"+chr(27)+"(s0p8v16.6h0s0b7t"


          case xfont=3        && For Condensed LandScape
               HPLAND=chr(27)+"&l1O"
               HPCONDENSED=chr(27)+"(15Q"+chr(27)+"(s0p8v16.6h0s0b7T"
               _plength=43
               ??? HPLAND
               ??? HPCONDENSED

       endcase


*---- KYOCERA FS-850  [Prnt_rtk.prg takes care of formatting !!]-----
  CASE upper(XPRINTER)="KYO"  && Formatting done in KYO own mode
       kyoreset="!R! RES ; EXIT ; "
       kyoheb17cpi="!R! FONT 8421 ; EXIT; "
       kyoheb12cpi="!R! FONT 8411 ; EXIT; "
       kyohebLAND=chr(27)+"&l1O"
       kyohebL17=chr(27)+"(15Q"+chr(27)+"(s0p8v16.6h0s0b7T"
       kyoPORT = "!R! SPOP; EXIT ; "

       do case
          case xfont=4   && Prod.Control 17cpi landscape reports
               _plength=43
               ! capture s=avx_server q=prod_floor>nul
               ??? kyoreset
               ??? kyohebLAND
               ??? kyohebL17

          case xfont=3   && Prod.Control 12 cpi Reports
               _plength=62
               ! capture s=avx_server q=prod_floor>nul
               ??? kyoreset
               ??? kyoheb12cpi
          case xfont=2   && Prod.Control 17 cpi Reports
               _plength=62
               ! capture s=avx_server q=prod_floor>nul
               ??? kyoreset
               ??? kyoheb17cpi
          case xfont=1   && Redirect to Kyocera's queue
               ! capture s=avx_server q=prod_floor>nul
          case xfont=0   && Reset to HP III emulation mode
               ! capture s=avx_server q=prod_floor>nul
               ??? kyoreset
               ??? kyoPORT
               ! capture>nul   && Back to the native queue
       endcase
ENDCASE

retu