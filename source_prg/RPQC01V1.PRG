* Rpqc01v1.prg

#include "avxdefs.ch"

STATIC oMyRep,cFileName

Function MyRpqc01v1(ReportFunc)
LOCAL oRep
MEMVAR aBuffer

oRep := TheReport():New(ReportFunc,"Production line report")

IF ReportFunc == "RPQC01V1" //old "classic" report VR 27-01-02 TID:2127212
oRep:SetSort({"Proc(D)+Prior+B/N" ,;
              "Prior+Slack"       ,;//"B/N",;
              "Btype+Size+Value+Proc(D)" ,;
              "Proc(D)+Btype+B/N"  ,;
              "Btype+Size+Proc(D)" ,;
              "TopDown+ExpFin+Stat+Purp+Proc(D)-Tot.",;//Type+Line+Size+Value+Proc(D)-Totals EDOMO 00122717 25-06-01
              "Type+Line+Size+Proc(D)+Value",;
              "Wrkst+Proc(D)+Btype - FF" ,;  // using aReports this option will call a different rh2
              "Wrkst+Proc(D)+Btype" ,;
              "Processes+Type+Line+Size+Value",;
              "Proc(D)+Prior+Slack+Purp+Stat+B/N",;//id: 37281079
				  "TopDown+Proc(D)";
             } )

oRep:SetExprSort({"Str(10000-val(cp_pccode),5)+b_prior+GetNewBidSort()" ,;
                   "B_Prior+STR(100000+Slack_01(Alias()))",;//"GetNewBidSort()",;
                   "b_type+size_id+Str(value_id,8,3)+Str(10000-val(cp_pccode),5)" ,;
                   "Str(10000-val(cp_pccode),5)+b_type+GetNewBidSort()",;
                   "b_type+size_id+Str(10000-val(cp_pccode),5)",;
                   "ptype_id+pline_id+size_id+str(value_id,8,3) + DTOS(ExpFinDate) + Getseqbstat(b_stat) + Getseqpurp(b_purp) + Str(10000-val(cp_pccode),5)",;//VR 00122717 25-06-01
                   "ptype_id+pline_id+size_id+Str(10000-val(cp_pccode),5)+str(value_id,8,3)",;
                   "cpwkstn_id+Str(10000-val(cp_pccode),5)+b_type",;
                   "cpwkstn_id+Str(10000-val(cp_pccode),5)+b_type",;
                   "cpproc_id+ptype_id+pline_id+size_id+str(value_id,8,3)",;
						 "Str(10000-val(cp_pccode),5)+B_Prior+STR(100000+Slack_01(Alias()))+Getseqpurp(b_purp)+Getseqbstat(b_stat)+ B_ID",;
						 "ptype_id+pline_id+size_id+str(value_id,8,3) + cpproc_id";
                    } )
oRep:SetCrit({       "Scheduled finish date"     ,;
                     "No of days at current proc",;
                     "Status"                    ,;
                     "Purpose"                   ,;
                     "Priority"                  ,;
                     "Product type"              ,;
                     "Product line"              ,;
                     "Size"                      ,;
                     "Value"                     ,;
                     "Promised Batches"          ,;
                     "P/O Type"                  ,;
                     "Workstation"               ,;
                     "Processes"                 ,;
                     "Pc_code"                   ,;
                     "Remark"                    ,;
                     "Pack instruction's"        ,;
                     "Route filter"              ,;
                     "Slack from...to"           ,;
                     "Pc_code and date"          ,;
                     "ESN(XX)"                   ,;
                     "ESN(Y)"                     ;
                   })

oRep:SetCheck({     {|o| GetFinishDate(o,{Ctod("  /  /  "),Ctod("  /  /  ")}) } ,;
                    {|o| GetNoOfDays(o,{0})}    ,;
                    {|o| critBrowse( o ,{"all"},"c_bstat",    , {|| c_bstat->b_stat } ,{|| c_bstat->bstat_nme } )},;
                    {|o| critBrowse( o ,{"all"},"c_bpurp",    , {|| c_bpurp->b_purp } ,{|| c_bpurp->bpurp_nme } )},;
                    {|o| critBrowse( o ,{"all"},"c_prior",    , {|| c_prior->b_prior },{|| c_prior->bprior_nm } )},;
                    {|o| critBrowse( o ,{"all"},"c_ptype",    , {|| c_ptype->ptype_id},{|| c_ptype->ptype_nm }  )},;
                    {|o| critBrowse( o ,{"all"},"c_pline",    , {|| c_pline->pline_id},{|| c_pline->pline_nm }  )},;
                    {|o| critBrowse( o ,{"all"},"c_size" ,    , {||  c_size->size_id },{|| Space(60) }          )},;
                    {|o| critBrowse( o ,{"all"},"c_value",    , {|| Str(c_value->value_id,8,3) } ,{|| c_value->value_nm } )},;
                    {|o| GetQtyChoice(o,{"No"}) } ,;
                    {|o| critBrowse( o ,{"all"},"c_potype",, {|| c_potype->poln_type }   ,{|| c_potype->potype_nm } )   },;
                    {|o| Get42Wkstn(o,{"0000","9999"})}    ,;
                    {|o| critBrowse( o ,{"all"},"c_proc" ,    , {|| c_proc->proc_id  }  ,{|| c_proc->proc_nme } )},;
                    {|o| Getpccode2(o,{"0000","9999"})}   ,;
                    {|o| GetRemark(o,{space(50)})}       ,;
                    {|o| GetQtyChoice(o,{"Yes"},{"With instruction's","Without instruction's"},"Pack instruction's") } ,;
                    {|o| GetRoute(o,{space(4)}) }        ,;
                    {|o| GetSlack(o,{0,50}) }            ,;
                    {|o| GetPcAndDate(o,{"0000","9999",ctod("  /  /  "),ctod("  /  /  ")})}   ,;
                    {|o| critBrowse( o ,{"all"},"c_esnxx",    , {|| c_esnxx->esnxx_id } ,{|| c_esnxx->esnxx_nm } )},;
                    {|o| critBrowse( o ,{"all"},"c_esny" ,    , {|| c_esny->esny_id }   ,{|| c_esny->esny_nm } )} ;
              })
oRep:SetBuffer({ "Scheduled finish date", "No of days at current proc",;
                  "c_bstat","c_bpurp","c_prior","c_ptype","c_pline",;
                  "c_size","c_value","Promised Batches","P/O Type",;
                  "Workstation","Processes","Pc_code","Remark",;
                  "Pack instruction's","Route filter","Slack",;
                  "Pc_code and date","ESN(XX)","ESN(Y)"})

oRep:SetQueryBlocks({ {|| d_line->b_dprom  <= aBuffer[ 1]   },;
             {|| D_line->(date()-d_line->pp_dfin) >= aBuffer[ 2] },;        //d_line->(Lupdate())-d_line->pp_dfin    changed by
             {|| D_line->b_STAT   $ aBuffer[ 3]   },;
             {|| D_line->b_purp   $ aBuffer[ 4]   },;
             {|| D_line->b_prior  $ aBuffer[ 5]   },;
             {|| D_line->ptype_id $ aBuffer[ 6]   },;
             {|| D_line->pline_id $ aBuffer[ 7]   },;
             {|| D_line->size_id  $ aBuffer[ 8]   },;
             {|| STR(D_line->value_id,8,3)$ aBuffer[ 9]},;
             {|| IIF(aBuffer[10] == "Yes",d_prom->(Found()) .AND. FoundOrder(aBuffer[11]),!d_prom->(Found()) ) },;
             {|| TRUE    },;
             {|| d_line->cpwkstn_id >= aBuffer[12][1] .AND. d_line->cpwkstn_id <= aBuffer[12][2]  },;
             {|| d_line->CPPROC_ID $ aBuffer[13] },;
             {|| d_line->CP_PCCODE >= aBuffer[14][1]  .AND. d_line->CP_PCCODE <= aBuffer[14][2] },;
             {|| IF(right(aBuffer[15],1)==".",;
                          left(aBuffer[15],len(aBuffer[15])-1),;
                          aBuffer[15];
                    ) $ d_line->b_remark },;
             {|| IIF(aBuffer[ 16] == 'Yes',!Empty(d_line->Pack_mem) ,Empty(d_line->Pack_mem) )   },;
             {|| trim(aBuffer[17]) $ d_line->route_id },;
             {|| slack_01("d_line") >= aBuffer[18][1] .AND. slack_01("d_line") <=  aBuffer[18][2]},;
             {|| checkMLine(aBuffer[19]) },;
             {|| d_line->esnxx_id  $ aBuffer[20]  },;
             {|| d_line->esny_id  $ aBuffer[21]   };
             })
oRep:SetFields( { "b_purp" , "b_id", "b_type" ,"size_id","value_id",;
                "tol_id","b_ncaps","diel_id","cp_bqtyw","cp_bqtys",;
                "cp_bqtyp","CurrentProcnNM('t_line01')","CP_PCCODE",;   //Added by David L. 8/09/97
                "b_prior", "SubStr(b_wkprom,2,4 )", "b_dprom",;
                "date()-pp_dfin", "Slack_01('t_line01')",;          //"Lupdate()-pp_dfin"  changed by
                "QtyPromised('t_line01')","transform(QtyExpected('t_line01'),'999,999')",;
                "b_stat","b_remark","esn_id" ,"Testforproctype('t_line01')","Getbloc()","Getbloc('Instr')","RepStage('t_line01')"})

oRep:SetTitles( {"P","B/N","T","Size","Value","Tol","#C","D","W","S","P",;
                "Current process","Code","Pr","Sched. wk#", "fin. date ","Days at",;
                "Slack","Promis qty","Expect qty","S","Remark","ESN"," ","Location/Date/Time","Location remark","Rep.Stage"})

ELSE//new report VR 27-01-02 TID:2127212
oRep:SetSort({"Proc(D)+Rep.Stage+Hours At"           ,;
              "Type+Line+Size+Proc(D)+R.St.+Ho.At"   ,;
              "Type+Line+Size+Val+Proc(D)+R.St+Ho.At",;
              "Wrkst+Proc(D)+Rep.Stage+Hours At"     ,;
              "Proj+Line+Size+Val+Prior+Slack"        ;
				})
oRep:SetExprSort({"Str(10000-val(cp_pccode),5)+RepStage(Alias())+Str(100000-HoursAt(Alias()))"                                                              ,;
						"ptype_id+pline_id+size_id+Str(10000-val(cp_pccode),5)+RepStage(Alias())+Str(100000-HoursAt(Alias()))"                  ,;
						"ptype_id+pline_id+size_id+str(value_id,8,3)+Str(10000-val(cp_pccode),5)+RepStage(Alias())+Str(100000-HoursAt(Alias()))",;
						"cpwkstn_id+Str(10000-val(cp_pccode),5)+RepStage(Alias())+Str(100000-HoursAt(Alias()))"                                 ,;
						"esn_p2+pline_id+size_id+str(value_id,8,3)+b_prior+Str(Slack_01(Alias()))";
		          })
oRep:SetCrit({       "Scheduled finish date"     ,;
                     "Min&Max Hours at Rep.Stage",;
                     "Status"                    ,;
                     "Purpose"                   ,;
                     "Priority"                  ,;
                     "Product type"              ,;
                     "Product line"              ,;
                     "Size"                      ,;
                     "Value"                     ,;
                     "Promised Batches"          ,;
                     "Project"                   ,;
                     "Workstation"               ,;
                     "Processes"                 ,;
                     "Pc_code"                   ,;
                     "Remark"                    ,;
                     "Dialecter"                 ,;
                     "Route filter"              ,;
                     "Slack from...to"           ,;
                     "Reporting Stage"           ,;
                     "ESN(XX)"                   ,;
                     "ESN(Y)"                     ;
                   })

oRep:SetCheck({     {|o| GetFinishDate(o,{Ctod("  /  /  "),Ctod("  /  /  ")}) } ,;
                    {|o| GetHours(o,{0,24}) }    ,;
                    {|o| critBrowse( o ,{"all"},"c_bstat",    , {|| c_bstat->b_stat } ,{|| c_bstat->bstat_nme } )},;
                    {|o| critBrowse( o ,{"all"},"c_bpurp",    , {|| c_bpurp->b_purp } ,{|| c_bpurp->bpurp_nme } )},;
                    {|o| critBrowse( o ,{"all"},"c_prior",    , {|| c_prior->b_prior },{|| c_prior->bprior_nm } )},;
                    {|o| critBrowse( o ,{"all"},"c_ptype",    , {|| c_ptype->ptype_id},{|| c_ptype->ptype_nm }  )},;
                    {|o| critBrowse( o ,{"all"},"c_pline",    , {|| c_pline->pline_id},{|| c_pline->pline_nm }  )},;
                    {|o| critBrowse( o ,{"all"},"c_size" ,    , {||  c_size->size_id },{|| Space(60) }          )},;
                    {|o| critBrowse( o ,{"all"},"c_value",    , {|| Str(c_value->value_id,8,3) } ,{|| c_value->value_nm } )},;
                    {|o| GetQtyChoice(o,{"No"}) } ,;//new criteria for TAPI
                    {|o| critBrowse( o ,{"all"},"c_prornd","c_prornd", {|| c_prornd->PROJ_ID },{|| c_prornd->proj_nm} )},;
                    {|o| Get42Wkstn(o,{"0000","9999"})}    ,;
                    {|o| critBrowse( o ,{"all"},"c_proc" ,    , {|| c_proc->proc_id  }  ,{|| c_proc->proc_nme } )},;
                    {|o| Getpccode2(o,{"0000","9999"})}   ,;
                    {|o| GetRemark(o,{space(50)})}       ,;
                    {|o| GetDialecter(o,{space(1)})}     ,;
                    {|o| GetRoute(o,{space(4)}) }        ,;
                    {|o| GetSlack(o,{0,50}) }            ,;
                    {|o| GetStage(o,{"Arrive"}) }   ,;  //new criteria for TAPI
                    {|o| critBrowse( o ,{"all"},"c_esnxx",    , {|| c_esnxx->esnxx_id } ,{|| c_esnxx->esnxx_nm } )},;
                    {|o| critBrowse( o ,{"all"},"c_esny" ,    , {|| c_esny->esny_id }   ,{|| c_esny->esny_nm } )} ;
              })
oRep:SetBuffer({ "Scheduled finish date", "Min&Max Hours at Rep.Stage",;
                  "c_bstat","c_bpurp","c_prior","c_ptype","c_pline",;
                  "c_size","c_value","Promised Batches","Project",;
                  "Workstation","Processes","Pc_code","Remark",;
                  "Dialecter","Route filter","Slack",;
                  "Reporting Stage","ESN(XX)","ESN(Y)"})

oRep:SetQueryBlocks({ {|| d_line->b_dprom  <= aBuffer[ 1]   },;
             {|| HoursAt("d_line") >= aBuffer[2][1] .AND. HoursAt("d_line") <=  aBuffer[2][2]},;
             {|| D_line->b_STAT     $ aBuffer[ 3]   },;
             {|| D_line->b_purp     $ aBuffer[ 4]   },;
             {|| D_line->b_prior    $ aBuffer[ 5]   },;
             {|| D_line->ptype_id   $ aBuffer[ 6]   },;
             {|| D_line->pline_id   $ aBuffer[ 7]   },;
             {|| D_line->size_id    $ aBuffer[ 8]   },;
             {|| STR(D_line->value_id,8,3) $ aBuffer[ 9]},;
             {|| IIF(aBuffer[10]    == "Yes",d_prom->(Found()) ,!d_prom->(Found()) ) },;
             {|| d_line->esn_p2       $  Alltrim(aBuffer[11])},;
             {|| d_line->cpwkstn_id >= aBuffer[12][1] .AND. d_line->cpwkstn_id <= aBuffer[12][2]  },;
             {|| d_line->CPPROC_ID  $ aBuffer[13] },;
             {|| d_line->CP_PCCODE  >= aBuffer[14][1]  .AND. d_line->CP_PCCODE <= aBuffer[14][2] },;
             {|| IF(right(aBuffer[15],1)==".",;
                          left(aBuffer[15],len(aBuffer[15])-1),;
                          aBuffer[15];
                    ) $ d_line->b_remark },;
             {|| d_line->diel_id    $ aBuffer[16]  },;
             {|| trim(aBuffer[17])  $ d_line->route_id },;
             {|| slack_01("d_line")    >= aBuffer[18][1] .AND. slack_01("d_line") <=  aBuffer[18][2]},;
             {|| RepStage("d_line") $ aBuffer[19]  },;
             {|| d_line->esnxx_id   $ aBuffer[20]  },;
             {|| d_line->esny_id    $ aBuffer[21]   } ;
             })

oRep:SetFields( {"b_purp"    												,; //1
                 "b_id"      												,; //2
					  "size_id"   												,; //3
					  "value_id"  												,; //4
                 "cp_bqtyw" 											  	,; //5
					  "cp_bqtys"   											,; //6
                 "cp_bqtyp"												,; //7
					  "CurrentProcnNM('t_line01')"		         	,; //8
                 "b_prior"													,; //9
					  "b_dprom"											,; //10
					  "RepStage('t_line01')"								,; //11
					  "GetDate01('t_line01')"								,; //12
					  "GetTime01('t_line01')"								,; //13
					  "HoursAt('t_line01')"									,; //14
                 "Slack_01('t_line01')"									,; //15
                 "QtyPromised('t_line01')"                     ,; //16
					  "transform(QtyExpected('t_line01'),'999,999')",; //17
                 "b_stat"													,; //18
					  "b_remark"                               		,; //19
					  "esn_id" 													,; //20
					  "Testforproctype('t_line01')" })
//"Padl(substr(ALLtrim(b_remark),-22),22)"		,; //19
oRep:SetTitles( {"P"					 ,; //1
	              "B/N"				 ,; //2
					  "Size"				 ,; //3
					  "Value"			 ,; //4
					  "W"					 ,; //5
					  "S"				    ,; //6
					  "P"					 ,; //7
                 "Current process",; //8
					  "Pr"				 ,; //9
					  "fin. date "		 ,; //10
					  "Rep.Stage"		 ,; //11
					  "Date"	   		 ,; //12
					  "Time"				 ,; //13
					  "Hours At"		 ,; //14
                 "Slack"			 ,; //15
					  "Promis qty"		 ,; //16
					  "Expect qty"		 ,; //17
					  "S"					 ,; //18
					  "Remark"			 ,; //19
					  "ESN"				 ,; //20
					  " "              })
ENDIF



oRep:SetDB({"D_LINE","C_LEADT","C_PROC",;
            "d_prom","c_expqty","c_thqty",;
            "c_ptype","c_pline","c_size","m_linemv","d_ord","d_ordreq","m_loc","d_Irrfin","c_expqty","c_tol","d_finqc"})
oRep:cbSetRelation := {|o| rel01(o)}
oRep:lSuperSmart    :=   .T.
oRep:aSuperCriteria :=   {4,7,13}//{3,4,7,13}
oRep:aSuperIndexes  :=   {;
                          {"dlptsvn",.f. },;
                          {"dlinepic",.f.},;
                          {"dlcpptsv",.f.} }//{"dlsptsvn",.f.},; Vitaly 17-06-02
// A request was made to see all batches. "C" and "D" batches
// will be by default unselected: RPQC01V1.DEF. ID:99021705
oRep:cPrepDbf := "t_line01"
oRep:cRepDbf := "t_line01"
IF ReportFunc == "RPQC01V1"
   oRep:cRepFileName := "RPQC01V1"
   oRep:aReports :={"RPQC01V1","RPQC01V1","RPQC01V1","RPQC01V1","RPQC01V1",;
                    "RPQC01V3","RPQC01V1","RPQC01V2","RPQC01V1","RPQC01V1","RPQC01V1","RPQC01V1"}
ELSE
   oRep:cRepFileName := "RPQC01V4" //VR 27-01-02 TID:2127212
   oRep:aReports :={"RPQC01V4","RPQC01V4","RPQC01V4","RPQC01V4","RPQC01V5"}
ENDIF
//rpqc01v2 will stick a form feed after each workstation
//rpqc01v3,v4 will group on value
oMyRep := oRep
cFileName := "t_line01"
oRep:lPrepDBF := TRUE
oRep:Exec()
oRep := NIL
oMyrep := NIL
deletetempfiles({"t_line01"})
Return NIL


FUNCTION CurrentQtyWSP(cFileName)
LOCAL cW , cS , cP
c_proc->( DbSeek( (cFileName)->cpproc_id ) )
cW := Transform( (cFileName)->cp_bqtyw,"999") + if( c_proc->must_w,"*"," " )
cS := Transform( (cFileName)->cp_bqtys,"99,999") + if( c_proc->must_s,"*"," " )
cP := Transform( (cFileName)->cp_bqtyp,"9,999,999") + if( c_proc->must_p,"*"," " )
RETURN cW + " "  + cS + " " + cP

FUNCTION CurrentProcnNM(cFileName)
// current proc seeked in previos function CurrentQtyWSP()
RETURN (cFileName)->cpproc_id+ ":" + myprocname((cFileName)->cpproc_id)


Function Testforproctype(cAlias)
LOCAL cRetVal := " "
m_linemv->(DBSEEK((cAlias)->b_id + (cAlias)->cpproc_id))
DO WHILE ((cAlias)->b_id + (cAlias)->cpproc_id) == (m_linemv->b_id + m_linemv->cpproc_id)
     IF(m_linemv->fin)
       m_linemv->(DBSKIP())
     ELSE
       IF(m_linemv->proc_type == "R")
         cRetVal := "*"
         EXIT
       ELSE
         cRetVal := " "
         EXIT
       ENDIF
    ENDIF
ENDDO
Return cRetVal



FUNCTION QtyExpected(cFileName,cFileExp,lExp)
LOCAL aCalc
LOCAL nQty,nSelect
default cFileExp to cFileName
default lExp to FALSE
nSelect := Select()
//GenOpenfiles({"c_proc","c_expqty","c_thqty"})
Select (nSelect)
aCalc := AvailbleForAnyOne(cFileName)//BatchCalc1(cFileName,cFileExp,lExp)
nQty := aCalc[1][2]//aCalc[2]
RETURN  nQty

function checkMLine(aCrits)
local lRetval := .F.
IF select("m_linemv") == 0
     genopenfiles({"m_linemv"})
ENDIF
m_linemv->(ordsetfocus("ilnmvbs"))
if m_Linemv->(dbseek(d_line->b_id))
     while m_linemv->b_id == d_line->b_id
          IF m_linemv->cp_pccode >= aCrits[1] .AND. m_linemv->cp_pccode <= aCrits[2] .AND.;
             m_linemv->cp_dfin >= aCrits[3] .AND. m_linemv->cp_dfin <= aCrits[4]
               lRetVal := .T.
               exit
          ENDIF
          m_linemv->(dbskip())
     end
endif
return lRetVal
//////////////////////////vitaly 00122715/257
STATIC FUNCTION Rel01(o)
*
*
*
LOCAL nSelect := Select()
FIELD B_id,poln_id,porq_id
IF SELECT("d_line") <> 0
   D_prom->(ORDSETFOCUS(1))
	D_ord->(ORDSETFOCUS(1))
	D_ordreq->(ORDSETFOCUS(1))
	SELECT "d_line"
   SET RELATION  TO  B_id INTO D_prom
	SELECT "d_prom"
	SET RELATION  TO  str(poln_id,9,2) INTO D_ord,str(porq_id,9,2) INTO D_ordreq
ENDIF
SELECT SELECT(nSelect)

RETURN ""
////////////////////////////
Function FoundOrder(aPOstats)

Local cType := "T"
Local cBid := d_prom->b_id
Local lRetVal := FALSE

While cBid == d_prom->b_id .AND. !lRetVal
		IF d_ord->(Found())
	      cType := d_ord->poln_type
      ELSEIF d_ordreq->(Found())
	      cType := d_ordreq->poln_type
      ENDIF
		IF Empty(GetBuffer("P/O Type")) .OR. (cType <> "T" .AND. cType $ aPOstats)
			lRetVal := TRUE
		ENDIF
      d_prom->(dbskip(1))
End

Return lRetVal

FUNCTION QtyPromised(cFileName)

LOCAL nPromQty := 0
Local cType := "T"
local aToFrom := GetBuffer("Scheduled finish date")

d_prom->(ordsetfocus("iprombn"))
IF d_prom->( DbSeek( (cFileName)->b_id ) )
While (cFileName)->b_id == d_prom->b_id
		IF !Empty(GetBuffer("P/O Type")) .AND. d_ord->(Found())
	      cType := d_ord->poln_type
      ELSEIF !Empty(GetBuffer("P/O Type")) .AND. d_ordreq->(Found())
	      cType := d_ordreq->poln_type
      ENDIF
		IF (Empty(GetBuffer("P/O Type")) .OR. (cType <> "T" .AND. cType $ GetBuffer("P/O Type"))) .AND. ;
			(Empty(aToFrom) .OR. (d_ord->(Found())  .AND. (!Empty(aToFrom) .AND. mrkorddue(.T.,'d_ord') <=  aToFrom )))
			nPromQty += d_prom->qty_prom
		ENDIF
      d_prom->(dbskip(1))
End
ENDIF

RETURN  nPromQty

////////////////////////////
Function HoursAt(cFileName) //VR 27-01-02 TID:2127212
local cRetVal
IF RepStage(cFileName) == "F"
	cRetVal := Hours({(cFileName)->pp_dfin,(cFileName)->pp_tfin},{date(),time()})
ELSEIF RepStage(cFileName) == "A"
	cRetVal := Hours({(cFileName)->cp_darr,(cFileName)->cp_tarr},{date(),time()})
ELSEIF RepStage(cFileName) == "S"
	cRetVal := Hours({(cFileName)->cp_dsta,(cFileName)->cp_tsta},{date(),time()})
ELSE
	cRetVal := 999
ENDIF
Return cRetVal

Function GetTime01(cFileName) //VR 27-01-02 TID:2127212
LOCAL nRetVal
IF RepStage(cFileName) == "F"
	nRetVal := (cFileName)->pp_tfin
ELSEIF RepStage(cFileName) == "A"
	nRetVal := (cFileName)->cp_tarr
ELSEIF RepStage(cFileName) == "S"
	nRetVal := (cFileName)->cp_tsta
ELSE
	nRetVal := "  :  "
ENDIF
Return nRetVal

Function GetDate01(cFileName) //VR 27-01-02 TID:2127212
LOCAL nRetVal
IF RepStage(cFileName) == "F"
	nRetVal := (cFileName)->pp_dfin
ELSEIF RepStage(cFileName) == "A"
	nRetVal := (cFileName)->cp_darr
ELSEIF RepStage(cFileName) == "S"
	nRetVal := (cFileName)->cp_dsta
ELSE
	nRetVal := CTOD("  /  /  ")
ENDIF
Return nRetVal

Function GetHforSort(cStage) //VR 07-05-02 TID:2127212
LOCAL nRetVal
IF cStage == "F"
	nRetVal := str(field->FINHOURS)
ELSEIF cStage == "A"
	nRetVal := str(field->ARRHOURS)
ELSEIF cStage == "S"
	nRetVal := str(field->STAHOURS)
ELSE
	nRetVal := str(field->HOURS)
ENDIF
Return nRetVal

Function RepStage(cFileName) //VR 27-01-02 TID:2127212
local cRetVal
IF Empty((cFileName)->cp_darr) .AND. Empty((cFileName)->cp_dsta) .AND. !Empty((cFileName)->pp_dfin)
	cRetVal := "F"
ELSEIF !Empty((cFileName)->cp_darr) .AND. Empty((cFileName)->cp_dsta) .AND. !Empty((cFileName)->pp_dfin)
	cRetVal := "A"
ELSEIF !Empty((cFileName)->cp_darr) .AND. !Empty((cFileName)->cp_dsta) .AND. !Empty((cFileName)->pp_dfin)
	cRetVal := "S"
ELSE
	cRetVal := "?"
ENDIF
Return cRetVal

Static function Hours(aFirst,aSecond) //VR 27-01-02 TID:2127212
local cRetVal
local nTemp_1,nTemp_2,nTempGen
SET DECIMALS TO 2
SET FIXED ON

cRetVal := (aSecond[1] - aFirst[1]) * 24 //hourse
nTemp_1 := (VAL(SubStr(aSecond[2],1,2)) * 60 + VAL(SubStr(aSecond[2],4,2)))
nTemp_2 := (VAL(SubStr(aFirst[2],1,2))  * 60 + VAL(SubStr(aFirst[2],4,2)))
nTempGen := ((nTemp_1 - nTemp_2) % 60) / 60 / 1.66   //it's 0.28 from 2.28
cRetVal := ( ;
	          cRetVal + ;
	         ( ((nTemp_1 - nTemp_2) / 60) - ( ((nTemp_1 - nTemp_2) % 60) / 60) );
			  ) + ;
			  nTempGen
Return cRetVal


Function DecimalToHourse(nHours)
local nRetVal
nHouRs := nHours * 60
nRetVal := (nHouRs/60 - nHours%60/60) + nHours%60/60/1.66
Return nRetVal

FUNCTION Slack_01(cFileName)

LOCAL nRetVal

nRetVal := (cFileName)->b_dprom - (cFileName)->ExpFinDate

RETURN nRetVal

function Getbloc(cMode)

local cRet := "Prod"
local nOldarea := Select()

m_loc->(ordsetfocus(1))


if m_loc->(dbseek(t_line01->b_id))
	While m_loc->b_id == t_line01->b_id
			m_loc->(dbskip(1))
	End
	m_loc->(dbskip(-1))
     IF  Empty(cMode)
          cRet := Trnslt(m_loc->loc)+" #"+dtoc(m_loc->dadd_rec)+" #"+m_loc->tlu_rec
     ELSE
          cRet := TrnsltRmk(m_loc->cremark)+"--"+IIF(!Empty(m_loc->C_FRMRK),m_loc->C_FRMRK," ")
     ENDIF

endif

dbselectarea(nOldarea)

return cRet

Function Trnslt(cHebloc)

local cEngLoc := " "

do case
   case "˜…–‰" $ cHebloc
         cEngLoc := "Prod.Floor"
   case "š‰…˜" $ cHebloc
         cEngLoc := "Ronit/ENG"
   case "Œˆ‰" $ cHebloc
         cEngLoc := "Meital/ENG"
   case "„‰˜‰€" $ cHebloc
         cEngLoc := "Irina/ENG"
   case "‰™" $ cHebloc
         cEngLoc := "Shani/ENG"
  case "‘‰˜‰€" $ cHebloc
         cEngLoc := "Iris/ENG"
  case "„ƒ…„‰" $ cHebloc
       cEngLoc := "Yehuda/ENG"
  case "Œ€‰˜" $ cHebloc
       cEngLoc := "Ariel/ENG"
  case "˜‰ƒ" $ cHebloc
       cEngLoc := "Janet/ENG"
  case "€‘" $ cHebloc
       cEngLoc := "Basanel/ENG"
  case "„ƒ…„‰" $ cHebloc
       cEngLoc := "Yehuda/ENG"
end

Return cEngLoc

function TrnsltRmk(cremark)

local cEngRmk := " "

do case
   case "š…‰‡„" $ cremark
        cEngRmk := "Lack of guidelines"
   case "š…‰—š" $ cremark
        cEngRmk := "Abnormalities"
   case "Œ™" $ cremark
        cEngRmk := "Inspection of an engineer  at the end of a stage"
   case "š…‘…˜”" $ cremark
        cEngRmk := "Delay by Prod.Control"
end

Return cEngRmk
