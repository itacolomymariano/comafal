#INCLUDE "rwmake.ch"
#INCLUDE "topconn.ch"

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณCRF011    บ Autor ณ AP7 IDE            บ Data ณ  29/10/01   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDeSCricao ณ Relacao de Fretes de Vendas                                บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
/*/

User Function CRF011()

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Declaracao de Variaveis                                             ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

Local aOrd 		:= {}
Local cDeSC1	:= "Este programa tem como objetivo imprimir relatorio "
Local cDeSC2	:= "de acordo com os parametros informados pelo usuario."
Local cDeSC3	:= "Relacao de Fretes - Notas de Vendas)"
Local cPict		:= ""
Local titulo	:= "Relacao de Fretes de Vendas"
Local nLin		:= 80
Local Cabec1	:= " Nota   EMISSAO   Cliente                                 Cidade                       Peso  Devol Tot Frete   Tot Nota  Perc  PVxPC"
Local Cabec2	:= ""
Local aArea		:= GetArea()

Private CbTxt        := ""
Private lEnd         := .F.
Private lAbortPrint  := .F.
Private limite       := 132
Private tamanho      := "M"
Private nomeprog     := "CRF011"
Private nTIPO        := 15
Private aReturn      := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey     := 0
Private cPerg       := "CRF011"
Private cbtxt		:= Space(10)
Private cbcont		:= 00
Private CONTFL		:= 01
Private m_pag		:= 01
Private wnrel		:= "CRF011"
Private cString 	:= "SF2"

dbSelectArea("SF2")
DbSetOrder(1)


AjustaSX1()

pergunte(cPerg,.F.)

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Monta a interface padrao com o usuario...                           ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

wnrel := SetPrint(cString,NomeProg,cPerg,@titulo,cDeSC1,cDeSC2,cDeSC3,.F. ,aOrd,.F.,Tamanho,,.T.)
If nLastKey == 27
	Return
EndIf

SetDefault(aReturn,cString)

If nLastKey == 27
	Return
EndIf

nTIPO := If(aReturn[4]==1,15,18)

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Processamento. RPTSTATUS monta janela com a regua de processamento. ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

RptStatus({|| RunReport(Cabec1,Cabec2,Titulo,nLin) },Titulo)

RestArea(aArea)

Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณRUNREPORT บ Autor ณ AP5 IDE            บ Data ณ  29/10/01   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDeSCrio ณ Funcao auxiliar chamada pela RPTSTATUS. A funcao RPTSTATUS บฑฑ
ฑฑบ          ณ monta a janela com a regua de procesSAmento.               บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Programa principal                                         บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
/*/

Static Function RunReport(Cabec1,Cabec2,Titulo,nLin)

Local nOrdem := 0
Local cQuery := ""
Local aStruSC5 	:= SC5->(dbStruct())
Local aStruSD2 	:= SD2->(dbStruct()) 
Local aStruSF2 	:= SF2->(dbStruct())
Local nSC5	 	:= 0
Local nSD2	 	:= 0 
Local nSF2	 	:= 0
Local aArea		:= GetArea()
Local cCFOPVenda  	    := "'5101','6101','5102','6102','5103','6103','5104','6104','5105','6105',"
	  cCFOPVenda  	    += "'5106','6106','6107','6108','5109','6109','5110','6110','5111','6111','5112','6112',"
	  cCFOPVenda  	    += "'5113','6113','5114','6114','5115','6115','5116','6116','5117','6117','5118','6118','5119','6119',"
	  cCFOPVenda  	    += "'5120','6120','5122','6122','5123','6123','7101','7102','7105','7106','7127'"

cQuery := "SELECT A4_CODTRAN,F2_TRANSP, D2_DOC, D2_EMISSAO, "
cQuery += "SUM(D2_QUANT) as QtdVen, SUM(D2_QTDEDEV) as QtdDev, "
cQuery += "SUM(D2_CUSTO1) as CUSTO1 "
cQuery += "FROM "
cQuery += RetSqlName('SF2') + " F2, "
cQuery += RetSqlName('SD2') + " D2, "
cQuery += RetSqlName('SC5') + " C5, " 
cQuery += RetSqlName('SA4') + " A4 "
cQuery += "WHERE A4.A4_FILIAL = '"+xFilial("SA4")+"' "
cQuery += "AND A4.D_E_L_E_T_ <> '*' "
cQuery += "AND A4.A4_COD >= '" + MV_PAR01 + "' "
cQuery += "AND A4.A4_COD <= '" + MV_PAR02 + "' "
cQuery += "AND A4.A4_COD = F2.F2_TRANSP "
cQuery += "AND F2.F2_FILIAL = '"+xFilial("SF2")+"' "
cQuery += "AND F2_TRANSP >= '" + MV_PAR01 + "' "
cQuery += "AND F2_TRANSP <= '" + MV_PAR02 + "' "
cQuery += "AND F2_EMISSAO >= '" + dtos(MV_PAR03) + "' "
cQuery += "AND F2_EMISSAO <= '" + dtos(MV_PAR04) + "' "
cQuery += "AND F2_DOC = D2_DOC "
cQuery += "AND F2_SERIE = D2_SERIE "
cQuery += "AND F2_FRETE > 0 " 
cQuery += "AND F2.D_E_L_E_T_ <> '*' "
cQuery += "AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += "AND D2_EMISSAO >= '" + dtos(MV_PAR03) + "' "
cQuery += "AND D2_EMISSAO <= '" + dtos(MV_PAR04) + "' "
cQuery += "AND D2_CLIENTE >= '" + MV_PAR06 + "' "
cQuery += "AND D2_CLIENTE <= '" + MV_PAR07 + "' "
cQuery += "AND D2.D2_CLIENTE <> '004502' "
cQuery += "AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ") "           
cQuery += "AND D2.D2_LOCAL = '01'  AND D2.D2_TIPO = 'N'  AND D2.D_E_L_E_T_ <> '*' "
cQuery += "AND C5.C5_FILIAL = '"+xFilial("SC5")+"' "
cQuery += "AND C5_NUM = D2_PEDIDO "
cQuery += "AND C5_TPFRETE = 'C' "
cQuery += "AND C5.D_E_L_E_T_ <> '*' " 
cQuery += "GROUP BY A4_CODTRAN,F2_TRANSP, D2_DOC, D2_EMISSAO "
cQuery += "ORDER BY A4_CODTRAN,F2_TRANSP, D2_DOC, D2_EMISSAO ASC "

cQuery:=ChangeQuery(cQuery)

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cALiasREG",.T.,.T.)
	
For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField("cALiasREG",aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2
 
For nSF2 := 1 To len(aStruSF2)
	If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
		TcSetField("cALiasREG",aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
	EndIf
Next nSF2

dbSelectArea("cAliasREG")
dbgotop()

/*
Nota   EMISSAO   Cliente                                 Cidade                       Peso  Devol Tot Frete   Tot Nota  Perc  PVxPC
------------------------------------------------------------------------------------------------------------------------------------
TRANSPortadora: xxxxxx-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

999999 99/99/9999 99999999-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx xxxxxxxxxxxxxxxxxxxxxxxxx  99,999 99,999  9.999,99  99.999,99 99,99 999,99
999999 99/99/9999 99999999-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx xxxxxxxxxxxxxxxxxxxxxxxxx  99,999 99,999  9.999,99  99.999,99 99,99 999,99
------- ------ --------- ---------- ----- ------
999,999 99,999 99.999,99 999.999,99 99,99 999,99
------------------------------------------------------------------------------------------------------------------------------------
Total Geral ===>                  999,999 99,999 99.999,99 999.999,99 99,99 999,99
1         2         3         4         5         6         7         8         9        10        11        12        13
012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901
*/

nFatTotal := 0
nDevTotal := 0
nFreTotal := 0
nQtdTotal := 0
nCusTotal := 0
nFatTrans := 0
nDevTrans := 0
nFreTrans := 0
nQtdTrans := 0
nCusTrans := 0

While cAliasREG->(!EOF())
	If lAbortPrint
		@nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
		Exit
	EndIf
	
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTIPO)
	nLin := 8

	DbSelectArea("SA4")     
	DbSetOrder(1)

	If EMPTY(ALLTRIM(cAliasREG->A4_CODTRAN))                            
		If dbseek(xFilial('SA4') + cAliasREG->F2_TRANSP )
			cTRANSP := cAliasREG->F2_TRANSP
		EndIf	
	Else	
		If Dbseek(xFilial('SA4') + cAliasREG->A4_CODTRAN)
			cTRANSP := cAliasREG->A4_CODTRAN
		EndIf	
	EndIf	

	@ prow()+1,000 PSAY 'Transportadora: ' + SA4->A4_COD + '-' + SA4->A4_NOME
	@ prow()+1,000 PSAY ' '
	
	dbSelectArea("cAliasREG")
	
	While cAliasREG->(!EOF()) .and. Iif(Empty(Alltrim(cAliasREG->A4_CODTRAN)),cAliasREG->F2_TRANSP == cTRANSP,cAliasREG->A4_CODTRAN == cTRANSP)  .and. ;
		prow() <= 55
		
		SF2->(DbSetOrder(1))
		SF2->(DbSeek(xFilial('SF2')+cAliasREG->D2_DOC))
		SC6->(DbSetOrder(4))
		SC6->(DbSeek(xFilial('SC6')+cAliasREG->D2_DOC))
		nValor := 0
		nCusto := 0
		cNumped:=SC6->c6_num
		nCusto += cAliasREG->CUSTO1
		nValor += SF2->F2_VALFAT
		
		cCodi:=space(6)
		cLoja:=space(2)
		cNome:=space(40)
		
		SC5->(DbSetOrder(1))
		SC5->(dbseek(xFilial('SC5')+cNumped))
		
		SA1->(DbSetOrder(1))
		SA1->(dbseek(xFilial('SA1')+SF2->(F2_CLIENTE+F2_LOJA)))
		cCodi:=SA1->A1_COD
		cLoja:=SA1->A1_LOJA
		cNome:=SA1->A1_NOME
		cMuni:=SA1->A1_MUN
		
		@ prow()+1,000 PSAY Right(alltrim(cAliasREG->D2_DOC),6)
		@ prow()  ,007 PSAY cAliasREG->D2_EMISSAO picture '@D'
		@ prow()  ,018 PSAY cCodi + cLoja + '-' + left(cNome,30)
		@ prow()  ,058 PSAY left(cMuni,25)
		@ prow()  ,085 PSAY cAliasREG->QtdVen picture '@E 99.999'
		@ prow()  ,092 PSAY cAliasREG->QtdDev picture '@E 99.999'
		@ prow()  ,100 PSAY SF2->F2_FRETE picture '@E 9,999.99'
		@ prow()  ,110 PSAY nValor picture '@E 99,999.99'
		@ prow()  ,120 PSAY (SF2->F2_FRETE/nValor)*100 picture '@E 99.99'
		@ prow()  ,126 PSAY (((nValor/nCusto)-1)*100) picture '@E 999.99'
		
		nQtdTotal += cAliasREG->QtdVen
		nQtdTrans += cAliasREG->QtdVen
		nDevTotal += cAliasREG->QtdDev
		nDevTrans += cAliasREG->QtdDev
		nFreTotal += SF2->F2_FRETE
		nFreTrans += SF2->F2_FRETE
		nCusTotal += nCusto
		nCusTrans += nCusto
		nFatTotal += nValor
		nFatTrans += nValor

		dbSelectArea("cAliasREG")	
		cAliasREG->(DbSkip())
		
	EndDo
	@ prow()+1,084 PSAY '------- ------ --------- ---------- ----- ------'
	@ prow()+1,082 PSAY nQtdTrans picture '@E 999.999'
	@ prow()  ,090 PSAY nDevTrans picture '@E 99.999'
	@ prow()  ,098 PSAY nFreTrans picture '@E 999,999.99'
	@ prow()  ,109 PSAY nFatTrans picture '@E 9999,999.99'
	@ prow()  ,120 PSAY (nFreTrans/nFatTrans)*100 picture '@E 99.99'
	@ prow()  ,126 PSAY ((nFatTrans/nCusTrans)-1)*100 picture '@E 999.99'
	
	nFatTrans := 0
	nDevTrans := 0
	nFreTrans := 0
	nQtdTrans := 0
	nCusTrans := 0
	@ prow()+1,000 PSAY repl( '-', 132 )
EndDo                             

@ prow()+1,000 PSAY 'Total Geral ===>'
@ prow()  ,075 PSAY nQtdTotal picture '@E 999,999.999'
@ prow()  ,092 PSAY nDevTotal picture '@E 99.99'
@ prow()  ,098 PSAY nFreTotal picture '@E 9999999.99'
@ prow()  ,109 PSAY nFatTotal picture '@E 9999,999.99'
@ prow()  ,120 PSAY (nFreTotal/nFatTotal)*100 picture '@E 99.99'
@ prow()  ,126 PSAY ((nFatTotal/nCusTotal)-1)*100 picture '@E 999.99'

SET DEVICE TO SCREEN

If aReturn[5]==1
	dbCommitAll()
	SET PRINTER TO
	OurSpool(wnrel)
EndIf

MS_FLUSH()             
   
DbSelectArea("cAliasREG")
cAliasREG->(DbCloseArea())

RestArea(aArea)

Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณAJUSTASX1 บAutor  ณEduardo Zanardo     บ Data ณ  19/03/03   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDeSC.     ณAjusta o SX1                                                บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ CRF011                                                     บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function AjustaSx1()

Local aArea := GetArea()

PutSx1(cPerg,"01","De Transportadora ?","","","mv_ch1","C",6,0,0,"G","","SA4","","",;
				"MV_PAR01","","","","","","","","","","","","","","","","")  
PutSx1(cPerg,"02","Ate Transportadora ?","","","mv_ch2","C",6,0,0,"G","","SA4","","",;
				"MV_PAR02","","","","","","","","","","","","","","","","")				
PutSx1(cPerg,"03","De Emissao ?","","","mv_ch3","D",8,0,0,"G","","","","",;
				"MV_PAR03","","","","","","","","","","","","","","","","")  
PutSx1(cPerg,"04","Ate Emissao ?","","","mv_ch4","D",8,0,0,"G","","","","",;
				"MV_PAR04","","","","","","","","","","","","","","","","")				
PutSx1(cPerg,"05","Valor Frete/Ton ?","","","mv_ch5","N",4,2,0,"G","","","","",;
				"MV_PAR02","","","","","","","","","","","","","","","","")				
PutSx1(cPerg,"06","De Cliente ?","","","mv_ch6","C",6,0,0,"G","","SA1","","",;
				"MV_PAR06","","","","","","","","","","","","","","","","")  
PutSx1(cPerg,"07","Ate Cliente ?","","","mv_ch7","C",6,0,0,"G","","SA1","","",;
				"MV_PAR07","","","","","","","","","","","","","","","","")				
//PutSx1(cPerg,"08","Gera Financeiro ?","","","mv_ch8","N",1,0,0,"C","","","","",;
//				"MV_PAR08","Sim","","","","Nao","","","","","","","","","","","",;
//				{"Gera Financeiro:","SIM: Gera titulo nos contas a pagar.","NAO: Nใo gera nenhum titulo."},{},{})				
				
RestArea(aArea)

Return(.T.)
