#INCLUDE "rwmake.ch"
#INCLUDE "topconn.ch"

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณFREVEN    บ Autor ณ AP7 IDE            บ Data ณ  29/10/01   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDeSCricao ณ IncluSAo do Frete para notas de Entrada                    บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
/*/

User Function CRF016()

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Declaracao de Variaveis                                             ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

Local aOrd 		:= {}
Local cDeSC1	:= "Este programa tem como objetivo imprimir relatorio "
Local cDeSC2	:= "de acordo com os parametros informados pelo usuario."
Local cDeSC3	:= "Relacao de Fretes - Notas de Beneficiamento)"
Local cPict		:= ""
Local titulo	:= "Relacao de Fretes de Vendas"
Local nLin		:= 80
Local Cabec1	:= " Nota   EMISSAO   Cliente                                 Cidade                       Peso  Devol Tot Frete   Tot Nota  Perc  PVxPC"
Local Cabec2	:= ""
Local imprime   := .T.
Local nFreteT 	:= 0
Local nTQdte 	:= 0
Local nTValF 	:= 0
Local nNFVal 	:= 0
Local aArea		:= GetArea()

Private CbTxt        := ""
Private lEnd         := .F.
Private lAbortPrint  := .F.
Private limite       := 132
Private tamanho      := "M"
Private nomeprog     := "CRF016"
Private nTIPO        := 15
Private aReturn      := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey     := 0
Private cPerg       := "CRF016"
Private cbtxt		:= Space(10)
Private cbcont		:= 00
Private CONTFL		:= 01
Private m_pag		:= 01
Private wnrel		:= "CRF016"
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

/*/
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
Local nTQdte := 0
Local nTValF := 0
Local nNFVal := 0
Local cQuery := ""
Local aStruSC5 	:= SC5->(dbStruct())
Local aStruSD2 	:= SD2->(dbStruct()) 
Local aStruSF2 	:= SF2->(dbStruct())
Local aStruSD1 	:= SD2->(dbStruct()) 
Local aStruSF1 	:= SF2->(dbStruct())
Local nSC5	 	:= 0
Local nSD2	 	:= 0 
Local nSF2	 	:= 0
Local nSD1	 	:= 0 
Local nSF1	 	:= 0
Local aArea		:= GetArea()
Local cTes		:= Alltrim(GetNewPar("MV_TESRTBN","999"))

cQuery := "SELECT F2_TRANSP, F2_DOC, F2_EMISSAO, "
cQuery += "SUM(D2_QUANT) as QtdVen, SUM(D2_QTDEDEV) as QtdDev, "
cQuery += "SUM(D2_CUSTO1) as CUSTO1 "
cQuery += "FROM "
cQuery += RetSqlName('SF2') + " F2, "
cQuery += RetSqlName('SD2') + " D2, "
cQuery += RetSqlName('SC5') + " C5 
cQuery += "WHERE "
cQuery += "F2.F2_FILIAL = '"+xFilial("SF2")+"' "
cQuery += "AND F2_TRANSP >= '" + MV_PAR01 + "' "
cQuery += "AND F2_TRANSP <= '" + MV_PAR02 + "' "
cQuery += "AND F2_EMISSAO >= '" + dtos(MV_PAR03) + "' "
cQuery += "AND F2_EMISSAO <= '" + dtos(MV_PAR04) + "' "
cQuery += "AND F2_DOC = D2_DOC "
cQuery += "AND F2_SERIE = D2_SERIE "
cQuery += "AND F2_FRETE > 0 " 
cQuery += "AND F2_TIPO = 'B' "
cQuery += "AND F2.D_E_L_E_T_ <> '*' "
cQuery += "AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += "AND D2.D2_CF = '5901' "           
cQuery += "AND D2.D_E_L_E_T_ <> '*' "           
cQuery += "AND C5.C5_FILIAL = '"+xFilial("SC5")+"' "
cQuery += "AND C5_NUM = D2_PEDIDO "
cQuery += "AND C5_TPFRETE = 'C' "
cQuery += "AND C5.D_E_L_E_T_ <> '*' "
cQuery += "GROUP BY F2_TRANSP, F2_DOC, F2_EMISSAO "
cQuery += "ORDER BY F2_TRANSP, F2_DOC"

cQuery:=ChangeQuery(cQuery)

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cALiasREG",.T.,.T.)
	
For nSC5 := 1 To len(aStruSC5)
	If aStruSC5[nSC5][2] <> "C" .And. FieldPos(aStruSC5[nSC5][1])<>0
		TcSetField("cALiasREG",aStruSC5[nSC5][1],aStruSC5[nSC5][2],aStruSC5[nSC5][3],aStruSC5[nSC5][4])
	EndIf
Next nSC5

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

cQuery := "SELECT F1_DOC, F1_EMISSAO, F1_FORNECE, F1_LOJA, F1_TIPO, "
cQuery += "SUM(D1_QUANT) as QtdVen, F1_VALBRUT, D1_X_TRANS "
cQuery += "FROM "
cQuery += RetSqlName('SF1') + " F1, "
cQuery += RetSqlName('SD1') + " D1 "
cQuery += "WHERE "
cQuery += "F1.F1_FILIAL = '"+xFilial("SF1")+"' "
cQuery += "AND F1_EMISSAO >= '" + dtos(MV_PAR03) + "' "
cQuery += "AND F1_EMISSAO <= '" + dtos(MV_PAR04) + "' "
cQuery += "AND F1_SERIE = D1_SERIE "
cQuery += "AND F1_DOC = D1_DOC "
cQuery += "AND F1.D_E_L_E_T_ <> '*' "           
cQuery += "AND D1.D1_FILIAL = '"+xFilial("SD1")+"' "
cQuery += "AND D1_X_TRANS >= '" + MV_PAR01 + "' "
cQuery += "AND D1_X_TRANS <= '" + MV_PAR02 + "' "                                            
cQuery += "AND D1_TES IN ("+cTes+") " 
cQuery += "AND D1.D_E_L_E_T_ <> '*' "
cQuery += "GROUP BY D1_X_TRANS, F1_DOC, F1_EMISSAO, F1_FORNECE, F1_LOJA, F1_TIPO, F1_VALBRUT "
cQuery += "ORDER BY D1_X_TRANS, F1_DOC ASC "

cQuery:=ChangeQuery(cQuery)

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cALiasENT",.T.,.T.)

For nSD1 := 1 To len(aStruSD1)
	If aStruSD1[nSD1][2] <> "C" .And. FieldPos(aStruSD1[nSD1][1])<>0
		TcSetField("cALiasENT",aStruSD1[nSD1][1],aStruSD1[nSD1][2],aStruSD1[nSD1][3],aStruSD1[nSD1][4])
	EndIf
Next nSD1
 
For nSF1 := 1 To len(aStruSF1)
	If aStruSF1[nSF1][2] <> "C" .And. FieldPos(aStruSF1[nSF1][1])<>0
		TcSetField("cALiasENT",aStruSF1[nSF1][1],aStruSF1[nSF1][2],aStruSF1[nSF1][3],aStruSF1[nSF1][4])
	EndIf
Next nSF1

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
	
	SA4->( DbSetOrder(1) )
	SA4->( dbseek( xFilial('SA4') + cAliasREG->F2_TRANSP ) )
	If !EMPTY(ALLTRIM(SA4->A4_CODTRAN))                            
		SA4->(Dbseek(xFilial('SA4') + SA4->A4_CODTRAN))
	EndIf	

	@ prow()+1,000 PSAY 'Transportadora: ' + SA4->A4_COD + '-' + ;
	SA4->A4_NOME
	
	@ prow()+1,000 PSAY ' '
	
	cTRANSP := cAliasREG->F2_TRANSP
	while cAliasREG->( !EOF() .and. F2_TRANSP == cTRANSP ) .and. ;
		prow() < 55 .and. !lAbortPrint
		
		SF2->(DbSetOrder(1))
		SF2->(DbSeek(xFilial('SF2')+cAliasREG->F2_DOC))
		SC6->(DbSetOrder(4))
		SC6->(DbSeek(xFilial('SC6')+cAliasREG->F2_DOC))
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
		
		If SC5->C5_TIPO $ "B#D"
			SA2->(DbSetOrder(1))
			SA2->(dbseek( xFilial('SA2')+SF2->(F2_CLIENTE+F2_LOJA)))
			cCodi:=SA2->A2_COD
			cLoja:=SA2->A2_LOJA
			cNome:=SA2->A2_NOME
			cMuni:=SA2->A2_MUN
		Else
			SA1->(DbSetOrder(1))
			SA1->(dbseek(xFilial('SA1')+SF2->(F2_CLIENTE+F2_LOJA)))
			cCodi:=SA1->A1_COD
			cLoja:=SA1->A1_LOJA
			cNome:=SA1->A1_NOME
			cMuni:=SA1->A1_MUN
		EndIf
		
		@ prow()+1,000 PSAY cAliasREG->F2_DOC
		@ prow()  ,007 PSAY cAliasREG->F2_EMISSAO picture '@D'
		@ prow()  ,018 PSAY cCodi + cLoja + '-' + left(cNome,30)
		@ prow()  ,058 PSAY left(cMuni,25)
		@ prow()  ,085 PSAY cAliasREG->QtdVen picture '@E 99.999'
		@ prow()  ,092 PSAY cAliasREG->QtdDev picture '@E 99.999'
		@ prow()  ,100 PSAY IIf(cAliasREG->(QtdVen=QtdDev),0,SF2->F2_FRETE) picture '@E 9,999.99'
		@ prow()  ,110 PSAY IIf(cAliasREG->(QtdVen=QtdDev),0,nValor) picture '@E 99,999.99'
		@ prow()  ,120 PSAY IIf(cAliasREG->(QtdVen=QtdDev),0,(SF2->F2_FRETE/nValor)*100) picture '@E 99.99'
		@ prow()  ,126 PSAY IIf(cAliasREG->(QtdVen=QtdDev),0,((nValor/nCusto)-1)*100) picture '@E 999.99'
		
		nQtdTotal += cAliasREG->QtdVen
		nQtdTrans += cAliasREG->QtdVen
		nDevTotal += cAliasREG->QtdDev
		nDevTrans += cAliasREG->QtdDev
		nFreTotal += IIf(cAliasREG->(QtdVen=QtdDev),0,SF2->F2_FRETE)
		nFreTrans += IIf(cAliasREG->(QtdVen=QtdDev),0,SF2->F2_FRETE)
		nCusTotal += IIf(cAliasREG->(QtdVen=QtdDev),0,nCusto)
		nCusTrans += IIf(cAliasREG->(QtdVen=QtdDev),0,nCusto)
		nFatTotal += IIf(cAliasREG->(QtdVen=QtdDev),0,nValor)
		nFatTrans += IIf(cAliasREG->(QtdVen=QtdDev),0,nValor)
		
		cAliasREG->( dbskip() )
		
	EndDo
	If cAliasREG->( EOF() .or. F2_TRANSP <> cTRANSP ) .or. lAbortPrint
		@ prow()+1,084 PSAY '------- ------ --------- ---------- ----- ------'
		@ prow()+1,084 PSAY nQtdTrans picture '@E 999.999'
		@ prow()  ,092 PSAY nDevTrans picture '@E 99.999'
		@ prow()  ,099 PSAY nFreTrans picture '@E 99,999.99'
		@ prow()  ,109 PSAY nFatTrans picture '@E 999,999.99'
		@ prow()  ,120 PSAY (nFreTrans/nFatTrans)*100 picture '@E 99.99'
		@ prow()  ,126 PSAY ((nFatTrans/nCusTrans)-1)*100 picture '@E 999.99'
		
		nFatTrans := 0
		nDevTrans := 0
		nFreTrans := 0
		nQtdTrans := 0
		nCusTrans := 0
		flagcab   := 0
		@ prow()+1,000 PSAY repl( '-', 132 )
		
		dbSelectArea("cALiasENT")
		dbgotop()
		
		while cALiasENT->(!EOF() .and. D1_X_TRANS == cTRANSP )
			If lAbortPrint
				@nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
				Exit
			EndIf
			
			If cALiasENT->F1_TIPO $ "B#D"
				SA1->(DbSetOrder(1))
				SA1->(DbSeek(xFilial('SA1') + cALiasENT->( F1_FORNECE + F1_LOJA)))
				cCodi:=SA1->A1_COD
				cLoja:=SA1->A1_LOJA
				cNome:=SA1->A1_NOME
				cMuni:=SA1->A1_MUN
			Else
				SA2->(DbSetOrder(1))
				SA2->(DbSeek(xFilial('SA2')+ cALiasENT->(F1_FORNECE+F1_LOJA)))
				cCodi:=SA2->A2_COD
				cLoja:=SA2->A2_LOJA
				cNome:=SA2->A2_NOME
				cMuni:=SA2->A2_MUN
			EndIf
			
			NFreteT := MV_PAR05
			
			/*
			Nota   EMISSAO   Cliente                                 Cidade                       Peso  Tot Frete  Tot Nota
			----------------------------------------------------------------------------------------------------------------
			999999 99/99/9999 99999999-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx xxxxxxxxxxxxxxxxxxxxxxxxx  99,999 9.999,99  99.999,99
			999999 99/99/9999 99999999-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx xxxxxxxxxxxxxxxxxxxxxxxxx  99,999 9.999,99  99.999,99
			------------------------------------------------------------------------------------------------------------------------------------
			Total Geral ===>          														     999,999  99,999  99.999,99
			1         2         3         4         5         6         7         8         9        10        11        12        13
			012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901
			*/

			If flagcab = 0
				@ prow()+2,000 PSAY "Relacao de Fretes de Vendas - Notas de Entrada"
				@ prow()+1,000 PSAY " Nota   EMISSAO   Cliente                                 Cidade                       Peso  Tot Frete     Tot Nota"
				@ prow()+1,000 PSAY repl( '-', 132 )
				flagcab := 1
			EndIf
			
			@ prow()+1,000 PSAY cALiasENT->F1_DOC
			@ prow()  ,007 PSAY cALiasENT->F1_EMISSAO picture '@D'
			@ prow()  ,018 PSAY cCodi + cLoja + '-' + left(cNome,30)
			@ prow()  ,058 PSAY left(cMuni,25)
			@ prow()  ,087 PSAY cALiasENT->QtdVen    picture '@E 99.999'
			@ prow()  ,094 PSAY QtdVen*nFreteT picture '@E 9,999.99'
			@ prow()  ,106 PSAY cALiasENT->F1_VALBRUT     picture '@E 99,999.99' //F1_VALBRUT
			
			nTQdte := nTQdte + cALiasENT->QtdVen
			nTValF := nTValF + (QtdVen*nFreteT)
			nNFVal := nNFVal + cALiasENT->F1_VALBRUT
			
			cALiasENT->(DbSkip())
		EndDo
		If flagcab = 1
			@ prow()+1,085 PSAY '-------- --------  -----------'
			@ prow()  ,087 PSAY nTQdte picture '@E 999.999'
			@ prow()  ,094 PSAY nTValF picture '@E 99,999.99'
			@ prow()  ,106 PSAY nNFVal picture '@E 999,999,999.99'
			@ prow()+1,000 PSAY repl( '-', 132 )	
		EndIf
	EndIf
	nTQdte := 0
	nTValF := 0
	nNFVal := 0
EndDo                             

@ prow()+1,000 PSAY 'Total Geral ===>'
@ prow()  ,080 PSAY nQtdTotal picture '@E 999,999.999'
@ prow()  ,092 PSAY nDevTotal picture '@E 99.999'
@ prow()  ,099 PSAY nFreTotal picture '@E 999999.99'
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
DbSelectArea("cAliasENT")
cALiasENT->(DbCloseArea())

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
				
RestArea(aArea)

Return(.T.)
