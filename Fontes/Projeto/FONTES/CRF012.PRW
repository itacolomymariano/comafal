#INCLUDE "rwmake.ch"                         
#INCLUDE "Topconn.ch"

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณFatr71    บ Autor ณ Daniel             บ Data ณ  08/07/02   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescricao ณ Carta Frete					                              บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ AP5 IDE                                                    บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
/*/

User Function CRF012()

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Declaracao de Variaveis                                             ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

Local aOrd 			:= {}
Local cDesc1        := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2        := "de acordo com os parametros informados pelo usuario."
Local cDesc3        := "Carta Frete"
Local cPict         := ""
Local titulo       	:= "Carta Frete"
Local nLin         	:= 80
Local Cabec1       	:= ""
Local Cabec2       	:= ""
Local cTitFrete		:= ""
Local cPre			:= ""
Local cNomBen		:= ""

Private CbTxt       := ""
Private cString     := ""
Private lEnd        := .F.
Private lAbortPrint := .F.
Private limite      := 120
Private tamanho     := "M"
Private nomeprog    := "FATR71" // Coloque aqui o nome do programa para impressao no cabecalho
Private nTipo       := 15
Private aReturn     := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey    := 0
Private cPerg       := "FATR71"
Private cbtxt      	:= Space(10)
Private cbcont     	:= 00
Private CONTFL     	:= 01
Private m_pag      	:= 01
Private wnrel      	:= "FATR71" // Coloque aqui o nome do arquivo usado para impressao em disco
Private aResSup		:= {}


ValidPerg()

pergunte(cPerg,.F.)

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Monta a interface padrao com o usuario...                           ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

wnrel := SetPrint(cString,NomeProg,cPerg,@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,.T.,Tamanho,,.F.)

If nLastKey == 27
	Return
Endif


SetDefault(aReturn,cString)

nTipo := If(aReturn[4]==1,15,18)

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Processamento. RPTSTATUS monta janela com a regua de processamento. ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

RptStatus({|| RunReport(Cabec1,Cabec2,Titulo,nLin) },Titulo)
Return

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณRUNREPORT บ Autor ณ AP5 IDE            บ Data ณ  20/05/02   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescrio ณ Funcao auxiliar chamada pela RPTSTATUS. A funcao RPTSTATUS บฑฑ
ฑฑบ          ณ monta a janela com a regua de processamento.               บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Programa principal                                         บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
/*/

Static Function RunReport(Cabec1,Cabec2,Titulo,nLin)

Local nTotalRecords := 0
Local cTrans 	 := ""
Local cPlaca 	 := ""
Local cPosto 	 := ""
Local cNomeTrans := ""
Local cCgc 		 := ""

Local nSF2 		:= 0
Local nSD2 		:= 0
Local nSC5 		:= 0
Local nSA1 		:= 0  

Local aStruSF2 	:= SF2->(dbStruct())
Local aStruSD2 	:= SD2->(dbStruct())
Local aStruSA1 	:= SA1->(dbStruct())
Local aStruSC5 	:= SC5->(dbStruct())

Private cDtEnt 	  := ""
Private nTotFrete := 0
Private nTotPeso  := 0
Private nTotvalor := 0
Private nAdian 	  := 0
Private dEmissao  := Date()
Private nTotFin   := 0
Private nAdFin	  := 0
Private nPerc	  := 0

cQuery := "SELECT F2.F2_TRANSP, F2.F2_DOC, F2.F2_SERIE, A1.A1_COD,"
cQuery += "A1.A1_LOJA, A1.A1_NOME, F2.F2_EMISSAO, F2.F2_VALBRUT, F2.F2_FRETE, F2.F2_PLIQUI "
cQuery += "FROM "
cQuery += RetSqlName("SF2")+" F2, "
cQuery += RetSqlName("SD2")+" D2, "
cQuery += RetSqlName("SC5")+" C5, "
cQuery += RetSqlName("SA1")+" A1 "
cQuery += "WHERE "
cQuery += "		F2.F2_FILIAL 	= '"+xFilial("SF2")+"' AND "
cQuery += "		F2.F2_DOC 		= D2.D2_DOC AND "
cQuery += "		F2.F2_SERIE 	= D2.D2_SERIE AND "
cQuery += "		F2.F2_TRANSP 	= '" + MV_PAR01 + "' AND "
cQuery += "		F2.F2_EMISSAO 	= '" + dtos(MV_PAR02) + "' AND "
cQuery += "		F2.F2_CLIENTE 	= A1_COD AND "
cQuery += "		F2.F2_LOJA 		= A1_LOJA AND "
cQuery += "		F2.D_E_L_E_T_ 	<> '*' AND "
cQuery += "		C5.C5_FILIAL 	= '"+xFilial("SC5")+"' AND "
cQuery += " 	C5_NUM 			= D2.D2_PEDIDO AND "
cQuery += "		C5_TPFRETE 		= 'C' AND "
cQuery += "		F2_FRETE 		> 0  AND "
cQuery += "		C5.D_E_L_E_T_ 	<> '*' AND "  
cQuery += "		D2.D2_FILIAL 	= '"+xFilial("SD2")+"' AND "             
cQuery += "		D2.D_E_L_E_T_ 	<> '*' AND "                             
cQuery += "		A1.A1_FILIAL 	= '"+xFilial("SA1")+"' AND "             
cQuery += "		A1.D_E_L_E_T_ 	<> '*' "               
cQuery += "GROUP BY F2.F2_TRANSP, F2.F2_DOC, F2.F2_SERIE, "
cQuery += "A1.A1_COD, A1.A1_LOJA, A1.A1_NOME, F2.F2_EMISSAO, "
cQuery += "F2.F2_VALBRUT,F2.F2_FRETE, F2.F2_PLIQUI "
cQuery += "ORDER BY F2.F2_DOC"

cQuery:=ChangeQuery(cQuery)

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasTC1",.T.,.T.)

For nSC5 := 1 To len(aStruSC5)
	If aStruSC5[nSC5][2] <> "C" .And. FieldPos(aStruSC5[nSC5][1])<>0
		TcSetField("cAliasTC1",aStruSC5[nSC5][1],aStruSC5[nSC5][2],aStruSC5[nSC5][3],aStruSC5[nSC5][4])
	EndIf
Next nSC5

For nSF2 := 1 To len(aStruSF2)
	If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
		TcSetField("cAliasTC1",aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
	EndIf
Next nSF2

For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField("cAliasTC1",aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2

For nSA1 := 1 To len(aStruSA1)
	If aStruSA1[nSA1][2] <> "C" .And. FieldPos(aStruSA1[nSA1][1])<>0
		TcSetField("cAliasTC1",aStruSA1[nSA1][1],aStruSA1[nSA1][2],aStruSA1[nSA1][3],aStruSA1[nSA1][4])
	EndIf
Next nSA1

cTrans := MV_PAR01
cPlaca := MV_PAR03
cDtEnt := MV_PAR04
nPerc  := MV_PAR05
cPosto := MV_PAR06

cAliasTC1->(dbGoTop())

While !cAliasTC1->(EOF()) .and. !lAbortPrint
	//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
	//ณ Verifica o cancelamento pelo usuario...                             ณ
	//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู
	
	If lAbortPrint
		@nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
		Exit
	Endif
	
	//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
	//ณ Impressao do cabecalho do relatorio. . .                            ณ
	//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู
	
	SA4->( dbSetOrder( 1 ) )
	If SA4->( dbSeek( xFilial("SA4") + cTrans, .F.) )
		cNomeTrans := SA4->A4_NOME
		substring(cNomeTrans := SA4->A4_NOME,1,25)
		cCGC := SA4->A4_CGC
	Else
		cNomeTrans := "INDEFINIDO"
	EndIf
	
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin := 5
	@nLin++
	@nLin++, 70 PSay "VIA MOTORISTA (NAO PAGAVEL)"
	@nLin, 00 PSay "TRANSPORTADOR -"
	@nLin++, 17 PSay cNomeTrans
	@nLin, 00 PSay "PLACA -"
	@nLin++, 09 PSay cPlaca
	@nLin, 00 PSay "CGC -"
	@nLin++,07 PSAY cCgc PICTURE IF(LEN(ALLTRIM(cCgc))>11,"@R !!9.999.999/9999-99","@R 999.999.999-99")
	@nLin++
	@nLin++, 00 PSay "NOTAS"
        @nLin++, 00 PSay "CODIGO/LJ         NOME CLIENTE                NOTA   SERIE   DATA EMI     VALOR      FRETE      P.BRUTO "
        @nLin++, 00 Psay "---------   ------------------------------   ------  -----  ----------  ----------  ----------  ------- "
	
	cAliasTC1->(dbGoTop())
	While !cAliasTC1->(EOF()) .and. !lAbortPrint
		@nLin  , 00 PSay cAliasTC1->(A1_COD + "/" + A1_LOJA)
		@nLin  , 12 PSay Subs(cAliasTC1->A1_NOME,1,30)
		@nLin  , 45 PSay cAliasTC1->F2_DOC
		@nLin  , 54 PSay cAliasTC1->F2_SERIE
		@nLin  , 61 PSay cAliasTC1->F2_EMISSAO 	PICTURE "@D"
		@nLin  , 72 PSay cAliasTC1->F2_VALBRUT	PICTURE "@E 999,999.99"
		@nLin  , 83 PSay cAliasTC1->F2_FRETE		PICTURE "@E 99,999.99"
        @nLin  , 93 PSay cAliasTC1->F2_PLIQUI         PICTURE "@E 999,999.99"

		nTotValor += cAliasTC1->F2_VALBRUT
		nTotFrete += cAliasTC1->F2_FRETE
        ntotpeso  += cAliasTC1->F2_PLIQUI

		nLin++ // Avanca a linha de impressao
		
		cAliasTC1->(dbSkip()) // Avanca o ponteiro do registro no arquivo
	Enddo
        @nLin++
        @nLin  , 00  PSay "TOTAIS:"
		@nLin  , 72 PSay nTotValor	PICTURE "@E 999,999.99"        
        @nLin  , 83 PSay NTOTFRETE  PICTURE "@E 99,999.99"
        @nLin  , 93 PSay NTOTPESO   PICTURE "@E 999,999.99"
        @nLin++
        @nLin, 00 Psay "----------------------------------------------------------------------------------------------------------"
        @nLin++

        @nLin, 00 PSay "POSTO -"
	If ALLTRIM(cPosto) <> ""
		@nLin++, 08 PSay cPosto
		@nLin, 00 PSay "Estou ciente que efetuarei despesas no posto acima listado, no momento da troca da carta-frete, bem como dou quitacao dos valores"
		@nLin++
		@nLin, 00 PSay "abaixo descriminados."
		@nLin++
	Else
		@nLin++, 08 PSay "NAO HA POSTO"
	Endif
	@nLin++
	@nLin, 00 PSay "TOTAL DO FRETE -"
	@nLin++, 17 PSay nTotFrete PICTURE "@E 99,999.99"
	@nLin++
        nAdian := (nTotFrete*nPerc)/100
        @nLin++
	@nLin, 00 PSay "SALDO A PAGAR -"
	@nLin++, 17 PSay nTotFrete - nAdian PICTURE "@E 99,999.99"
	@nLin++
        @nLin, 00 PSay "DATA DA EMISSAO -"
	@nLin++, 20 PSay dEmissao PICTURE "@D"
	@nLin, 00 PSay "DATA LIMITE PARA ENTREGA -"
	@nLin++, 27 PSay cDtEnt PICTURE "@D"
	
	//  @nLin++, 60 PSay "--------------------------------"
	//  @nLin++, 60 PSay "           Motorista"
	
	@nLin := nLin + 1
	@nLin++, 00 PSay "OBS: Os custos financeiros decorrentes do atraso na entrega serao cobrados do transportador"
	@nLin++
	@nLin++, 00 PSay AllTrim( GetMv("MV_XCFRT1") ) //Responsavel pela liberacao da Carta Frete
	@nLin++, 00 PSay AllTrim( GetMv("MV_XCFRT2") ) //Responsavel pela liberacao da Carta Frete
	nTotFrete := 0
    ntotpeso:=0
    nAdian := 0
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin := 6
	@nLin++, 70 PSay "VIA ADIANTAMENTO"
	@nLin, 00 PSay "TRANSPORTADOR -"
	@nLin++, 17 PSay cNomeTrans
	@nLin, 00 PSay "PLACA -"
	@nLin++, 09 PSay cPlaca
	@nLin, 00 PSay "CGC -"
	@nLin++,07 PSAY cCgc PICTURE IF(LEN(ALLTRIM(cCgc))>11,"@R !!9.999.999/9999-99","@R 999.999.999-99")
	@nLin++
	@nLin++, 00 PSay "NOTAS"
	@nLin++
    @nLin++, 00 PSay "CODIGO/LJ         NOME CLIENTE                NOTA   SERIE   DATA EMI     VALOR      FRETE      P.BRUTO"
    @nLin++, 00 Psay "---------   ------------------------------   ------  -----  ----------  ----------  ----------  -------"
	
	cAliasTC1->(dbGoTop())
	While !cAliasTC1->(EOF()) .and. !lAbortPrint
		@nLin  , 00 PSay cAliasTC1->(A1_COD + "/" + A1_LOJA)
		@nLin  , 12 PSay Subs(cAliasTC1->A1_NOME,1,30)
		@nLin  , 45 PSay cAliasTC1->F2_DOC
		@nLin  , 54 PSay cAliasTC1->F2_SERIE
		@nLin  , 61 PSay cAliasTC1->F2_EMISSAO 	PICTURE "@D"
		@nLin  , 72 PSay cAliasTC1->F2_VALBRUT	PICTURE "@E 999,999.99"
		@nLin  , 83 PSay cAliasTC1->F2_FRETE		PICTURE "@E 99,999.99"
                @nLin  , 93 PSay cAliasTC1->F2_PLIQUI         PICTURE "@E 999,999.99"

		nTotValor += cAliasTC1->F2_VALBRUT
		nTotFrete += cAliasTC1->F2_FRETE
        ntotpeso  += cAliasTC1->F2_PLIQUI

		nLin++ // Avanca a linha de impressao
		
		cAliasTC1->(dbSkip()) // Avanca o ponteiro do registro no arquivo
	Enddo

	@nLin++
	@nLin  , 00  PSay "TOTAIS:"
	@nLin  , 72 PSay nTotValor 	PICTURE "@E 999,999.99"
	@nLin  , 83 PSay NTOTFRETE  PICTURE "@E 99,999.99"
	@nLin  , 93 PSay NTOTPESO   PICTURE "@E 999,999.99"
	@nLin++
	@nLin  , 00 Psay "----------------------------------------------------------------------------------------------------------"
	@nLin++
	@nLin, 00 PSay "POSTO -"

	If ALLTRIM(cPosto) <> ""
		@nLin++, 08 PSay cPosto
		@nLin, 00 PSay "Estou ciente que efetuarei despesas de 25% no posto acima listado, no momento da troca da carta-frete, bem como dou quitaฦo dos valores "
		@nLin++
		@nLin, 00 PSay "abaixo descriminados."
		@nLin++
	Else
		@nLin++, 08 PSay "NAO HA POSTO"
	Endif
	@nLin++
	If MV_PAR07 = 1
		cTitFrete	:= GetSxeNum("SE2","E2_NUM",,1)
		@nLin, 00 PSay "TITULO FINANCEIRO -"
                cPreFrt         := AllTrim( GetMv("MV_XCFPRE") )
		@nLin, 21 PSay AllTrim(cPreFrt)
		@nLin++, 26 PSay AllTrim(cTitFrete)
	Endif
        @nLin++
        @nLin, 00 PSay "ADIANT.EFETUADO-"
	nAdian := (nTotFrete*nPerc)/100
	@nLin++, 17 PSay nAdian PICTURE "@E 99,999.99"
	@nLin++
	@nLin, 00 PSay "DATA DA EMISSAO -"
	@nLin++, 20 PSay dEmissao PICTURE "@D"
	@nLin, 00 PSay "DATA LIMITE PARA ENTREGA -"
	@nLin++, 27 PSay cDtEnt PICTURE "@D"
	
	@nLin++, 60 PSay "--------------------------------"
	@nLin++, 60 PSay "           Motorista"
	
	@nLin := nLin + 1
	@nLin++, 00 PSay "OBS: So aceitar esta carta frete, contendo o carimbo da Comafal mais um dos carimbos abaixo:"
	@nLin++
	@nLin++, 00 PSay AllTrim( GetMv("MV_XCFRT1") ) //Responsavel pela liberacao da Carta Frete
	@nLin++, 00 PSay AllTrim( GetMv("MV_XCFRT2") ) //Responsavel pela liberacao da Carta Frete
	nTotFrete := 0
    nTotpeso  := 0
	nAdian := 0
	nTotValor := 0
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin := 6
	
	@nLin++, 70 PSay "VIA COMAFAL (SALDO)"
	@nLin, 00 PSay "TRANSPORTADOR -"
	@nLin++, 17 PSay cNomeTrans
	@nLin, 00 PSay "PLACA -"
	@nLin++, 09 PSay cPlaca
	@nLin, 00 PSay "CGC -"
	@nLin++,07 PSAY cCgc PICTURE IF(LEN(ALLTRIM(cCgc))>11,"@R !!9.999.999/9999-99","@R 999.999.999-99")
	@nLin++
	@nLin++, 00 PSay "NOTAS"
	@nLin++
        @nLin++, 00 PSay "CODIGO/LJ         NOME CLIENTE                NOTA   SERIE   DATA EMI     VALOR      FRETE      P.BRUTO "
        @nLin++, 00 Psay "---------   ------------------------------   ------  -----  ----------  ----------  ----------  ------- "
        @nLin++

        cAliasTC1->(dbGoTop())
	While !cAliasTC1->(EOF()) .and. !lAbortPrint
		@nLin  , 00 PSay cAliasTC1->(A1_COD + "/" + A1_LOJA)
		@nLin  , 12 PSay Subs(cAliasTC1->A1_NOME,1,30)
		@nLin  , 45 PSay cAliasTC1->F2_DOC
		@nLin  , 54 PSay cAliasTC1->F2_SERIE
		@nLin  , 61 PSay cAliasTC1->F2_EMISSAO 	PICTURE "@D"
		@nLin  , 72 PSay cAliasTC1->F2_VALBRUT	PICTURE "@E 999,999.99"
		@nLin  , 83 PSay cAliasTC1->F2_FRETE		PICTURE "@E 99,999.99"
                @nLin  , 93 PSay cAliasTC1->F2_PLIQUI         PICTURE "@E 999,999.99" 
        nTotValor += cAliasTC1->F2_VALBRUT
		nTotFrete +=  cAliasTC1->F2_FRETE
        ntotpeso  +=  cAliasTC1->F2_PLIQUI

		nLin++ // Avanca a linha de impressao
		
		cAliasTC1->(dbSkip()) // Avanca o ponteiro do registro no arquivo
	Enddo

        @nLin++
        @nLin  , 00  PSay "TOTAIS:"
        @nLin  , 72 PSay NTOTVALOR  PICTURE "@E 99,999.99"
        @nLin  , 83 PSay NTOTFRETE  PICTURE "@E 99,999.99"
        @nLin  , 93 PSay NTOTPESO   PICTURE "@E 999,999.99"
        @nLin++
        @nLin++, 00 Psay "----------------------------------------------------------------------------------------------------------"
	@nLin++
	@nLin, 00 PSay "POSTO -"
	If ALLTRIM(cPosto) <> ""
		@nLin++, 08 PSay cPosto
	Else
		@nLin++, 08 PSay "NAO HA POSTO"
	Endif
	@nLin++
	@nLin++
	If MV_PAR07 = 1
		cTitFrete	:= GetSxeNum("SE2","E2_NUM",,1)
		@nLin, 00 PSay "TITULO FINANCEIRO -"
                cPreFrt         := AllTrim( GetMv("MV_XCFPRE") )
		@nLin, 21 PSay AllTrim(cPreFrt)
		@nLin++, 26 PSay AllTrim(cTitFrete)
	Endif
        @nLin++
        @nLin, 00 PSay "TOTAL DO FRETE -"
	@nLin++, 17 PSay nTotFrete PICTURE "@E 99,999.99"
	@nLin, 00 PSay "ADIANTAMENTO   -"
	nAdian := (nTotFrete*nPerc)/100
	@nLin++, 17 PSay nAdian PICTURE "@E 99,999.99"
	@nLin++
	@nLin, 00 PSay "SALDO A PAGAR -"
	@nLin++, 17 PSay nTotFrete - nAdian PICTURE "@E 99,999.99"
	@nLin++
	@nLin, 00 PSay "DATA DA EMISSAO -"
	@nLin++, 20 PSay dEmissao PICTURE "@D"
	@nLin, 00 PSay "DATA LIMITE PARA ENTREGA -"
	@nLin++, 27 PSay cDtEnt PICTURE "@D"
	
	@nLin++, 60 PSay "--------------------------------"
	@nLin++, 60 PSay "           Motorista"
	
	@nLin := nLin + 1
	@nLin++, 00 PSay "OBS: So aceitar esta carta frete, contendo o carimbo da Comafal mais um dos carimbos abaixo:"
	@nLin++
	@nLin++, 00 PSay AllTrim( GetMv("MV_XCFRT1") ) //Responsavel pela liberacao da Carta Frete
	@nLin++, 00 PSay AllTrim( GetMv("MV_XCFRT2") ) //Responsavel pela liberacao da Carta Frete
	nTotFin := 	nTotFrete
	nAdFin  := 	nAdian
	nTotFrete := 0
	nAdian := 0
EndDo

If MV_PAR07 = 1
	ExecBlock("GeraFin",.F.,.F.)
Endif

cAliasTC1->( dbCloseArea() )

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Finaliza a execucao do relatorio...                                 ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

SET DEVICE TO SCREEN

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Se impressao em disco, chama o gerenciador de impressao...          ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

If aReturn[5]==1
	dbCommitAll()
	SET PRINTER TO
	OurSpool(wnrel)
Endif

MS_FLUSH()

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณVALIDPERG บ Autor ณ AP5 IDE            บ Data ณ  12/03/02   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescrio ณ Verifica a existencia das perguntas criando-as caso seja   บฑฑ
ฑฑบ          ณ necessario (caso nao existam).                             บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Programa principal                                         บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
Static Function ValidPerg
Local _sAlias 	:= Alias()
Local aRegs 	:= {}
Local i,j		:= 0

dbSelectArea("SX1")
dbSetOrder(1)
cPerg := PADR(cPerg,6)

// Grupo/Ordem/Pergunta/Variavel/Tipo/Tamanho/Decimal/Presel/GSC/Valid/Var01/Def01/Cnt01/Var02/Def02/Cnt02/Var03/Def03/Cnt03/Var04/Def04/Cnt04/Var05/Def05/Cnt05
aAdd(aRegs,{cPerg,"01","TRANSPORTADORA    ","TRANSPORTADORA    ","TRANSPORTADORA    ","mv_ch1","C",6,0,0,"G","","MV_PAR01","","","","","","","","","","","","","","","","","","","","","","","","","SA4"})
aAdd(aRegs,{cPerg,"02","DATA DA EMISSAO   ","DATA DA EMISSAO   ","DATA DA EMISSAO   ","mv_ch2","D",8,0,0,"G","","MV_PAR02","","","","","","","","","","","","","","","","","","","","","","","","",""})
aAdd(aRegs,{cPerg,"03","PLACA             ","PLACA             ","PLACA             ","mv_ch3","C",7,0,0,"G","","MV_PAR03","","","","","","","","","","","","","","","","","","","","","","","","",""})
aAdd(aRegs,{cPerg,"04","DATA ENTREGA      ","DATA ENTREGA      ","DATA ENTREGA      ","mv_ch4","D",8,0,0,"G","","MV_PAR04","","","","","","","","","","","","","","","","","","","","","","","","",""})
aAdd(aRegs,{cPerg,"05","PERCENTUAL A PAGAR","PERCENTUAL A PAGAR","PERCENTUAL A PAGAR","mv_ch5","N",3,0,0,"G","","MV_PAR05","","","","","","","","","","","","","","","","","","","","","","","","",""})
aAdd(aRegs,{cPerg,"06","POSTO			  ?","POSTO			   ?","POSTO			?","mv_ch6","C",22,0,0,"G","","MV_PAR06","","","","","","","","","","","","","","","","","","","","","","","","",""})
aAdd(aRegs,{cPerg,"07","GERA FINANCEIRO    ","GERA FINANCEIRO   ","GERA FINANCEIRO   ","mv_ch7","N",1,0,0,"C","","MV_PAR07","SIM","","","","","NAO","","","","","","","","","","","","","","","","","","",""})
aAdd(aRegs,{cPerg,"08","FORNECEDOR         ","FORNECEDOR        ","FORNECEDOR        ","mv_ch8","C",6,0,0,"G","","MV_PAR08","","","","","","","","","","","","","","","","","","","","","","","","","SA2"})
aAdd(aRegs,{cPerg,"09","LOJA-FORNECEDOR    ","LOJA-FORNECEDOR   ","LOJA-FORNECEDOR   ","mv_ch9","C",2,0,0,"G","","MV_PAR09","","","","","","","","","","","","","","","","","","","","","","","","",""})
//
For i:=1 to Len(aRegs)
	If !dbSeek(cPerg+aRegs[i,2])
		RecLock("SX1",.T.)
		For j:=1 to FCount()
			If j <= Len(aRegs[i])
				FieldPut(j,aRegs[i,j])
			Endif
		Next
		MsUnlock()
	Endif
Next

dbSelectArea(_sAlias)
Return

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณGeraFin   บ Autor ณ Daniel             บ Data ณ  12/07/02   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescricao ณ Gera financeiro apartir das informacoes do relatorio       บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ AP5 IDE                                                    บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
/*/
User Function GeraFin()
Local cTitFrete	:= ""
Local cPre		:= ""
Local cNomBen	:= ""
Local cFornec	:= ""
Local cLoja 	:= ""
Local cPreFrt	:= ""
Local cNomFor	:= ""

cTitFrete	:= GetSxeNum("SE2","E2_NUM",,1)
cPreFrt		:= AllTrim( GetMv("MV_XCFPRE") )
cFornec := MV_PAR08
cLoja	:= MV_PAR09

calias:=alias()
crec:=recno()

If SA2->( dbSeek( xFilial("SA2") +cFornec+cLoja, .F.) )
	cNomFor := SA2->A2_NREDUZ
	SE2->(dbSetOrder( 1 ))
	If !SE2->(dbSeek(xFilial("SE2")+cPreFrt+cTitFrete+"CFT"+cFornec+cLoja)) .and. nTotFin > 0
		RecLock("SE2",.T.)
		SE2->E2_FILIAL 	:= xFilial("SE2")
		SE2->E2_PREFIXO := AllTrim( GetMv("MV_XCFPRE") ) //PREFIXO
		SE2->E2_NUM 	:= cTitFrete
		SE2->E2_TIPO    := "CFT" //TIPO
		SE2->E2_NATUREZ := AllTrim( GetMv("MV_XCFNAT") ) //NATUREZA
		SE2->E2_FORNECE := cFornec
		SE2->E2_NOMFOR	:= cNomFor
		SE2->E2_LOJA	:= cLoja
		SE2->E2_EMISSAO	:= dEmissao
		SE2->E2_VENCTO	:= cDtEnt
		SE2->E2_VENCREA	:= cDtEnt
		SE2->E2_VALOR	:= nTotFin
		SE2->E2_RATEIO	:= 'N'
		SE2->E2_FLUXO	:= 'S'
		SE2->E2_DESDOBR	:= 'N'
		SE2->E2_VALOR	:= nTotFin
		SE2->E2_MOEDA	:= 1
		SE2->E2_ORIGEM	:= 'FINA050'
		If nAdFin > 0
			SE2->E2_BAIXA	:= dEmissao
			SE2->E2_EMIS1	:= dEmissao
			SE2->E2_BCOPAG	:= AllTrim( GetMv("MV_XCFBCO") ) //BANCO
			SE2->E2_VALLIQ	:= nTotFin - nAdFin
			SE2->E2_SALDO	:= nTotFin - nAdFin
			SE2->E2_VENCORI	:= cDtEnt
		Else
			SE2->E2_SALDO	:= nTotFin
		Endif
		
		Replace E2_NUM with cTitFrete
		ConfirmSX8()
		
		SE2->(MsUnLock())
		
		SA2->( dbSetOrder( 1 ) )
		If SA2->( dbSeek( xFilial("SA2") +cFornec+cLoja, .F.) )
			cNomBen := SA2->A2_NOME
			substring(cNomBen := SA2->A2_NOME,1,25)
			SE5->( dbSetOrder( 1 ) )
        
                        If !SE5->(dbSeek(xFilial("SE5")+cPre+cTitFrete+""+"CFT"+cFornec+cLoja+"01")) .and. nAdFin > 0
				RecLock("SE5",.T.)
				SE5->E5_FILIAL 	:= xFilial("SE5")
				SE5->E5_DATA  	:= dEmissao
                SE5->E5_TIPO    := "CFT" // AllTrim( GetMv("MV_XCFTIP") ) //TIPO
				SE5->E5_VALOR  	:= nTotFin - nAdFin
				SE5->E5_NATUREZ := AllTrim( GetMv("MV_XCFNAT") ) //NATUREZA
				SE5->E5_BANCO  	:= AllTrim( GetMv("MV_XCFBCO") ) //BANCO
				SE5->E5_AGENCIA	:= AllTrim( GetMv("MV_XCFAG") ) //BANCO
				SE5->E5_CONTA	:= AllTrim( GetMv("MV_XCFCTO") ) //BANCO
				SE5->E5_RECPAG  := 'P'
				SE5->E5_BENEF	:= cNomBen
				SE5->E5_NUMERO 	:= cTitFrete
				SE5->E5_HISTOR  := 'AD. FRETE'
				SE5->E5_TIPODOC := 'BA'
				SE5->E5_VLMOED2 := nTotFin - nAdFin
				SE5->E5_LA		:= 'N'
				SE5->E5_PREFIXO	:= AllTrim( GetMv("MV_XCFPRE") ) //PREFIXO
				SE5->E5_CLIFOR	:= cFornec
				SE5->E5_LOJA	:= cLoja
				SE5->E5_DTDIGIT	:= dEmissao
				SE5->E5_MOTBX	:= 'NOR'
                                SE5->E5_SEQ     := '01'
				SE5->E5_DTDISPO := dEmissao
			Endif
        
        	EndIf
                SE5->(dbCloseArea())
	Else
		MsgAlert("Erro ao cadastrar Titulo")
	Endif
Else
	MsgAlert("Fornecedor invalido")
Endif
nAdFin := 0

Dbselectarea(calias)
Goto crec
Return

