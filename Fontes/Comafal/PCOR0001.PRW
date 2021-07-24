#include "rwmake.ch"    
#INCLUDE "topconn.ch"

#IFNDEF WINDOWS
	#DEFINE PSAY SAY
#ENDIF

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Programa  ³  PCOR0001  ³ Autor ³   Valrélio Gomes    ³ Data ³ 23/10/08 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Relação do Orçamento                                       ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/
User Function PCOR0001()

SetPrvt("CBTXT,CBCONT,NORDEM,ALFA,Z,M")
SetPrvt("TAMANHO,LIMITE,TITULO,CDESC1,CDESC2,CDESC3")
SetPrvt("ARETURN,NOMEPROG,CPERG,NLASTKEY,LCONTINUA,NLIN")
SetPrvt("WNREL,M_PAG,CSTRING,CABEC1,CABEC2,MPAG")
SetPrvt("NTOTAL,NSUBTOTAL,CSETOR,NTOT,")
SetPrvt("nPer01,nPer02,nPer03,nPer04,nPer05,nPer06,nPer07,nPer08,nPer09,nPer10,nPer11,nPer12")

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Define Variaveis Ambientais                                                ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Variaveis utilizadas para parametros                                       ³
//³ mv_par01             // Orcamento                                          ³
//³ mv_par02             // Versao                                             ³
//³ mv_par03             // De  Conta Orcamentaria                             ³
//³ mv_par04             // Ate Conta Orcamentaria                             ³
//³ mv_par05             // Da  Filial                                         ³
//³ mv_par06             // Ate Filial                                         ³
//³ mv_par07             // Visao   1-Item X C.Orc  2-C.Orc X Item             ³
//³ mv_par08             // Nivel de Visao     1-Grupo 2-Sub-Grupo 3-Analitico ³
//³ mv_par09             // Exibe Zerados      1-Sim   2-Nao       3-Ambos     ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
CbTxt  := CbCont := ""
nOrdem := Alfa   := Z  := M := 0

tamanho  := "G"
limite   := 254
titulo   := PADC("Orcamento COMAFAL ",74)
cDesc1   := PADC("",74)
cDesc2   := PADC("",74)
cDesc3   := ""
aReturn  := { "Financeiro", 1,"Administracao", 1, 2, 1,"",1 }
nomeprog := "PCOR0001"
cPerg    := "PCO001"
nLastKey := 0
lContinua:= .T.
nLin     := 9
wnrel    := "PCOR0001_"+cUserName

M_PAG    := 1

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Verifica as perguntas selecionadas, busca o padrao da Nfiscal           ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

Pergunte(cPerg,.F.)               // Pergunta no SX1

cString := "AK2"

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Envia controle para a funcao SETPRINT                        ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

wnrel:=SetPrint(cString,wnrel,cPerg,Titulo,cDesc1,cDesc2,cDesc3,.T.)

If nLastKey == 27
	Return
Endif

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Verifica Posicao do Formulario na Impressora                 ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
SetDefault(aReturn,cString)

If nLastKey == 27
	Return
Endif

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Inicializa  de variaveis                                  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
Processa( {|| mapa() } )// Substituido pelo assistente de conversao do AP5 IDE em 06/12/01 ==> Processa( {|| Execute(mapa) } )
Return

// Substituido pelo assistente de conversao do AP5 IDE em 06/12/01 ==> Function mapa
Static Function mapa()

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Inicializa  de variaveis                                  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
aEstrut := {}
Aadd( aEstrut, { "FILIAL"  ,"C", Getsx3cache("AK2_FILIAL","X3_TAMANHO"),0 } )
Aadd( aEstrut, { "ORCAME"  ,"C", Getsx3cache("AK2_ORCAME","X3_TAMANHO"),0 } )
Aadd( aEstrut, { "CO"      ,"C", Getsx3cache("AK2_CO"    ,"X3_TAMANHO"),0 } )
Aadd( aEstrut, { "ITCTB"   ,"C", Getsx3cache("AK2_ITCTB" ,"X3_TAMANHO"),0 } )
Aadd( aEstrut, { "CC"      ,"C", Getsx3cache("AK2_CC"    ,"X3_TAMANHO"),0 } )
Aadd( aEstrut, { "VERSAO"  ,"C", 06, 0 } )
Aadd( aestrut, { "PERIODO" ,"D", 08, 0 } )
Aadd( aEstrut, { "TPSALDO" ,"C", 02, 0 } )
Aadd( aEstrut, { "VALOR01" ,"N", 14, 2 } )
Aadd( aEstrut, { "VALOR02" ,"N", 14, 2 } )
Aadd( aestrut, { "VALOR03" ,"N", 14, 2 } )
Aadd( aEstrut, { "VALOR04" ,"N", 14, 2 } )
Aadd( aEstrut, { "VALOR05" ,"N", 14, 2 } )
Aadd( aEstrut, { "VALOR06" ,"N", 14, 2 } )
Aadd( aEstrut, { "VALOR07" ,"N", 14, 2 } )
Aadd( aestrut, { "VALOR08" ,"N", 14, 2 } )
Aadd( aEstrut, { "VALOR09" ,"N", 14, 2 } )
Aadd( aEstrut, { "VALOR10" ,"N", 14, 2 } )
Aadd( aEstrut, { "VALOR11" ,"N", 14, 2 } )
Aadd( aEstrut, { "VALOR12" ,"N", 14, 2 } )

cArqTmp := CriaTrab( aEstrut, .t. )
use ( cArqTmp ) alias TMP new
cfiltro := ""
cChave  := "FILIAL+ORCAME+CO+TPSALDO+VERSAO"
cArqInd := CriaTrab( nil,.f. )
IndRegua( "TMP", cArqInd, cChave, , cFiltro, "Indexando Arquivo de Trabalho..." )


//----------------------------------|| Gera o Saldo Orçado dentro de um período informado. ||---------------------------------------------------
if !Empty(mv_par03) .Or. !Empty(mv_par04)
	
	lAsTop	:= .T.
	cQuery := "SELECT * FROM " + RetSqlName("AK2")+" AK2 "
	cQuery += "WHERE AK2.AK2_ORCAME= '" + MV_PAR01 + "' AND "
	cQuery += "AK2.AK2_VERSAO = '" + MV_PAR02 + "' AND "
	cQuery += "AK2.AK2_CO    Between '" + MV_PAR03 + "' AND '" + MV_PAR04 + "' AND "
	cQuery += "AK2.AK2_FILIAL Between '" + MV_PAR05 + "' AND '" + MV_PAR06 + "' AND "
	
	If mv_par09==1
		cQuery += "      AK2.AK2_VALOR <= 0 AND "
	ElseIf mv_par09==2
		cQuery += "      AK2.AK2_VALOR > 0 AND "
	EndIf
	
	cQuery += " D_E_L_E_T_<>'*' "
	cQuery += " ORDER BY AK2.AK2_FILIAL,AK2.AK2_CO,AK2.AK2_CLASSE,AK2.AK2_OPER,AK2.AK2_CHAVE"
	
	dbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery), 'QAK2', .F., .T.)
	aStru := AK2->(dbStruct())
	
	For ni := 1 to Len(aStru)
		If aStru[ni,2] != 'C'
			TCSetField('QAK2', aStru[ni,1], aStru[ni,2],aStru[ni,3],aStru[ni,4])
		Endif
	Next
	
EndIf

QAK2->( dbGoTop() )

Do While QAK2->(!Eof())
	
	TMP->(DbSeek(QAK2->AK2_FILIAL+QAK2->AK2_ORCAME+QAK2->AK2_CO+"OR"+QAK2->AK2_VERSAO))
	
	If TMP->(!Found())
		RecLock( "TMP", .T. )
		
		TMP->FILIAL  := QAK2->AK2_FILIAL
		TMP->ORCAME  := QAK2->AK2_ORCAME
		TMP->CO      := QAK2->AK2_CO
		TMP->ITCTB   := QAK2->AK2_ITCTB
		TMP->CC      := QAK2->AK2_CC
		TMP->VERSAO  := QAK2->AK2_VERSAO
		TMP->PERIODO := QAK2->AK2_PERIODO
		TMP->TPSALDO := "OR"
		TMP->VALOR01 := IIf(month(QAK2->AK2_PERIODO)==1 ,QAK2->AK2_VALOR,0)
		TMP->VALOR02 := IIf(month(QAK2->AK2_PERIODO)==2 ,QAK2->AK2_VALOR,0)
		TMP->VALOR03 := IIf(month(QAK2->AK2_PERIODO)==3 ,QAK2->AK2_VALOR,0)
		TMP->VALOR04 := IIf(month(QAK2->AK2_PERIODO)==4 ,QAK2->AK2_VALOR,0)
		TMP->VALOR05 := IIf(month(QAK2->AK2_PERIODO)==5 ,QAK2->AK2_VALOR,0)
		TMP->VALOR06 := IIf(month(QAK2->AK2_PERIODO)==6 ,QAK2->AK2_VALOR,0)
		TMP->VALOR07 := IIf(month(QAK2->AK2_PERIODO)==7 ,QAK2->AK2_VALOR,0)
		TMP->VALOR08 := IIf(month(QAK2->AK2_PERIODO)==8 ,QAK2->AK2_VALOR,0)
		TMP->VALOR09 := IIf(month(QAK2->AK2_PERIODO)==9 ,QAK2->AK2_VALOR,0)
		TMP->VALOR10 := IIf(month(QAK2->AK2_PERIODO)==10,QAK2->AK2_VALOR,0)
		TMP->VALOR11 := IIf(month(QAK2->AK2_PERIODO)==11,QAK2->AK2_VALOR,0)
		TMP->VALOR12 := IIf(month(QAK2->AK2_PERIODO)==12,QAK2->AK2_VALOR,0)
		MsUnLock()
	Else
		RecLock( "TMP", .F. )
		
		TMP->VALOR01 += IIf(month(QAK2->AK2_PERIODO)==1 ,QAK2->AK2_VALOR,0)
		TMP->VALOR02 += IIf(month(QAK2->AK2_PERIODO)==2 ,QAK2->AK2_VALOR,0)
		TMP->VALOR03 += IIf(month(QAK2->AK2_PERIODO)==3 ,QAK2->AK2_VALOR,0)
		TMP->VALOR04 += IIf(month(QAK2->AK2_PERIODO)==4 ,QAK2->AK2_VALOR,0)
		TMP->VALOR05 += IIf(month(QAK2->AK2_PERIODO)==5 ,QAK2->AK2_VALOR,0)
		TMP->VALOR06 += IIf(month(QAK2->AK2_PERIODO)==6 ,QAK2->AK2_VALOR,0)
		TMP->VALOR07 += IIf(month(QAK2->AK2_PERIODO)==7 ,QAK2->AK2_VALOR,0)
		TMP->VALOR08 += IIf(month(QAK2->AK2_PERIODO)==8 ,QAK2->AK2_VALOR,0)
		TMP->VALOR09 += IIf(month(QAK2->AK2_PERIODO)==9 ,QAK2->AK2_VALOR,0)
		TMP->VALOR10 += IIf(month(QAK2->AK2_PERIODO)==10,QAK2->AK2_VALOR,0)
		TMP->VALOR11 += IIf(month(QAK2->AK2_PERIODO)==11,QAK2->AK2_VALOR,0)
		TMP->VALOR12 += IIf(month(QAK2->AK2_PERIODO)==12,QAK2->AK2_VALOR,0)
		
		MsUnLock()
		
	EndIf

	If TMP->(!DbSeek(QAK2->AK2_FILIAL+QAK2->AK2_ORCAME+QAK2->AK2_CO+"QT"+QAK2->AK2_VERSAO))

		RecLock( "TMP", .T. )
		TMP->FILIAL  := QAK2->AK2_FILIAL
		TMP->ORCAME  := QAK2->AK2_ORCAME
		TMP->CO      := QAK2->AK2_CO
		TMP->ITCTB   := QAK2->AK2_ITCTB
		TMP->CC      := QAK2->AK2_CC
		TMP->VERSAO  := QAK2->AK2_VERSAO
		TMP->PERIODO := QAK2->AK2_PERIODO
		
		TMP->TPSALDO := "QT"        

		TMP->VALOR01 := 0  ;	TMP->VALOR02 := 0
		TMP->VALOR03 := 0  ;	TMP->VALOR04 := 0
		TMP->VALOR05 := 0  ;	TMP->VALOR06 := 0
		TMP->VALOR07 := 0  ;	TMP->VALOR08 := 0
		TMP->VALOR09 := 0  ;	TMP->VALOR10 := 0
		TMP->VALOR11 := 0  ; TMP->VALOR12 := 0
		MsUnLock()
    EndIf

	QAK2->(DbSkip())
	
Enddo                                     
//----------------------------------|| Gera o Saldo Quitado dentro de um período informado. ||---------------------------------------------------
if !Empty(mv_par03) .Or. !Empty(mv_par04)
	
	dIni := Ctod("01/01/"+SubStr(mv_par01,len(Alltrim(mv_par01))-3,4))
	dFim := Ctod("31/12/"+SubStr(mv_par01,len(Alltrim(mv_par01))-3,4))
	
	lAsTop	:= .T.
	cQuery1 := " SELECT * FROM " + RetSqlName("SE5")+" SE5 "
	cQuery1 += " WHERE SE5.E5_RECPAG= 'P' AND "
	cQuery1 += " SE5.E5_FILIAL Between '" + MV_PAR05 + "' AND '" + MV_PAR06 + "' AND "
	cQuery1 += " SE5.E5_DATA Between '" + dTos(dIni) + "' AND '" + dTos(dFim) + "' AND "
	cQuery1 += " D_E_L_E_T_<>'*' "
	cQuery1 += " ORDER BY SE5.E5_FILIAL,SE5.E5_DATA"
	
	dbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery1), 'QSE5', .F., .T.)
	aStru1 := SE5->(dbStruct())
	
	For ni := 1 to Len(aStru1)
		If aStru1[ni,2] != 'C'
			TCSetField('QSE5', aStru1[ni,1], aStru1[ni,2],aStru1[ni,3],aStru1[ni,4])
		Endif
	Next
EndIf

QSE5->( dbGoTop() )

Do While QSE5->(!Eof())

// --- Elimina valores extornados ---
	DbSelectArea("SE5")
	SE5->(DbSetOrder(7))
	SE5->(DbSeek(xFilial("SE5")+QSE5->(E5_PREFIXO+E5_NUMERO+E5_PARCELA+E5_TIPO+E5_CLIFOR+E5_LOJA+E5_SEQ)))
	
	lPula := .T.
	
	Do While SE5->(!Eof()) .And. QSE5->(E5_PREFIXO+E5_NUMERO+E5_PARCELA+E5_TIPO+E5_CLIFOR+E5_LOJA+E5_SEQ)==SE5->(E5_PREFIXO+E5_NUMERO+E5_PARCELA+E5_TIPO+E5_CLIFOR+E5_LOJA+E5_SEQ) .And. lpula==.T.
   	If SE5->E5_RECPAG=="R"
	      lPula := .F.
	   EndIf    
	   SE5->(DbSkip())
	   Loop
	EndDo

	If lPula==.F.
	   QSE5->(DbSkip())
	   Loop
	EndIf
// ----------------------------------
	DbSelectArea("SED")
	SED->(DbSetOrder(1))
	SED->(DbSeek(xFilial("SED")+Alltrim(QSE5->E5_NATUREZ)))  // voltar filial

	If SED->ED_ORCAME < mv_par03 .Or. SED->ED_ORCAME > mv_par04
	   QSE5->(DbSkip())
	   Loop
	EndIf

	If Empty(SED->ED_ORCAME)
	   QSE5->(DbSkip())
	   Loop
	EndIf

	DbSelectArea("AK5")
	AK5->(DbSetOrder(1))
			
	If AK5->(!DbSeek(AK5->AK5_FILIAL+Alltrim(SED->ED_ORCAME)))
	   QSE5->(DbSkip())
	   Loop
    EndIf

	TMP->(DbSeek(QSE5->E5_FILIAL+MV_PAR01+SED->ED_ORCAME+"QT"+MV_PAR02))   //FILIAL+ORCAME+CO+VERSAO+TPSALDO 

	If TMP->(!Found())
		RecLock( "TMP", .T. )
		
		TMP->FILIAL  := QSE5->E5_FILIAL
		TMP->CC      := QSE5->E5_CCC
		TMP->PERIODO := QSE5->E5_DATA

		TMP->VERSAO  := MV_PAR02            //--
		TMP->ORCAME  := MV_PAR01            //--
		TMP->CO      := SED->ED_ORCAME      //--
		TMP->TPSALDO := "QT"                //--


		TMP->VALOR01 := IIf(month(QSE5->E5_DATA)==1 ,QSE5->E5_VALOR,0)
		TMP->VALOR02 := IIf(month(QSE5->E5_DATA)==2 ,QSE5->E5_VALOR,0)
		TMP->VALOR03 := IIf(month(QSE5->E5_DATA)==3 ,QSE5->E5_VALOR,0)
		TMP->VALOR04 := IIf(month(QSE5->E5_DATA)==4 ,QSE5->E5_VALOR,0)
		TMP->VALOR05 := IIf(month(QSE5->E5_DATA)==5 ,QSE5->E5_VALOR,0)
		TMP->VALOR06 := IIf(month(QSE5->E5_DATA)==6 ,QSE5->E5_VALOR,0)
		TMP->VALOR07 := IIf(month(QSE5->E5_DATA)==7 ,QSE5->E5_VALOR,0)
		TMP->VALOR08 := IIf(month(QSE5->E5_DATA)==8 ,QSE5->E5_VALOR,0)
		TMP->VALOR09 := IIf(month(QSE5->E5_DATA)==9 ,QSE5->E5_VALOR,0)
		TMP->VALOR10 := IIf(month(QSE5->E5_DATA)==10,QSE5->E5_VALOR,0)
		TMP->VALOR11 := IIf(month(QSE5->E5_DATA)==11,QSE5->E5_VALOR,0)
		TMP->VALOR12 := IIf(month(QSE5->E5_DATA)==12,QSE5->E5_VALOR,0)
		MsUnLock()
	Else
		RecLock( "TMP", .F. )

		TMP->VALOR01 += IIf(month(QSE5->E5_DATA)==1 ,QSE5->E5_VALOR,0)
		TMP->VALOR02 += IIf(month(QSE5->E5_DATA)==2 ,QSE5->E5_VALOR,0)
		TMP->VALOR03 += IIf(month(QSE5->E5_DATA)==3 ,QSE5->E5_VALOR,0)
		TMP->VALOR04 += IIf(month(QSE5->E5_DATA)==4 ,QSE5->E5_VALOR,0)
		TMP->VALOR05 += IIf(month(QSE5->E5_DATA)==5 ,QSE5->E5_VALOR,0)
		TMP->VALOR06 += IIf(month(QSE5->E5_DATA)==6 ,QSE5->E5_VALOR,0)
		TMP->VALOR07 += IIf(month(QSE5->E5_DATA)==7 ,QSE5->E5_VALOR,0)
		TMP->VALOR08 += IIf(month(QSE5->E5_DATA)==8 ,QSE5->E5_VALOR,0)
		TMP->VALOR09 += IIf(month(QSE5->E5_DATA)==9 ,QSE5->E5_VALOR,0)
		TMP->VALOR10 += IIf(month(QSE5->E5_DATA)==10,QSE5->E5_VALOR,0)
		TMP->VALOR11 += IIf(month(QSE5->E5_DATA)==11,QSE5->E5_VALOR,0)
		TMP->VALOR12 += IIf(month(QSE5->E5_DATA)==12,QSE5->E5_VALOR,0)
		
		MsUnLock()
		
	EndIf

	If TMP->(!DbSeek(QSE5->E5_FILIAL+MV_PAR01+SED->ED_ORCAME+"OR"+MV_PAR02))    

		RecLock( "TMP", .T. )
		
		TMP->FILIAL  := QSE5->E5_FILIAL
		TMP->CC      := QSE5->E5_CCC
		TMP->PERIODO := QSE5->E5_DATA

		TMP->VERSAO  := MV_PAR02            //--
		TMP->ORCAME  := MV_PAR01            //--
		TMP->CO      := SED->ED_ORCAME      //--
		TMP->TPSALDO := "OR"        

		TMP->VALOR01 := 0  ;	TMP->VALOR02 := 0
		TMP->VALOR03 := 0  ;	TMP->VALOR04 := 0
		TMP->VALOR05 := 0  ;	TMP->VALOR06 := 0
		TMP->VALOR07 := 0  ;	TMP->VALOR08 := 0
		TMP->VALOR09 := 0  ;	TMP->VALOR10 := 0
		TMP->VALOR11 := 0  ; TMP->VALOR12 := 0
		MsUnLock()
    EndIf
	 
	QSE5->(DbSkip())
	
Enddo                                     

//-----------------------------------|| Fim do processo de captura de Dados para composição dos Saldos ||----------------------------------------

cCancel := "***** CANCELADO PELO OPERADOR *****"
m_pag   := 1

#IFDEF WINDOWS
	RptStatus({|| RptDetail()})// Substituido pelo assistente de conversao do AP5 IDE em 06/12/01 ==>    RptStatus({|| Execute(RptDetail)})
	Return
	// Substituido pelo assistente de conversao do AP5 IDE em 06/12/01 ==>    Function RptDetail
	Static Function RptDetail()
#ENDIF

SetRegua( TMP->( LastRec() ) )    //Ajusta numero de elementos da regua de relatorios

nLin    := 70
cfiltro := ""
cChave  := Iif(mv_par07==1,"FILIAL+ORCAME+CO+TPSALDO+VERSAO","FILIAL+ORCAME+CO+TPSALDO+VERSAO")   //Valrélio analisar tipos de indices
cArqInd := CriaTrab( nil,.f. )

IndRegua( "TMP", cArqInd, cChave, , cFiltro, "Indexando Arquivo de Trabalho..." )

nFlag   := 0

Cabec1  := "CONTA ORCAMENTARIA  "
Cabec2  := "SALDO                                                  JANEIRO     FEVEREIRO         MARCO         ABRIL          MAIO         JUNHO         JULHO        AGOSTO      SETEMBRO       OUTUBRO      NOVEMBRO      DEZEMBRO"
//            -123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789
//            0---------1---------2---------3---------4---------5---------6---------7---------8---------9---------0---------1---------2---------3---------4---------5---------6---------7---------8---------9---------0---------1---------2---------3---------4---------5---------6---------7---------8---------9---------0---------1
//            300101001-IMPRESSAO DE CONSUMO MATERIAL ESCRITO   99.999.999,99 99.999.999,99 99.999.999,99 99.999.999,99 99.999.999,99 99.999.999,99 99.999.999,99 99.999.999,99 99.999.999,99 99.999.999,99 99.999.999,99 99.999.999,99 99.999.999,99

nPer01:=nPer02:=nPer03:=nPer04:=nPer05:=nPer06:=nPer07:=nPer08:=nPer09:=nPer10:=nPer11:=nPer12:= 0

DbSelectArea("TMP")
TMP->( dbGotop() )

Do While TMP->( !Eof() )
	
	If nLin >= 58
		nLin := Cabec(titulo,cabec1,cabec2,nomeprog,tamanho,15) + 2 //Impressao do cabecalho
	Endif
	
	cCO := TMP->CO
	AK5->(DbSeek(AK5->AK5_FILIAL+TMP->CO))
	
	@ nlin, 001 pSay Alltrim(TMP->CO)+"-"+AK5->AK5_DESCRI
	
	nlin := nlin + 1
	
	Do While TMP->(!Eof()) .And. TMP->CO == cCO
		
		cSaldo := TMP->TPSALDO
		
		@ nlin, 001 pSay Iif(TMP->TPSALDO=="OR","OR-SALDO ORCADO","QT-SALDO QUITADO")
		
		Do While TMP->(!Eof()) .And. TMP->CO == cCO .And. TMP->TPSALDO == cSaldo
			
			@ nlin, 040 pSay Transform(TMP->VALOR01,"@E 999,999,999.99")
			@ nlin, 056 pSay Transform(TMP->VALOR02,"@E 999,999,999.99")
			@ nlin, 072 pSay Transform(TMP->VALOR03,"@E 999,999,999.99")
			@ nlin, 088 pSay Transform(TMP->VALOR04,"@E 999,999,999.99")
			@ nlin, 104 pSay Transform(TMP->VALOR05,"@E 999,999,999.99")
			@ nlin, 120 pSay Transform(TMP->VALOR06,"@E 999,999,999.99")
			@ nlin, 136 pSay Transform(TMP->VALOR07,"@E 999,999,999.99")
			@ nlin, 152 pSay Transform(TMP->VALOR08,"@E 999,999,999.99")
			@ nlin, 168 pSay Transform(TMP->VALOR09,"@E 999,999,999.99")
			@ nlin, 184 pSay Transform(TMP->VALOR10,"@E 999,999,999.99")
			@ nlin, 200 pSay Transform(TMP->VALOR11,"@E 999,999,999.99")
			@ nlin, 216 pSay Transform(TMP->VALOR12,"@E 999,999,999.99")
			//      		@ nlin, 242 pSay Transform(TMP->VALOR12,"@E 999,999,999.99")
			
			If TMP->TPSALDO=="OR"
				
				nPer01 := TMP->VALOR01   ;			   nPer02 := TMP->VALOR02
				nPer03 := TMP->VALOR03   ; 		   nPer04 := TMP->VALOR04
				nPer05 := TMP->VALOR05   ;			   nPer06 := TMP->VALOR06
				nPer07 := TMP->VALOR07   ;			   nPer08 := TMP->VALOR08
				nPer09 := TMP->VALOR09   ;			   nPer10 := TMP->VALOR10
				nPer11 := TMP->VALOR11   ;			   nPer12 := TMP->VALOR12
				
			Else
				
				nPer01 := Round((TMP->VALOR01*100/nPer01),2)
				nPer02 := Round((TMP->VALOR02*100/nPer02),2)
				nPer03 := Round((TMP->VALOR03*100/nPer03),2)
				nPer04 := Round((TMP->VALOR04*100/nPer04),2)
				nPer05 := Round((TMP->VALOR05*100/nPer05),2)
				nPer06 := Round((TMP->VALOR06*100/nPer06),2)
				nPer07 := Round((TMP->VALOR07*100/nPer07),2)
				nPer08 := Round((TMP->VALOR08*100/nPer08),2)
				nPer09 := Round((TMP->VALOR09*100/nPer09),2)
				nPer10 := Round((TMP->VALOR10*100/nPer10),2)
				nPer11 := Round((TMP->VALOR11*100/nPer11),2)
				nPer12 := Round((TMP->VALOR12*100/nPer12),2)
				
			EndIf
			TMP->(DbSkip())
			
		EndDo
		
		nlin := nlin + 1
		
	EndDo
	
	@ nlin, 001 pSay "PERCENTUAIS SOBRE ORCAMENTO"
	
	@ nlin, 046 pSay Transform(nPer01,"@E 999.99")+" %"
	@ nlin, 062 pSay Transform(nPer02,"@E 999.99")+" %"
	@ nlin, 078 pSay Transform(nPer03,"@E 999.99")+" %"
	@ nlin, 094 pSay Transform(nPer04,"@E 999.99")+" %"
	@ nlin, 110 pSay Transform(nPer05,"@E 999.99")+" %"
	@ nlin, 126 pSay Transform(nPer06,"@E 999.99")+" %"
	@ nlin, 142 pSay Transform(nPer07,"@E 999.99")+" %"
	@ nlin, 158 pSay Transform(nPer08,"@E 999.99")+" %"
	@ nlin, 174 pSay Transform(nPer09,"@E 999.99")+" %"
	@ nlin, 190 pSay Transform(nPer10,"@E 999.99")+" %"
	@ nlin, 206 pSay Transform(nPer11,"@E 999.99")+" %"
	@ nlin, 222 pSay Transform(nPer11,"@E 999.99")+" %"
	
	nPer01:=nPer02:=nPer03:=nPer04:=nPer05:=nPer06:=nPer07:=nPer08:=nPer09:=nPer10:=nPer11:=nPer12:= 0
	nlin := nlin + 1
	
	@ nlin, 001 pSay Replicate("-",260)
	
	nlin := nlin + 2
EndDo

//EndIf

dbSelectArea("QAK2")   ;   QAK2->(dbCloseArea())
dbSelectArea("QSE5")   ;   QSE5->(dbCloseArea())
dbSelectArea("TMP")    ;   TMP->(dbCloseArea())

ChKFile("AK2")

Roda(0,"","M")
If aReturn[5] == 1
	Set Printer To
	Commit
	ourspool(wnrel)
Endif
MS_FLUSH()

Return
// ok versao 10