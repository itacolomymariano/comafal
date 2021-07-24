#Include "PROTHEUS.CH"
#Include "TopConn.Ch"
#INCLUDE "SaldoSR.CH"
#INCLUDE "SALDOSP.CH"

#DEFINE QUEBR				1
#DEFINE CLIENT				2
#DEFINE TITUL				3
#DEFINE TIPO				4
#DEFINE NATUREZA			5
#DEFINE EMISSAO			6
#DEFINE VENCTO				7
#DEFINE VENCREA			8
#DEFINE BANC				9
#DEFINE VL_ORIG			10
#DEFINE VL_NOMINAL		11
#DEFINE VL_CORRIG			12
#DEFINE VL_VENCIDO		13
#DEFINE NUMBC				14
#DEFINE VL_JUROS			15
#DEFINE ATRASO				16
#DEFINE HISTORICO			17
#DEFINE VL_SOMA			18


#DEFINE QUEBR				1
#DEFINE FORNEC				2
#DEFINE TITUL				3
#DEFINE TIPO				4
#DEFINE NATUREZA			5
#DEFINE EMISSAO			6
#DEFINE VENCTO				7
#DEFINE VENCREA			8
#DEFINE VL_ORIG			9
#DEFINE VL_NOMINAL		10
#DEFINE VL_CORRIG			11
#DEFINE VL_VENCIDO		12
#DEFINE PORTADOR			13
#DEFINE VL_JUROS			14
#DEFINE ATRASO				15
#DEFINE HISTORICO			16
#DEFINE VL_SOMA			17

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �RGERCOMAF �Autor  �Five Solutions      � Data �  05/01/2008 ���
�������������������������������������������������������������������������͹��
���Desc.     � Relat�rio de Resumo Gerencial COMAFAL                      ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � COMAFAL - PE, SP e RS.                                    ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

User Function RGerComaf

//���������������������������������������������������������������������Ŀ
//� Declaracao de Variaveis                                             �
//�����������������������������������������������������������������������

Local cDesc1         := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2         := "de acordo com os parametros informados pelo usuario."
Local cDesc3         := "Resumo Gerencial COMAFAL"
Local cPict          := ""
Local titulo       := "Resumo Gerencial COMAFAL"
Local nLin         := 80

Local Cabec1       := ""//M�s                               M�s 01      M�s 02        M�s 03"
Local Cabec2       := ""
Local imprime      := .T.
Local aOrd := {}


Private nChamada := 0 //Ita-03/03/2008
Private nPAChamada := 0

Private cNomeMes := "" //Five Solutions - 28/05/2008

Private lEnd         := .F.
Private lAbortPrint  := .F.
Private CbTxt        := ""
Private limite           := 220
Private tamanho          := "G"
Private nomeprog         := "R2GerComaf" // Coloque aqui o nome do programa para impressao no cabecalho
Private nTipo            := 15
Private aReturn          := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey        := 0
Private cPerg       := "RGERCO"
//Private cPerg      	:= "CRF010"
Private cbtxt      := Space(10)
Private cbcont     := 00
Private CONTFL     := 01
Private m_pag      := 01
Private wnrel      := "RGerComaf" // Coloque aqui o nome do arquivo usado para impressao em disco
Private dDataDe    := CTOD("")
Private dDataAte   := CTOD("")
Private d2AnoMes   := ""
Private d3AnoMes   := ""
Private cAlmoxMP   := ""
Private cAlmoxPA   := ""

Private nLRel := 68 //75 //68 //Quantidade M�xima de Linhas por P�gina


Private dDtParam  := CTOD("")
Private d2DtParam := CTOD("")
Private d3DtParam := CTOD("")

//nRecDREFin

//CFOP�s de Vendas
Private cCFOPVenda  	    := "5101,6101,5102,6102,5103,6103,5104,6104,5105,6105,"
        cCFOPVenda  	    += "5106,6106,6107,6108,5109,6109,5110,6110,5111,6111,5112,6112,"
        cCFOPVenda  	    += "5113,6113,5114,6114,5115,6115,5116,6116,5117,6117,5118,6118,5119,6119,"
        cCFOPVenda  	    += "5120,6120,5122,6122,5123,6123"//,7101,7102,7105,7106,7127"
        //Ita - Teste em Casa -  cCFOPVenda  	    += "5120,6120,5122,6122,5123,6123,7101,7102,7105,7106,7127,512"

//CFOP�s de Devolu��o de Vendas
//Ita - Teste em Casa - Private cCFOPDvVd   := "1201,2201,3201,1202,2202,3202,112,1202" //devolucao de venda
Private cCFOPDvVd   := "1201,2201,3201,1202,2202,3202" //devolucao de venda
/*
//Naturezas - (M.O. + BENEFICIAMENTO + MANUTEN��O + SERVICO TERCEIRIZADO + EPI + INSUMOS + ENERGIA + AGUA + ALUGUEL)
Private cNatProd := "100301,100302,100303,100304,100305,100306,100307,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100409,100410,"
        cNatProd += "100411,100412,100417,200314,600601,200317,200318,200319,200417,200418,200419,200320,200420,200308,200408,200312,200412,200311,200411,200310,200410,"
        cNatProd += "200309,200303,200403"
*/

//Nova Vers�o de Naturezas - Patr�cia: 13/02/2008
//Naturezas - (M.O. + BENEFICIAMENTO + MANUTEN��O + SERVICO TERCEIRIZADO + EPI + INSUMOS + ENERGIA + AGUA + ALUGUEL) / (QTD. TONELADAS PRODUZIDAS) 
Private cNatProd := "100301,100302,100303,100304,100305,100306,100307,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100409,100410,100411,"
        cNatProd += "100412,100417,200314,600601,200317,200318,200319,200417,200418,200419,200320,200420,200308,200408,200312,200412,200311,200411,200310,200410,200309,200303,200403"
/*
//Naturezas de Gastos Gerais do Comercial
Private cGGNatComl := "100120,100201,100202,100203,100204,100206,100207,100208,100209,100210,100211,100212,100213,100214,100215,"
        cGGNatComl += "100216,100222,200201,200202,200203,200204,200205,200206,200207,600101,600102,600303,700302"
*/
//Nova vers�o de Naturezas - Patr�cia 13/02/2008
//TOTAL DE GASTOS DA �REA COMERCIAL
Private cGGNatComl := "100120,100201,100202,100203,100204,100206,100207,100208,100209,100210,100211,100212,100213,100214,100215,100216,100222,"
        cGGNatComl += "200201,200202,200203,200204,200205,200206,200207,600101,600102,600303,"
        cGGNatComl += "100217" //Five Solutions Consultoria - 26/08/2008 - Solicita��o: Patr�cia M�llo/Roberto Lira
        //cGGNatComl += "100315" //Five Solutions Consultoria - 26/08/2008 - Solicita��o: Patr�cia M�llo/Roberto Lira
        //cNtGstMtOrInd += "100217"//Five Solutions Consultoria - 26/08/2008 - Solicita��o: Patr�cia M�llo/Roberto Lira

//Naturezas de Gastos Gerais M�o de Obra
//Private cNatGGMO   := "100120,100201,100202,100204,100206,100207,100208,100209,100210,100211,100213,100214,100215,100216,100222"

//Altera��o Patr�cia 23/02/2008 - Retirar 100216 de Gastos M�o-de-Obra Comercial
//Naturezas de Gastos Gerais M�o de Obra
Private cNatGGMO   := "100120,100201,100202,100204,100206,100207,100208,100209,100210,100211,100213,100214,100215,100222"



/*
//Naturezas de Despesas Fixas
Private cDFNatGG   := "100202,100203,100204,100206,100207,100208,100209,100214,100216"
*/
//Nova vers�o de Naturezas - Patr�cia 13/02/2008
//DESPESA FIXA 
Private cDFNatGG   := "100202,100204,100206,100207,100208,100209,100214,100216"
/*
//Naturezas de Despesas Peri�dicas
Private cDPNatGG   := "100120,100201,100212,100213"
*/
//Nova vers�o de Naturezas - Patr�cia 13/02/2008
//DESPESA PERI�DICA 
Private cDPNatGG   := "100120,100201,100213" 

//Naturezas de Rescis�es
Private cNResGG    := "100211,100222"
//Naturezas de Premia��es
Private cNtPremGG  := "100210,100215"
//Naturaza de Comissao
Private cNtComissao := "100203"

//Naturezas de Marketing
Private cNatMark   := "200204,200207"
//Naturezas co Despesas de Viagens
Private cNtDespV   := "200203"
//Naturezas com Fretes de Vendas
Private cNtFrtVend := "600101"
//Naturezas de Outras Despesas do Comercial
Private cNtOutDesp := "200205"

//Naturezas de Outras Naturezas
//Private nOutNatGG  := "100212,200201,200202,200206,600102,600303,700302"

//Altera��o Patr�cia - 23/02/2008 - Incluir 100216 me Outras Naturas do Comercial
//Naturezas de Outras Naturezas
Private nOutNatGG  := "100212,200201,200202,200206,600102,600303,100216"


/*
//Naturezas de GASTOS META ORCAMENTO COML
Private cNtGMtOrCml:= "100204,100206,100207,100208,100209,100214,200201,200202,200203,200204,200205,200206,200207,600303"
*/
//Nova vers�o de Naturezas - Patr�cia 13/02/2008
Private cNtGMtOrCml:= "100204,100206,100207,100208,100209,100214,200201,200202,200203,200205,200206,600303,"
        cNtGMtOrCml+= "100217"//Five Solutions Consultoria - 26/08/2008 Solicita��o: Patr�cia M�llo/Roberto Lira.
        //cNtGMtOrCml+= "100315"//Five Solutions Consultoria - 26/08/2008 Solicita��o: Patr�cia M�llo/Roberto Lira.
        //cNtGstMtOrInd += "100217"//Five Solutions Consultoria - 26/08/2008 - Solicita��o: Patr�cia M�llo/Roberto Lira
/*
//NATUREZAS PARA CUSTO DE PRODUCAO
Private cCUSTPROD := "100301,100302,100303,100304,100305,100306,100307,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100409,"
        cCUSTPROD += "100410,100411,100412,100417,200314,600601,200317,200318,200319,200417,200418,200419,200320,200420,200308,200408,200312,200412,200311,200411,200310,"
        cCUSTPROD += "200410,200309,200303,200403"
*/
//Nova vers�o de naturezas - Patr�cia 13/02/2008
//(M.O. + BENEFICIAMENTO +MANUTEN��O +SERVICO TERCEIRIZADO+EPI+INSUMOS+ENERGIA+AGUA+ALUGUEL)
Private cCUSTPROD := "100301,100302,100303,100304,100305,100306,100307,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100409,100410,100411,"
        cCUSTPROD += "100412,100417,200314,600601,200317,200318,200319,200417,200418,200419,200320,200420,200308,200408,200312,200412,200311,200411,200310,200410,200309,200303,200403"
/*
//GASTOS GERAIS IND.
Private cGASTGERIND := "100301,100302,100303,100304,100305,100306,100307,100308,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100408,100409,100410,100411,100412,"
        cGASTGERIND += "100417,200109,200110,200111,200132,200135,200301,200302,200303,200304,200305,200306,200307,200308,200309,200310,200311,200312,200313,200314,200315,200316,200317,200318,200319,200320,"
        cGASTGERIND += "200321,200322,200360,200361,200362,200363,200364,200365,200366,200367,200368,200369,200370,200372,200373,200374,200375,200376,200377,200401,200402,200403,200404,200405,200406,200407,"
        cGASTGERIND += "200408,200410,200411,200412,200416,200417,200418,200419,200420,200421,600201,600301,600501,600601,700301,700303,700304,700306,700307,700308,700309"
*/
//Nova vers�o de naturezas - Patr�cia 13/02/2008
//TOTAL DE GASTOS REALIZADOS PELA INDUSTRIA
Private cGASTGERIND := "100301,100302,100303,100304,100305,100306,100307,100308,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100408,100409,100410,100411,100412,100417,200109,"
        cGASTGERIND += "200110,200111,200132,200135,200301,200302,200303,200304,200305,200306,200307,200308,200309,200310,200311,200312,200313,200314,200315,200316,200317,200318,200319,200320,200321,200322,200360,200361,"
        cGASTGERIND += "200362,200363,200364,200365,200366,200367,200368,200369,200370,200372,200373,200374,200375,200376,200377,200401,200402,200403,200404,200405,200406,200407,200408,200410,200411,200412,200416,200417,"
        cGASTGERIND += "200418,200419,200420,200421,600201,600301,600501,600601,700301,700303,700304,700306,700307,700308,700309,"
        cGASTGERIND += "100315" //Five Solutions Consultoria - 26/08/2008 Soliti��o: Patricia M�llo/Roberto Lira 
        //cGASTGERIND += "100217" //Five Solutions Consultoria - 26/08/2008 Soliti��o: Patricia M�llo/Roberto Lira
        //cGGNatComl += "100315" //Five Solutions Consultoria - 26/08/2008 - Solicita��o: Patr�cia M�llo/Roberto Lira

//Natures de Gastos Totais com Ativo Imobilizado
Private cImobTotal := "200109,200110,200111,200306,200406,700309"

//TOTAL DE GASTOS REALIZADO COM M�QUINAS E MOTORES
Private cNatMaqMot := "200111,200306,200406"
//TOTAL DE GASTOS REALIZADO COM IMOBILIZADO IMPORTADO 
Private cNatImpInd := "700309"
//TOTAL DE GASTOS REALIZADO COM IMOBILIZADO DA �REA DE INFORM�TICA 
Private cNatInfInd := "200109"
//TOTAL DE GASTOS REALIZADO COM M�VEIS E UTENS�LIOS 
Private cMovUteInd := "200110"
//TOTAL DE INVESTIMENTOS REALIZADOS NA EMPRESA (SOMAT�RIO DAS NATUREZAS QUE POSSUEM GASTOS COM INVESTIMENTOS)
Private cInvTotInd := "200132,200135,200313,200322,200360,200361,200362,200363,200364,200365,200366,200367,200368,200369,200370,200372,200373,200374,200375,200376,200377"
//INVESTIMENTOS NAS EMPRESAS CABOMAR E COMAFAL-RS
Private cNtEmpresas:= "200135,200360"
//INVESTIMENTOS LAN�ADOS NAS NATUREZAS DE M�QUINAS E EQUIPAMENTOS 
Private cNtMaquinas:= "200361,200362,200363,200364,200369,200370,200375"
//VALOR DE FINAME LAN�ADO NA NATUREZA 
Private cNtFiname := "200313"
//VALORES LAN�ADOS REF. LEASING 
Private cNtLeasing:= "200132"
//SOMAT�RIO DAS NATUREZAS QUE POSSUEM DESPESAS REF. A BENFEITORIAS 
Private cNtBenfeit:= "200322, 200365,200366,200368,200372,200373,200374,200376,200377"
//VALORES LAN�ADOS NA NATUREZA ISO 9001 - 2000  
Private cNtISO92  := "200367"
//SOMAT�RIO DAS NATUREZAS QUE POSSUEM GASTOS RELACIONADOS A MATERIA PRIMA 
Private cNtMPTotal:= "700303,700306,700307,700308,600201"
//MATERIA PRIMA NACIONAL 
Private cNtNacMP := "700303"
//MAT�RIA PRIMA IMPORTADA   
Private cNtImpMP := "700306"
//TRIBUTOS DA IMPORTA��O  
Private cNtTrbImp:= "700307"
//DESPESAS DE IMPORTA��O 
Private cNtDespImp:="700308"
//FRETE IMPORTA��O 
Private cNtFrtImp:= "600201"
//INSUMOS 
Private cNtInsumos := "200308,200408,200312,200412"
/*
//DESPESAS COM MAO DE OBRA
Private cNtMaoObInd:="100301,100302,100303,100304,100305,100306,100307,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100409,100410,100411,100412,100417"
*/
//Nova vers�o Naturezas - Patr�cia 13/02/2008
//Private cNtMaoObInd:="100301,100302,100303,100304,100305,100306,100307,100308,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100408,100409,100410,100411,100412,100417"
//Nova altera��o - Patr�cia - 20/02/2008 - Movido naturezas 100308,100408 de M�o-de-Obra e Despesas peri�dicas p/ Outras Naturezas
Private cNtMaoObInd:="100301,100302,100303,100304,100305,100306,100307,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100409,100410,100411,100412,100417"
//DESPESA FIXA Ind�stria 
Private cNtMODFInd:= "100303,100403,100305,100405,100306,100406,100307,100407,100311,100411,100312,100412"                 
//DESPESA PERI�DICA Ind�stria
//Nova altera��o - Patr�cia - 20/02/2008 - Movido naturezas 100308,100408 de M�o-de-Obra e Despesas peri�dicas p/ Outras Naturezas
//Private cNtMODPInd:= "100301,100401,100304,100404,100308,100408,100309,100409"
Private cNtMODPInd:= "100301,100401,100304,100404,100309,100409"
//RESCIS�O Ind�stria 
Private cNtRecisInd:= "100310,100410,100317,100417"
//PREMIA��ES Ind�stria 
Private cNtPremInd := "100302,100402,100313"
//BENEFICIAMENTO  Ind�stria 
Private cNtBenefInd:= "200314,600601"
//MANUTEN��O Ind�stria 
Private cNtManuten:= "200317,200318,200319,200417,200418,200419"
//SERVI�OS TERCEIRIZADOS Ind�stria 
Private cNtSerTerInd:= "200303,200403"
//EPI Ind�stria 
Private cNtEPIInd := "200320,200420"
//ENERGIA Ind�stria 
Private cNtEnergInd:= "200311,200411"
//AGUA  Ind�stria 
Private cNtAguaInd := "200310,200410"
//ALUGUEL Ind�stria 
Private cNtAlugInd := "200309"
//CONSERVA��O DE PREDIO Ind�stria 
Private cNtCoPredInd:= "200321,200421"
//DESPESAS DE VIAGENS  Ind�stria 
Private cNtDespViaIn:= "200304,200404"
//OUTRAS DESPESAS DA IND�STRIA  Ind�stria 
//Private cNtOutDesInd := "200316,200416"


//OUTRAS DESPESAS DAS IND�STRIA
//'200321','200421','200304','200404','200316','200416','200301','200302','200305','200307','200315','200401','200402','200405','200407','600301','600501','700301','700304'
//Private cNtOutDesInd := "200321,200421,200304,200404,200316,200416,200301,200302,200305,200307,200315,200401,200402,200405,200407,600301,600501,700301,700304"
//Corre��o no ato da reuni�o
Private cNtOutDesInd := "200316"

/*
//OUTRAS NATUREZAS DA IND�STRIA 
Private cNtOutNatInd := "100308,100408,200301,200302,200305,200307,200315,200401,200402,200405,200407,600301,600501,700301,700304"
*/
//Nova vers�o Naturezas - Patr�cia: 13/02/2008
//Private cNtOutNatInd := "200301,200302,200305,200307,200315,200401,200402,200405,200407,600301,600501,700301,700304"
//Nova altera��o - Patr�cia - 20/02/2008 - Movido naturezas 100308,100408 de M�o-de-Obra e Despesas peri�dicas p/ Outras Naturezas
//Private cNtOutNatInd := "200301,200302,200305,200307,200315,200401,200402,200405,200407,600301,600501,700301,700304,100308,100408"
Private cNtOutNatInd := "100308,100408,200301,200302,200305,200307,200315,200401,200402,200405,200407,600301,600501,700301,700304"

/*
//TOTAL DE GASTOS REALIZADOS NAS NATUREZAS QUE COMP�E AS METAS DE OR�AMENTO
Private cNtGstMtOrInd := "100303,100304,100305,100306,100307,100311,100312,100403,100404,100405,100406,100407,100411,100412,200301,200302,200303,200304,200305,200308,200309,200310,200311,"
        cNtGstMtOrInd += "200312,200316,200317,200318,200319,200320,200321,200322,200401,200402,200403,200404,200405,200408,200410,200411,200412,200416,200417,200418,200419,200420,200421,"
        cNtGstMtOrInd += "600301"
*/
//Nova vers�o Naturezas - Patr�cia 13/02/2008
//TOTAL DE GASTOS REALIZADOS NAS NATUREZAS QUE COMP�E AS METAS DE OR�AMENTO
Private cNtGstMtOrInd := "100303,100305,100306,100307,100311,100312,100403,100405,100406,100407,100411,100412,200302,200304,200305,200308,200309,200310,200311,200312,200316,200317,200318,200319,200320,200321,"
        cNtGstMtOrInd += "200402,200404,200405,200408,200410,200411,200412,200416,200417,200418,200419,200420,200421,600301,"
        cNtGstMtOrInd += "100315"//Five Solutions Consultoria - 26/08/2008 - Solicita��o: Patr�cia M�llo/Roberto Lira 
        //cNtGstMtOrInd += "100217"//Five Solutions Consultoria - 26/08/2008 - Solicita��o: Patr�cia M�llo/Roberto Lira
        //cGGNatComl += "100315" //Five Solutions Consultoria - 26/08/2008 - Solicita��o: Patr�cia M�llo/Roberto Lira
/*        
//TOTAL DE GASTOS REALIZADOS PELO ADMINISTRATIVO
Private cGstGerADM  := "100101,100102,100103,100104,100105,100106,100107,100108,100109,100110,100111,100112,100113,100114,100115,100116,100117,100118,100119,100121,100122,"
        cGstGerADM  += "200101,200102,200103,200104,200105,200106,200107,200108,200112,200113,200114,200115,200116,200117,200118,200119,200120,200121,200122,200123,200124,"
        cGstGerADM  += "200126,200127,200128,200129,200130,200131,200133,200134,200136,200137,200138,200139,200140,200141,200142,200143,200208,200211,200501,200502,200503,"
        cGstGerADM  += "200504,200505,200506,200507,200508,200509,200510,500101,500102,500103,500104,600302,700206,700310,D.ESCONT,I.RF,I.SS,S.ANGRIA,T.ROCO"
*/
//Nova vers�o Naturezas - Patr�cia 13/02/2008
//TOTAL DE GASTOS REALIZADOS PELO ADMINISTRATIVO
//GASTOS GERAIS DO ADM
/*  Comentado para pegar nova vers�o de naturezas disponibilizada por Patr�cia em 28/05/2008
Private cGstGerADM  := "100101,100102,100103,100104,100105,100106,100107,100108,100109,100110,100111,100112,100113,100114,100115,100116,100117,100118,100119,100121,100122,200101,200102,"
        cGstGerADM  += "200103,200104,200105,200106,200107,200108,200112,200113,200114,200115,200116,200117,200118,200119,200120,200121,200122,200123,200124,200126,200127,200128,200129,"
        cGstGerADM  += "200130,200131,200133,200134,200136,200137,200138,200139,200140,200141,200142,200143,200208,200211,200501,200502,200503,200504,200505,200506,200507,200508,200509,"
        cGstGerADM  += "200510,500101,500102,500103,500104,600302,700206,700310,DESCONT,INSS,IRF,ISS,SANGRIA,TROCO"
        */
//Novas Naturezas do GASTOS GERAIS DO ADM em 28/05/2008 - Disponibilizada por Patricia.
Private cGstGerADM  := "100101,100102,100103,100104,100105,100106,100107,100108,100109,100110,100111,100112,100113,100114,100115,100116,100117,100118,100119,100121,"
        cGstGerADM  += "100122,200101,200102,200103,200104,200105,200106,200107,200108,200112,200113,200114,200115,200116,200117,200118,200119,200120,200121,200122,"
        cGstGerADM  += "200123,200124,200126,200127,200128,200129,200130,200131,200133,200134,200136,200137,200138,200139,200140,200141,200142,200143,200208,200211,"
        cGstGerADM  += "200501,200502,200503,200504,200505,200506,200507,200508,200509,200510,500101,500102,500103,500104,600302,700206,700310,COFINS,CSLL,INSS,IRF,ISS,PIS"
        
/*        
//MAO DE OBRA ADM 
Private cMOGstADM := "100101,100102,100103,100104,100105,100106,100107,100109,100110,100111,100114,100115,100116,100117,100118,100119,100122"
*/
//Nova vers�o dNaturezas 13/02/2008
//MAO DE OBRA ADM 
//Private cMOGstADM := "100101,100102,100103,100104,100105,100106,100107,100108,100109,100110,100111,100113,100114,100116,100117,100118,100119,100121,100122" 
//Nova vers�o dNaturezas 23/02/2008
//MAO DE OBRA ADM 
Private cMOGstADM := "100101,100102,100103,100104,100105,100106,100107,100109,100110,100111,100114,100116,100117,100118,100119,100121,100122" 
/*
//DESPESA FIXA ADM
Private cDFGstADM := "100103,100105,100106,100107,100117,100116,,100111,100114,100108"
*/
//Nova vers�o Naturezas - Patr�cia 13/02/2008
//Private cDFGstADM := "100103,100105,100106,100107,100108,100111,100114,100116,100117"
//DESPESA FIXA "100103,100105,100106,100107,100111,100114,100116,100117" - Vers�o 18/03/2008
Private cDFGstADM := "100103,100105,100106,100107,100111,100114,100116,100117"

//DESPESA PERI�DICA ADM 
//Private cDPGstADM := "100101,100104,100109,100113,100119,100121"
// DESPESA PERI�DICA "100101,100104,100109,100119,100121" - Vers�o 18/03/2008
Private cDPGstADM := "100101,100104,100109,100119,100121"

//RESCIS�O ADM
Private cResGstADM:= "100110,100118,100122"

//PREMIA��ES ADM
Private cPrmADMGst:= "100102"

//DESPESAS DE VIAGENS ADM
Private cDspViaADM:= "200105"

//OUTRAS DESPESAS ADM. 
Private cOutDADMGst:= "200123"

/*
//SERVI�OS TERCEIRIZADOS ADM
Private cSrvTADMGst := "200103"
*/
//Nova vers�o Naturezas Patr�cia 13/02/2008
//SERVI�OS TERCEIRIZADOS  
Private cSrvTADMGst := "200103,200122,200124,200133"


//Novos Indices
//HONOR�RIOS ADVOCAT�CIOS 
Private cHonorAdvNt := "100112"
//DESPESAS JUDICIAIS 
Private cDespJdcNt  := "200106"
//RESCIS�ES ACORDOS / ADV.  
Private cRescAcordNt:= "200130"

/*
//OUTRAS NATUREZAS DO ADMINISTRATIVO FINANCEIRO ADM
Private OutrasADM := "100108,100112,100113,100121,200101,200102,200103,200104,200106,200107,200108,200112,200113,200114,200115,200116,200117,200118,200119,200120,"
        OutrasADM += "200121,200122,200124,200126,200127,200128,200129,200130,200131,200133,200134,200136,200137,200138,200139,200140,200141,200142,200143,200208,"
        OutrasADM += "200211,200501,200502,200503,200504,200505,200506,200507,200508,200509,200510,500101,500102,500103,500104,600302,700206,700310,900601,"
        OutrasADM += "D.ESCONT,I.RF,I.SS,S.ANGRIA,T.ROCO"
*/
//Nova vers�o Naturezas - Patr�cia 13/02/2008
//OUTRAS NATUREZAS DO ADMINISTRATIVO FINANCEIRO
//Private OutrasADM := "100115,200101,200102,200104,200107,200108,200112,200113,200114,200115,200116,200117,200118,200119,200120,200121,200128,200129,200131,200134,200136,"
//        OutrasADM += "200137,200140,200141,500101,500102,500103,500104,600302,DESCONT,INSS,SANGRIA,TROCO,100108,100113"

//OUTRAS NATUREZAS DO ADMINISTRATIVO FINANCEIRO - Nova vers�o 18/03/2008
Private OutrasADM := "100108,100113,100115,200101,200102,200104,200107,200108,200112,200113,200114,200115,200116,200117,200118,200119,200120,200121,200128,200129,200131,"
        OutrasADM += "200134,200136,200137,200140,200141,500101,500102,500103,500104,600302,DESCONT,INSS,SANGRIA,TROCO"
        
/*
//TOTAL DE GASTOS REALIZADOS NAS NATUREZAS QUE COMP�E AS METAS DE OR�AMENTO ADM
Private AdmMetasOrc := "100103,100104,100105,100106,100107,100111,100112,100114,100115,100116,100117,200101,200102,200103,200104,200105,200106,200107,200108,200112,200113,200114,200115,200116,200117,200118,200121,200122,200123,200124,200129,200131,200133,200134,200136,200143,200501,200502,200503,200504,200505,200506,200507,200508,200509,200510,500102,500103,500104,600302"
*/
//Nova vers�o de Naturezas - Patr�cia 13/02/2008
//TOTAL DE GASTOS REALIZADOS NAS NATUREZAS QUE COMP�E AS METAS DE OR�AMENTO
//Private AdmMetasOrc := "100103,100105,100106,100107,100108,100111,100112,100114,100115,100116,100117,200102,200104,200105,200107,200112,200113,200114,200116,200117,200118,200121,200123,200124,200129,200131,"
//        AdmMetasOrc += "200133,200134,200136,200143,200501,200502,200503,200504,200505,200506,200507,200508,200509,200510,500102,500103,500104,600302"
//TOTAL DE GASTOS REALIZADOS NAS NATUREZAS QUE COMP�E AS METAS DE OR�AMENTO  - Vers�o 18/03/2008
//Five Solutions - 28/08/2008 - Conforme Solicita��o de Patr�cial M�lo e Autoriza��o da Diretoria COMAFAL
// Foram retiradas as naturezas 500102, 500103 e 500104�, respectivamente �Despesas de Viagem da Diretoria, INSS da Diretoria e Plano de Sa�de da Diretoria� das metas de or�amento do Administrativo
/*
Private AdmMetasOrc := "100103,100105,100106,100107,100111,100112,100114,100115,100116,100117,200102,200104,200105,200107,200112,200113,200114,200116,200117,200118,200121,200123,200124,200129,200131,"
        AdmMetasOrc += "200133,200134,200136,200143,200501,200502,200503,200504,200505,200506,200507,200508,200509,200510,500102,500103,500104,600302"
        */
Private AdmMetasOrc := "100103,100105,100106,100107,100111,100112,100114,100115,100116,100117,200102,200104,200105,200107,200112,200113,200114,200116,200117,200118,200121,200123,200124,200129,200131,"
        AdmMetasOrc += "200133,200134,200136,200143,200501,200502,200503,200504,200505,200506,200507,200508,200509,200510,600302"

//EMPR�STIMO ENTRE EMPRESAS 
Private cNtEmpEEmp := "700206"

//FUNDO FIXO "200143,200501,200502,200503,200504,200505,200506,200507,200508,200509,200510"
Private cNtFundFix := "200143,200501,200502,200503,200504,200505,200506,200507,200508,200509,200510"        

/*
//TOTAL DE IMPOSTOS
Private cNtImpCtrbTt := "100115,200137,500103,100114,100122,100214,100222,100312,100317,100412,100417,200126,200127,200208,200211,"
        cNtImpCtrbTt += "200138,I.SS,200139,700307,I.RF" 
*/
//Nova vers�o Naturezas - Patr�cia 13/02/2008
//IMPOSTOS
 Private cNtImpCtrbTt := "200139,IRF,200211,200208,200142,700310,200138,ISS,200127,200126"

//INSS 
Private cNtINSS := "100115,200137,500103"

//FGTS - GRFC 
Private cNtFGTS := "100114,100122,100214,100222,100312,100317,100412,100417"

//CPMF
Private cNtCPMF := "200126" 

//IPTU  
Private cNtIPTU := "200127"

//IPI 
Private cNtIPI := "200208"

//PIS / COFINS  
Private cNtPISCOF := "200142"

//TRIBUTOS DIVERSOS  
Private cNtTribut := "700310"

//ICMS 
Private cNtICMS := "200211"

/*
//ISS  
Private cNtISS := "200138,I.SS"
*/
//Nova vers�o Naturezas - Patr�cia 13/02/2008
//ISS  
Private cNtISS := "200138,ISS"

/*
//IR 
Private cNtIR := "200139,I.RF"
*/
//Nova vers�o Naturezas - Patr�cia 13/02/2008
Private cNtIR := "200139,IRF"

//TODOS OS IMPOSTOS DA IMPORTA��O  
Private cNtImpImport := "700307"


//Equalizando Relat�rio
Private cCFOPVdInt  	    := "'5101','6101','5102','6102','5103','6103','5104','6104','5105','6105',"
cCFOPVdInt  	    += "'5106','6106','6107','6108','5109','6109','5110','6110','5111','6111','5112','6112',"
cCFOPVdInt  	    += "'5113','6113','5114','6114','5115','6115','5116','6116','5117','6117','5118','6118','5119','6119',"
cCFOPVdInt  	    += "'5120','6120','5122','6122','5123','6123'"

//Variaveis Privadas de Impress�o
//Saldos em Aberto a Pagar
nSld1MPag := 0
nSld2MPag := 0
nSld3MPag := 0
nSld1MAnt := 0
nSld2MAnt := 0
nSld3MAnt := 0

//Saldos em Aberto a Receber
nSld1MRec := 0
nSld2MRec := 0
nSld3MRec := 0


// Variaveis do Fluxo do Contas a Receber
 nARe1M1a30  := 0
 nARe2M1a30  := 0
 nARe3M1a30  := 0
 nARe1M31a60 := 0
 nARe2M31a60 := 0
 nARe3M31a60 := 0

//Five Solutions Consultoria
//20 de agosto de 2008  
nARAc160dd := 0
nARAc260dd := 0
nARAc360dd := 0

 nARe1M61a90 := 0
 nARe2M61a90 := 0
 nARe3M61a90 := 0
 nARe1M90M   := 0
 nARe2M90M   := 0
 nARe3M90M   := 0

//Variaveis de Clientes
nClAtv1Mes := 0
nClIna1Mes := 0
nClAtv2Mes := 0
nClIna2Mes := 0
nClAtv3Mes := 0
nClIna3Mes := 0
nCliN01Mes := 0
nCliN02Mes := 0
nCliN03Mes := 0
nClAtd1Mes := 0
nClAtd2Mes := 0
nClAtd3Mes := 0

//Prospects
nASC1Mes := 0
nASC2Mes := 0
nASC3Mes := 0

// Variaveis de Pedidos

nFatPd1QtMes := 0
nPdPen1QtMes := 0
nFatPd2QtMes := 0
nPdPen2QtMes := 0
nFatPd3QtMes := 0
nPdPen3QtMes := 0
nFtPd1VlrMes := 0
nPdPen1VlrM  := 0
nFtPd2VlrMes := 0
nPdPen2VlrM  := 0
nFtPd3VlrMes := 0
nPdPen3VlrM  := 0

nFtMAPd1Qt   := 0
nFtMAPd2Qt   := 0
nFtMAPd3Qt   := 0
nFtPd1MAVl   := 0
nFtPd2MAVl   := 0
nFtPd3MAVl   := 0

nPdPen1MAQt  := 0
nPdPen2MAQt  := 0
nPdPen3MAQt  := 0

nPdPMA1VlM   := 0
nPdPMA2VlM   := 0
nPdPMA3VlM   := 0

// Variaveis de Faturamento

nFat1MesBru := 0
nFat2MesBru := 0
nFat3MesBru := 0

// Faturamento CIF / FOB
nFOB1MFt := 0
nFOB2MFt := 0
nFOB3MFt := 0

//FRETES (Total)

//FRETE VENDAS '600101'
cNtVendFrt := "600101"
//FRETE IMPORTA��O '600201'
cNtImporFrt := "600201"
//FRETE EXPORTA��O '600401'
cNtExporFrt := "600401"
//FRETE BENEFICIAMENTO '600601'
cNtBenefFrt := "600601"
//FRETE TRANSFERENCIA '600501'
cNtTransfFrt := "600501"
//FRETE COM�RCIO  '600303'
cNtComercFrt := "600303"
//FRETE IND�STRIA '600301'
cNtIndusFrt := "600301"
//FRETE ADMINISTRATIVO '600302'
cNtAdminFrt := "600302"

//Valores Pagos Ref. Fretes
n1FrtVeMVl := 0
n2FrtVeMVl := 0
n3FrtVeMVl := 0
n1FrtImpMVl := 0
n2FrtImpMVl := 0
n3FrtImpMVl := 0
n1FrtExpMVl := 0
n2FrtExpMVl := 0
n3FrtExpMVl := 0
n1FrtBenMVl := 0
n2FrtBenMVl := 0
n3FrtBenMVl := 0
n1FrtTrfMVl := 0
n2FrtTrfMVl := 0
n3FrtTrfMVl := 0
n1FrtComMVl := 0
n2FrtComMVl := 0
n3FrtComMVl := 0
n1FrtIndMVl := 0
n2FrtIndMVl := 0
n3FrtIndMVl := 0
n1FrtAdmMVl := 0
n2FrtAdmMVl := 0
n3FrtAdmMVl := 0

nFt1MCIF := 0
nFt2MCIF := 0
nFt3MCIF := 0

nFt1MSDEF:= 0
nFt2MSDEF:= 0
nFt3MSDEF:= 0

aCFTList := {}

nVl1MesDev := 0
nVl2MesDev := 0
nVl3MesDev := 0

nCus1MesDev := 0
nCus2MesDev := 0
nCus3MesDev := 0

n1MesFtLiq  := 0
n2MesFtLiq  := 0
n3MesFtLiq  := 0

nCMP1Mes    := 0
nCMP2Mes    := 0
nCMP3Mes    := 0

nMg1Mes := 0 
nMg2Mes := 0
nMg3Mes := 0

nProdTon1Mes := 0 
nProdTon2Mes := 0
nProdTon3Mes := 0

nVlr1MesBx := 0
nProd1Mes  := 0
nVlr2MesBx := 0
nProd2Mes  := 0
nVlr3MesBx := 0
nProd3Mes  := 0
      
n1MTonelVen  := 0
n2MTonelVen  := 0
n3MTonelVen  := 0

nQtTon1MVen := 0
nQtTon1MAPd := 0
nQtTon2MVen := 0
nQtTon2MAPd := 0
nQtTon3MVen := 0
nQtTon3MAPd := 0
            
nMatPrim1Mes :=0 
nMatPrim2Mes :=0 
nMatPrim3Mes :=0 

nCus1MesFat := 0
n1MTonelVen := 0
nCus2MesFat := 0
n2MTonelVen := 0
nCus3MesFat := 0
n3MTonelVen := 0
      
nOutras := 0

nCMTon1Mes := 0 
nCMTon2Mes := 0 
nCMTon3Mes := 0 

nProdTon1Mes := 0
nMatPrim1Mes := 0
nProdTon2Mes := 0
nMatPrim2Mes := 0 
nProdTon3Mes := 0
nMatPrim3Mes := 0

nPMVen1Mes  := 0
nPMVen2Mes  := 0
nPMVen3Mes  := 0

nMgMed1Mes := 0
nMgMed2Mes := 0
nMgMed3Mes := 0

nPMVen1Mes := 0
nCMTon1Mes := 0
nPMVen2Mes := 0
nCMTon2Mes := 0
nPMVen3Mes := 0
nCMTon3Mes := 0

//Comiss�o

nVl1MComis := 0
nVl2MComis := 0
nVl3MComis := 0

//Gastos Gerais

nVl1MGGComl := 0
nVl2MGGComl := 0
nVl3MGGComl := 0

nVl1MGGMO  := 0
nVl2MGGMO  := 0
nVl3MGGMO  := 0

nVl1MDFGG  := 0
nVl2MDFGG  := 0
nVl3MDFGG  := 0
      
nVl1MDPGG  := 0
nVl2MDPGG  := 0
nVl3MDPGG  := 0

nVl1MResGG  := 0
nVl2MResGG  := 0
nVl3MResGG  := 0
      
nVl1MPrmGG  := 0
nVl2MPrmGG  := 0
nVl3MPrmGG  := 0

nVl1MComis  := 0
nVl2MComis  := 0
nVl3MComis  := 0

nVl1MMktGG  := 0
nVl2MMktGG  := 0
nVl3MMktGG  := 0
      
nVl1MDVgGG  := 0
nVl2MDVgGG  := 0 
nVl3MDVgGG  := 0

nVl1MDFrtGG  := 0
nVl2MDFrtGG  := 0
nVl3MDFrtGG  := 0
      
nVl1MOutDGG  := 0
nVl2MOutDGG  := 0
nVl3MOutDGG  := 0
      
nVl1MONtGG  := 0
nVl2MONtGG  := 0
nVl3MONtGG  := 0

//Metas do Or�amento
/*
nMetasVlr  := 0
nMetasVlr  := 0
nMetasVlr  := 0
*/
      
nGst1MMtOrC  := 0
nGst2MMtOrC  := 0
nGst3MMtOrC  := 0

//Faturamento Sucata

nVl1SUCMFat  := 0
nVl2SUCMFat  := 0
nVl3SUCMFat  := 0
      
aTSUCList := {}
   
nVlM1CusSUC  := 0
nVlM2CusSUC := 0
nVlM3CusSUC := 0
      
nMg1SUCMes := 0
nMg2SUCMes := 0
nMg3SUCMes := 0

nQt1MSUCVen := 0
nQt2MSUCVen  := 0
nQt3MSUCVen  := 0
      
nPMV1SUCM  := 0
nPMV2SUCM  := 0
nPMV3SUCM  := 0

// Variaveis Ind�stria


nProvd1M := 0
nProvd2M := 0
nProvd3M := 0


n1MTonelVen := 0
n2MTonelVen  := 0
n3MTonelVen  := 0

nProd1Mes := 0
nProd2Mes := 0
nProd3Mes := 0

aMaqPrdM := {}

nQtdFM1MEst := 0
nQtdFM2MEst := 0
nQtdFM3MEst := 0
      
nVlrFM1MEst  := 0
nVlrFM2MEst := 0
nVlrFM3MEst  := 0

//Materia Prima Sliter

nQt1MPSlit := 0
nQt2MPSlit := 0
nQt3MPSlit := 0

nVl1MCSlit := 0
nVl2MCSlit := 0
nVl3MCSlit := 0
      
nVM1MMP := 0
nVM2MMP := 0
nVM3MMP := 0

nQt1PAM := 0
nQt2PAM := 0
nQt3PAM := 0

nCs1PAM  := 0
nCs2PAM  := 0
nCs3PAM  := 0

nVM1MPA := 0
nVM2MPA := 0
nVM3MPA := 0
      
nVlr1CusPro  := 0
nVlr2CusPro  := 0
nVlr3CusPro  := 0

// Gastos Gerais Ind�stria

nVGG1MInd  := 0
nVGG2MInd  := 0
nVGG3MInd  := 0

nVImob1MTt := 0
nVImob2MTt := 0
nVImob3MTt := 0
      
nVMqMt1MTt  := 0
nVMqMt2MTt  := 0
nVMqMt3MTt  := 0

nVlM1ImpInd  := 0
nVlM2ImpInd  := 0
nVlM3ImpInd  := 0

nVlInfM1Ind  := 0
nVlInfM2Ind  := 0
nVlInfM3Ind  := 0
      
nVlMU1MInd  := 0
nVlMU2MInd  := 0
nVlMU3MInd  := 0

nVl1MInvTt  := 0
nVl2MInvTt  := 0
nVl3MInvTt  := 0

nVl1MEmpInv  := 0
nVl2MEmpInv := 0
nVl3MEmpInv := 0

nVlMq1MInv  := 0
nVlMq2MInv  := 0
nVlMq3MInv  := 0

nVlFin1MInv  := 0
nVlFin2MInv  := 0
nVlFin3MInv  := 0
      
nVlLea1MInv  := 0
nVlLea2MInv  := 0
nVlLea3MInv  := 0


nVlBenf1MI  := 0
nVlBenf2MI  := 0
nVlBenf3MI  := 0

nVlISO1MI  := 0
nVlISO2MI  := 0
nVlISO3MI  := 0

nMP1MVlTt  := 0
nMP2MVlTt := 0
nMP3MVlTt := 0

nNacMP1MVl  := 0
nNacMP2MVl := 0
nNacMP3MVl := 0

nImpMP1MVl  := 0
nImpMP2MVl := 0
nImpMP3MVl  := 0

nTrbImp1MMP  := 0
nTrbImp2MMP := 0
nTrbImp3MMP  := 0


nDspImp1MMP  := 0
nDspImp2MMP  := 0
nDspImp3MMP  := 0

nFrtImp1MMP  := 0
nFrtImp2MMP := 0
nFrtImp3MMP := 0

nInsVl1MInd  := 0
nInsVl2MInd := 0
nInsVl3MInd := 0

nMOInd1MVl  := 0
nMOInd2MVl := 0
nMOInd3MVl := 0

nFixMO1MInd  := 0
nFixMO2MInd := 0
nFixMO3MInd := 0

nPerMO1MInd  := 0
nPerMO2MInd := 0
nPerMO3MInd  := 0
      
nRes1MIndVl  := 0
nRes2MIndVl := 0
nRes3MIndVl := 0
      
nPrm1IndVl := 0
nPrm2IndVl := 0
nPrm3IndVl := 0

nBnf1IndVl  := 0
nBnf2IndVl := 0
nBnf3IndVl := 0

nMan1IndVl  := 0
nMan2IndVl  := 0
nMan3IndVl  := 0
      
nSTc1IndVl  := 0
nSTc2IndVl := 0
nSTc3IndVl := 0

nEPI1IndVl  := 0
nEPI2IndVl := 0
nEPI3IndVl := 0

nEner1IndVl  := 0
nEner2IndVl := 0
nEner3IndVl := 0

nAgua1IndVl := 0
nAgua2IndVl := 0
nAgua3IndVl := 0

nAlg1IndVl  := 0
nAlg2IndVl  := 0
nAlg3IndVl  := 0

nCPr1IndVl  := 0
nCPr2IndVl  := 0
nCPr3IndVl  := 0

nDVg1IndVl  := 0
nDVg2IndVl  := 0
nDVg3IndVl  := 0
      
nODp1IndVl  := 0
nODp2IndVl  := 0
nODp3IndVl  := 0

nONt1IndVl  := 0
nONt2IndVl  := 0
nONt3IndVl  := 0

nGst1MtOrInd  := 0
nGst2MtOrInd := 0
nGst3MtOrInd := 0

//Variaveis Administrativo - Financeiro

nPd1MTrvTt := 0
nPd2MTrvTt := 0
nPd3MTrvTt := 0

nPdTrv1MVlr:=0
nPdTrv1AntM:=0
nPdTrv2MVlr:=0
nPdTrv2AntM:=0
nPdTrv3MVlr:=0
nPdTrv3AntM:=0

nTit1MTrvTt := 0
nTit2MTrvTt := 0
nTit3MTrvTt := 0

nTit1MTrvVlr:=0
nAntTit1MTrv:=0                  
nTit2MTrvVlr:=0
nAntTit2MTrv:=0
nTit3MTrvVlr:=0
nAntTit3MTrv:=0

// T�tulos

nAPg1MTit := 0
nAPg2MTit := 0
nAPg3MTit  := 0

nPg1MRealiz := 0
nPg2MRealiz := 0
nPg3MRealiz  := 0
      
nBx1MComp := 0
nBx2MComp := 0
nBx3MComp  := 0

nDBT1MCC := 0
nDBT2MCC := 0
nDBT3MCC := 0

nBx1MNorm := 0
nBx2MNorm := 0
nBx3MNorm := 0

nBx1MDevol := 0
nBx2MDevol := 0
nBx3MDevol  := 0

nPgAb1MVlr := 0
nPgAb2MVlr := 0
nPgAb3MVlr := 0

nSld1MPag := 0
nSld2MPag := 0
nSld3MPag := 0

nSld1MAnt := 0
nSld2MAnt := 0
nSld3MAnt := 0

nTt1MAPG := 0
nTt2MAPG := 0
nTt3MAPG := 0

nAPg1M1a30:=0
nAPg1M31a60:=0

//Five Solutions Consultoria
//20 de agosto de 2008
//Limite de demonstra��o de prazo de t�tulos no fluxo de caixa, para "acima de 60 dias"
nAcm1M60dd := 0
nAcm2M60dd := 0
nAcm3M60dd := 0

nAPg1M61a90:=0
nAPg1M90M:=0
nAPg2M1a30:=0
nAPg2M31a60:=0
nAPg2M61a90:=0
nAPg2M90M:=0
nAPg3M1a30:=0
nAPg3M31a60:=0
nAPg3M61a90:=0
nAPg3M90M:=0

nARe1MTit := 0
nARe2MTit := 0
nARe3MTit := 0

nRe1MRealiz := 0
nRe2MRealiz := 0
nRe3MRealiz := 0
      
nBxRE1MComp := 0
nBxRE2MComp := 0
nBxRE3MComp := 0
      
nBx1MDvRE := 0
nBx2MDvRE := 0
nBx3MDvRE := 0

nBxRE1MNorm := 0
nBxRE2MNorm := 0
nBxRE3MNorm := 0

nSld1MRec := 0
nSld2MRec := 0
nSld3MRec := 0

nPer1MInadpl := 0
nPer2MInadpl := 0
nPer3MInadpl := 0

nTt1MVc := 0
nTt2MVc := 0
nTt3MVc := 0

nAAtTt1MVc := 0
nAAtTt2MVc := 0
nAAtTt3MVc := 0

nA1ATt1MVc := 0
nA1ATt2MVc := 0
nA1ATt3MVc := 0

nA2ATt1MVc := 0
nA2ATt2MVc := 0
nA2ATt3MVc := 0

nAAntTt1MVc := 0
nAAntTt2MVc := 0
nAAntTt3MVc := 0


nTtFlx1MARE := 0
nTtFlx2MARE := 0
nTtFlx3MARE := 0

nARe1M1a30:=0
nARe1M31a60:=0
nARe1M61a90:=0
nARe1M90M:=0
nARe2M1a30:=0
nARe2M31a60:=0
nARe2M61a90:=0
nARe2M90M:=0
nARe3M1a30:=0
nARe3M31a60:=0
nARe3M61a90:=0
nARe3M90M:=0

// Gastos Gerais Adm/Fin
nVl1MAdmGG := 0
nVl2MAdmGG := 0
nVl3MAdmGG := 0
nVGG1MMOADM := 0
nVGG2MMOADM := 0
nVGG3MMOADM := 0
nVGG1MDFADM := 0
nVGG2MDFADM := 0
nVGG3MDFADM := 0
nVGG1MDPAD := 0
nVGG2MDPAD := 0
nVGG3MDPAD := 0
nVRes1MAdm := 0
nVRes2MAdm := 0
nVRes3MAdm := 0
nVPrm1MAdm := 0
nVPrm2MAdm := 0
nVPrm3MAdm := 0
nVDVia1MAdm := 0
nVDVia2MAdm := 0
nVDVia3MAdm := 0
nVOut1MAdm := 0
nVOut2MAdm := 0
nVOut3MAdm := 0
nVSvT1MAdm := 0
nVSvT2MAdm := 0
nVSvT3MAdm := 0
nOut1MAdm := 0
nOut2MAdm := 0
nOut3MAdm := 0
nGMt1MAdm := 0
nGMt2MAdm := 0
nGMt3MAdm := 0
nEmp1MAdmEE := 0
nEmp2MAdmEE := 0
nEmp3MAdmEE := 0

//Impostos e Contribui��o (Total)

nImp1MCTt := 0
nImp2MCTt := 0
nImp3MCTt := 0
nV1MINSS := 0
nV2MINSS := 0
nV3MINSS := 0
nV1MFGTS := 0
nV2MFGTS := 0
nV3MFGTS := 0
nV1MCPMF := 0
nV2MCPMF := 0
nV3MCPMF := 0
nV1MIPTU := 0
nV2MIPTU := 0
nV3MIPTU := 0
nV1MIPI := 0
nV2MIPI := 0
nV3MIPI := 0
nV1MICMS := 0
nV2MICMS := 0
nV3MICMS := 0
nV1MISS := 0
nV2MISS := 0
nV3MISS := 0
nV1MIRImp := 0
nV2MIRImp := 0
nV3MIRImp := 0
nV1MImpImport := 0
nV2MImpImport := 0
nV3MImpImport := 0


//Novas Linhas de Valores Acrescentadas - Solicita��o Patr�cia 14/02/2008

nHonor1MAdv := 0
nHonor2MAdv := 0
nHonor3MAdv := 0

nDsp1MJudc := 0
nDsp2MJudc := 0
nDsp3MJudc := 0

nRes1MAcAd := 0
nRes2MAcAd := 0
nRes3MAcAd := 0
      
nFnd1MFixo := 0
nFnd2MFixo := 0
nFnd3MFixo := 0

nPIS1MCOF := 0
nPIS2MCOF := 0
nPIS3MCOF := 0

n1MTribut := 0
n2MTribut := 0
n3MTribut := 0

//Metas de Or�amentos

/* Five Solutions - 28/05/2008 - Comentado para possibilitar uso do novo conceito
 * que � pegar os valores or�ados no cadastro de or�amentos - SE7

nComMtOrc := 0 //Meta Or�amento Com�rcio
nIndMtOrc := 0 //Meta Or�amento Ind�stria
nAdmMtOrc := 0 //Meta Or�amento Administra��o
*/

nCom1MMtOrc := 0 //Meta Or�amento Com�rcio M�s 1
nCom2MMtOrc := 0 //Meta Or�amento Com�rcio M�s 2
nCom3MMtOrc := 0 //Meta Or�amento Com�rcio M�s 3

nInd1MMtOrc := 0 //Meta Or�amento Ind�stria M�s 1
nInd2MMtOrc := 0 //Meta Or�amento Ind�stria M�s 2
nInd3MMtOrc := 0 //Meta Or�amento Ind�stria M�s 3

nAdm1MMtOrc := 0 //Meta Or�amento Administra��o M�s 1
nAdm2MMtOrc := 0 //Meta Or�amento Administra��o M�s 2
nAdm3MMtOrc := 0 //Meta Or�amento Administra��o M�s 3


//Demonstrativo de Resultado

//Fretes
nVl1MsFrete := 0
nVl2MsFrete := 0
nVl3MsFrete := 0

//Faturamento por Condi��o de Pagamento
nFat1AVista := 0
nFat2AVista := 0
nFat3AVista := 0

nFat1APrazo := 0
nFat2APrazo := 0
nFat3APrazo := 0

//Impostos
nIPI1Valor := 0
nIPI2Valor := 0
nIPI3Valor := 0

nICM1Vlor  := 0
nICM2Vlor  := 0
nICM3Vlor  := 0

//CALCULAR PIS (1,65%) E COFINS (7,60%) SOBRE O FATURAMENTO TOTAL DE VENDAS (LIVRO FISCAL DE SA�DA)
nPIS1Valor := 0
nPIS2Valor := 0
nPIS3Valor := 0

nCOF1Valor := 0
nCOF2Valor := 0
nCOF3Valor := 0


//SALARIOS PRODU��O: TODAS AS DESPESAS DA M.O. DA IND�STRIA INCLUINDO A SUA AREA ADMINISTRATIVA 
Private cNtMOINDADM := "100301,100302,100303,100304,100305,100306,100307,"
cNtMOINDADM += "100309,100310,100311,100312,100313,100317,100401,"
cNtMOINDADM += "100402,100403,100404,100405,100406,100407,100409,"
cNtMOINDADM += "100410,100411,100412,100417"


nMPMO1IndAdm := 0
nMPMO2IndAdm := 0
nMPMO3IndAdm := 0

nSalInd1M := 0
nSalInd2M := 0
nSalInd3M := 0

//TODAS AS DESPESAS DO COMERCIAL (EXCLUI A NATUREZA DE EMPR�STIMO E REEMBOLSO A CLIENTE)
//Private cDspCDER := "100201,100202,100203,100204,100205,100206,100207,100208,100209,100210,100211,100213,100214,100215,100216,100222,200201,200202,200203,200204,200205,200206,200207,600101,600102,600303,600401"
Private cDspCDER := "100120,100201,100202,100203,100204,100206,100207,100208,100209,100210,100211,100213,100214,100215,100216,100222,200201,200202,200203,200204,200205,200206,200207,600101,600102,600303,100205,600401"


nDesp1ComDRE := 0
nDesp2ComDRE := 0
nDesp3ComDRE := 0

//TOTAL DE DESPESAS ADMINISTRATIVAS (EXCLUI EMPR�STIMOS, IMPOSTOS(Fica o INSS da Diretoria, FGTS e GRFC) E OUTRAS...)
/*
Private cNtAdmDER := "100101,100102,100103,100104,100105,100106,100107,100108,100109,100110,100111,100112,100114,100116,100117,100118,100119,100122,200101,200102,200103,200104,"
        cNtAdmDER += "200105,200106,200107,200108,200112,200113,200114,200115,200116,200117,200118,200119,200120,200121,200122,200123,200124,200127,200129,200130,200131,"
        cNtAdmDER += "200133,200134,200136,200501,200502,200503,200504,200505,200506,200507,200508,200509,200510,500101,500102,500103,500104,600302,"
        */
Private cNtAdmDER := "100101,100102,100103,100104,100105,100106,100107,100108,100109,100110,100111,100112,100114,"
        cNtAdmDER += "100115,100116,100117,100118,100119,100121,100122,200101,200102,200103,200104,200105,200106,"
        cNtAdmDER += "200107,200108,200112,200113,200114,200115,200116,200117,200118,200119,200120,200121,200122,"
        cNtAdmDER += "200123,200124,200128,200129,200130,200131,200133,200134,200136,200137,200143,200501,200502,"
        cNtAdmDER += "200503,200504,200505,200506,200507,200508,200509,200510,500101,500102,500103,500104,600302,"
        cNtAdmDER += "INSS,SANGRIA,TROCO"
        


nDsp1ADMDER := 0
nDsp2ADMDER := 0
nDsp3ADMDER := 0

//DESPESAS FINANCEIRAS '200141','200126','D.ESCONT'
Private cNDpFinDER := "200141,200126,D.ESCONT"

nDspFi1nsDER := 0
nDspFi2nsDER := 0
nDspFi3nsDER := 0

//RECEITAS FINANCEIRAS '200140'
Private cNRctFinDRE := "200140"

nRecDR1EFin := 0
nRecDR2EFin := 0
nRecDR3EFin := 0


//Variaveis que guardam Anos para Verificar Titulos Vencidos
Private cAnoAtu    := ""
Private cAno1Menos := ""
Private cAno2Menos := ""
Private cAnoAntes  := ""


//Defini��es das Metas do Or�amento Comercial
/*
nMetasVlr := 0
If SM0->M0_CODFIL $ "01|06"
	nMetasVlr := (12950)
ElseIf SM0->M0_CODFIL == "02"
	nMetasVlr := (32800)
ElseIf SM0->M0_CODFIL == "03"
	nMetasVlr := (11700)
EndIf	
*/

Private cString := "SA1"

dbSelectArea("SA1")
dbSetOrder(1)

AjustaSx1()
pergunte(cPerg,.F.)

//���������������������������������������������������������������������Ŀ
//� Monta a interface padrao com o usuario...                           �
//�����������������������������������������������������������������������

wnrel := SetPrint(cString,NomeProg,cPerg,@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,.F.,Tamanho,,.F.)

//���������������������������������������������������������������������Ŀ
//� Grupo de Perguntas - RGERCO                                         �
//� MV_PAR01 - Data De                                                  �
//� MV_PAR02 - Data At�                                                 �
//� MV_PAR03 - Custo - (1.M�dio, 2.Reposi��o                            �
//� MV_PAR04 - Anal�tico - (1.Sim, 2.N�o)                               �
//� MV_PAR05 - Depto - (1.Todos,2.Faturamento,3.Ind�stria,4.Financeiro  �
//�����������������������������������������������������������������������



cAlmoxMP  := Alltrim(MV_PAR10)
cAlmoxPA  := Alltrim(MV_PAR11)

/* Five Solutions - 28/05/2008 - Comentado para possibilitar uso do novo conceito
 * que � pegar os valores or�ados no cadastro de or�amentos - SE7
 
nComMtOrc := MV_PAR07 //Meta Or�amento Com�rcio
nIndMtOrc := MV_PAR08 //Meta Or�amento Ind�stria
nAdmMtOrc := MV_PAR09 //Meta Or�amento Administra��o
*/

dDataDe  := MV_PAR01
dDataAte := MV_PAR02
nDeptos  := MV_PAR05


lComercial := IIF(nDeptos == 1 .Or. nDeptos == 2,.T.,.F.)
lIndustria := IIF(nDeptos == 1 .Or. nDeptos == 3,.T.,.F.)
lAdmFin    := IIF(nDeptos == 1 .Or. nDeptos == 4,.T.,.F.)
lLogistica := IIF(nDeptos == 1 .Or. nDeptos == 5,.T.,.F.)

lDemonsRes := IIF(MV_PAR06 == 1,.T.,.F.)

titulo       := "Resumo Gerencial COMAFAL Per�odo "+DtoC(dDataDe)+" a "+DtoC(dDataAte)+" Emp/Fil: "+Alltrim(SM0->M0_NOME)+"/"+Alltrim(SM0->M0_FILIAL)

n2ResMes := Val(Substr(DTOS(dDataDe),1,6)) + 1
d2AnoMes := Alltrim(Str(n2ResMes))
If Right(Alltrim(Str(n2ResMes)),2) == "13"
   cAno := Val(Substr(DTOS(dDataDe),1,4)) + 1
   d2AnoMes   := Alltrim(Str(cAno))+"01"
EndIf

n3ResMes := Val(Substr(DTOS(dDataDe),1,6)) + 2
d3AnoMes   := Alltrim(Str(n3ResMes))

If Right(Alltrim(Str(n3ResMes)),2) == "13"
   cAno := Val(Substr(DTOS(dDataDe),1,4)) + 1
   d3AnoMes   := Alltrim(Str(cAno))+"01"
EndIf


//Par�metros de Datas - 15/02/2008

dData02 := CTOD("01/"+Substr(d2AnoMes,5,2)+"/"+Substr(d2AnoMes,1,4))
dData03 := CTOD("01/"+Substr(d3AnoMes,5,2)+"/"+Substr(d3AnoMes,1,4))


If nLastKey == 27
	Return
Endif

SetDefault(aReturn,cString)

If nLastKey == 27
   Return
Endif

nTipo := If(aReturn[4]==1,15,18)


//���������������������������������������������������������������������Ŀ
//� Processamento para sele��o dos registros p/ impress�o.              �
//�����������������������������������������������������������������������
Processa({||RunSelRG()},"Selecionando Informa��es Gerenciais ...")

//���������������������������������������������������������������������Ŀ
//� Processamento. RPTSTATUS monta janela com a regua de processamento. �
//�����������������������������������������������������������������������

RptStatus({|| RunReport(Cabec1,Cabec2,Titulo,nLin) },Titulo)
Return

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Fun��o    �RUNREPORT � Autor � AP6 IDE            � Data �  05/01/08   ���
�������������������������������������������������������������������������͹��
���Descri��o � Funcao auxiliar chamada pela RPTSTATUS. A funcao RPTSTATUS ���
���          � monta a janela com a regua de processamento.               ���
�������������������������������������������������������������������������͹��
���Uso       � Programa principal                                         ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/

Static Function RunReport(Cabec1,Cabec2,Titulo,nLin)

Local nOrdem

dbSelectArea(cString)
dbSetOrder(1)

//Processa({||RunSelRG()},"Selecionando Informa��es Gerenciais ...")



//���������������������������������������������������������������������Ŀ
//� SETREGUA -> Indica quantos registros serao processados para a regua �
//�����������������������������������������������������������������������

//SetRegua(RecCount())

//dbGoTop()
lRoda := .T.
While lRoda//!EOF()

   //���������������������������������������������������������������������Ŀ
   //� Verifica o cancelamento pelo usuario...                             �
   //�����������������������������������������������������������������������

   If lAbortPrint
      @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
      Exit
   Endif

   //���������������������������������������������������������������������Ŀ
   //� Impressao do cabecalho do relatorio. . .                            �
   //�����������������������������������������������������������������������

   If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 8
   Endif

   If lComercial
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|COMERCIAL"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //                       10        20        30        40        50        60        70        80        90       100       110       120       130       140       150       160       170       180       190       200       210       220
      //              01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
      //              |                                   |                             |                             |                              |      |MERCADO EXTERNO              |                              |
      //                                                                                                                                                    |MES 01                       |MES 02                        |MES 03
      //                                                                999,999,999,999               999,999,999,999                999,999,999,999                      999,999,999,999                999,999,999,999           999,999,999,999         
      //                                                                         999.99                        999.99                         999.99                               999.99                         999.99                    999.99
      //                                                                 999,999,999.99                999,999,999.99                 999,999,999.99                       999,999,999.99                 999,999,999.99            999,999,999.99 
      @nLin,000 PSAY "|MERCADO INTERNO                     "
      @nLin,134 PSAY "|MERCADO EXTERNO                     "
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //                       10        20        30        40
      //              01234567890123456789012345678901234567890
      //                       10        20        30        40        50        60        70        80        90       100       110       120       130       140       150       160       170       180       190       200       210       220
      //              01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
      //              |                                   |                             |                             |                              |      |MERCADO EXTERNO              |                              |
      //                                                                                                                                                    |MES 01                       |MES 02                        |MES 03
      //                                                                999,999,999,999               999,999,999,999                999,999,999,999                      999,999,999,999                999,999,999,999           999,999,999,999         

      @nLin,000 PSAY "|MES                                |"
      @nLin,037 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,066 PSAY "|"
      @nLin,067 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,096 PSAY "|"
      @nLin,097 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,135 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,164 PSAY "|"
      @nLin,165 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,195 PSAY "|"
      @nLin,196 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|""
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|CLIENTES                           |"
      @nLin,050 PSAY  (nClAtv1Mes + nClIna1Mes ) Picture "@E 999,999,999,999"
      @nLin,066 PSAY "|"
      @nLin,080 PSAY  (nClAtv2Mes + nClIna2Mes ) Picture "@E 999,999,999,999"
      @nLin,096 PSAY "|"
      @nLin,111 PSAY  (nClAtv3Mes + nClIna3Mes ) Picture "@E 999,999,999,999"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Prospects                       |"
      @nLin,050 PSAY nASC1Mes Picture "@E 999,999,999,999" 
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nASC2Mes Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nASC3Mes Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //                       10        20        30        40
      //              01234567890123456789012345678901234567890
      //                       10        20        30        40        50        60        70        80        90       100       110       120       130       140       150       160       170       180       190       200       210       220
      //              01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
      //              |                                   |                             |                             |                              |      |MERCADO EXTERNO              |                              |
      //                                                                                                                                                    |MES 01                       |MES 02                        |MES 03
      //                                                                999,999,999,999               999,999,999,999                999,999,999,999                      999,999,999,999                999,999,999,999           999,999,999,999         
      //                                                                         999.99                        999.99                         999.99                               999.99                         999.99                    999.99
      @nLin,000 PSAY "| * Novos                           |"
      @nLin,050 PSAY nCliN01Mes Picture "@E 999,999,999,999" 
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nCliN02Mes Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nCliN03Mes Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Ativos                          |"
      @nLin,050 PSAY nClAtv1Mes Picture "@E 999,999,999,999" 
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nClAtv2Mes Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nClAtv3Mes Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Inativos                        |"
      @nLin,050 PSAY nClIna1Mes Picture "@E 999,999,999,999" 
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nClIna2Mes Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nClIna3Mes Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Atendidos                       |"
      @nLin,050 PSAY nClAtd1Mes Picture "@E 999,999,999,999" 
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nClAtd2Mes Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nClAtd3Mes Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //                       10        20        30        40        50        60        70        80        90       100       110       120       130       140       150       160       170       180       190       200       210       220
      //              01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
      //              |                                   |                             |                             |                              |      |MERCADO EXTERNO              |                              |
      //                                                                                                                                                    |MES 01                       |MES 02                        |MES 03
      //                                                                999,999,999,999               999,999,999,999                999,999,999,999                      999,999,999,999                999,999,999,999           999,999,999,999         
      //                                                                         999.99                        999.99                         999.99                               999.99                         999.99                    999.99

      @nLin,000 PSAY "| * % Atendimento                   |"      //"@E 999,999,999,999" 
      @nLin,059 PSAY (nClAtd1Mes / (nClAtv1Mes + nClIna1Mes))*100 Picture "@E 999.99" 
      @nLin,066 PSAY "|"
      @nLin,089 PSAY (nClAtd2Mes / (nClAtv2Mes + nClIna2Mes))*100 Picture "@E 999.99" 
      @nLin,096 PSAY "|"
      @nLin,120 PSAY (nClAtd3Mes / (nClAtv3Mes + nClIna3Mes))*100 Picture "@E 999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,157 PSAY  (0) Picture "@E 999.99"
      @nLin,164 PSAY "|"
      @nLin,188 PSAY  (0) Picture "@E 999.99"
      @nLin,195 PSAY "|"
      @nLin,211 PSAY  (0) Picture "@E 999.99"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|PEDIDOS EMITIDOS                    "
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Quantidade (M�s)                |"
      /* Teste - Itacolomy Mariano - 15/01/2007
      @nLin,050 PSAY nPed1MesQt Picture "@E 999,999,999,999"
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nPed2MesQt Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nPed3MesQt Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"
      */
      @nLin,050 PSAY (nFatPd1QtMes+nPdPen1QtMes) Picture "@E 999,999,999,999"
      @nLin,066 PSAY "|"
      @nLin,080 PSAY (nFatPd2QtMes+nPdPen2QtMes) Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY (nFatPd3QtMes+nPdPen3QtMes) Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //                       10        20        30        40        50        60        70        80        90       100       110       120       130       140       150       160       170       180       190       200       210       220
      //              01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
      //              |                                   |                             |                             |                              |      |MERCADO EXTERNO              |                              |
      //                                                                                                                                                    |MES 01                       |MES 02                        |MES 03
      //                                                                999,999,999,999               999,999,999,999                999,999,999,999                      999,999,999,999                999,999,999,999           999,999,999,999         
      //                                                                         999.99                        999.99                         999.99                               999.99                         999.99                    999.99
      //                                                                 999,999,999.99                999,999,999.99                 999,999,999.99                       999,999,999.99                 999,999,999.99            999,999,999.99 
      @nLin,000 PSAY "| * Valor R$ (M�s)                  |"
      /* Teste - Itacolomy Mariano - 15/01/2007
      @nLin,051 PSAY nPed1VlMes Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nPed2VlMes Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nPed3VlMes Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      */
      @nLin,051 PSAY (nFtPd1VlrMes+nPdPen1VlrM) Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY (nFtPd2VlrMes+nPdPen2VlrM) Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY (nFtPd3VlrMes+nPdPen3VlrM) Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|PEDIDOS FATURADOS                   "
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Quantidade (M�s)                |"
      @nLin,050 PSAY nFatPd1QtMes Picture "@E 999,999,999,999" 
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nFatPd2QtMes Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nFatPd3QtMes Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Valor R$ (M�s)                  |"
      @nLin,051 PSAY nFtPd1VlrMes Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nFtPd2VlrMes Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nFtPd3VlrMes Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|PEDIDOS PENDENTES                   "
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Quantidade (M�s)                |"
      @nLin,050 PSAY nPdPen1QtMes Picture "@E 999,999,999,999" 
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nPdPen2QtMes Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nPdPen3QtMes Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Valor R$ (M�s)                  |"
      @nLin,051 PSAY nPdPen1VlrM Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nPdPen2VlrM Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nPdPen3VlrM Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|PEDIDOS FATURADOS                   "
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Quantidade (Meses Ant.)         |"
      @nLin,050 PSAY nFtMAPd1Qt Picture "@E 999,999,999,999" 
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nFtMAPd2Qt Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nFtMAPd3Qt Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Valor R$ (Meses Ant.)           |"
      @nLin,051 PSAY nFtPd1MAVl Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nFtPd2MAVl Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nFtPd3MAVl Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|PEDIDOS PENDENTES                   "
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Quantidade (Meses Ant.)         |"
      @nLin,050 PSAY nPdPen1MAQt Picture "@E 999,999,999,999" 
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nPdPen2MAQt Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nPdPen3MAQt Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Valor R$ (Meses Ant.)           |"
      @nLin,051 PSAY nPdPMA1VlM Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nPdPMA2VlM Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nPdPMA3VlM Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|QTD. TOTAL PEDIDOS FATURADOS       |"
      @nLin,050 PSAY  (nFatPd1QtMes + nFtMAPd1Qt) Picture "@E 999,999,999,999"
      @nLin,066 PSAY "|"
      @nLin,080 PSAY  (nFatPd2QtMes + nFtMAPd2Qt) Picture "@E 999,999,999,999"
      @nLin,096 PSAY "|"
      @nLin,111 PSAY  (nFatPd3QtMes + nFtMAPd3Qt) Picture "@E 999,999,999,999"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif                  
      @nLin,000 PSAY "|VALOR TOTAL PEDIDOS FATURADOS      |"
      @nLin,051 PSAY (nFtPd1VlrMes + nFtPd1MAVl) Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY (nFtPd2VlrMes + nFtPd2MAVl) Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY (nFtPd3VlrMes + nFtPd3MAVl) Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif                  
      @nLin,000 PSAY "|QTD. TOTAL PEDIDOS PENDENTES       |"
      @nLin,050 PSAY  (nPdPen1QtMes + nPdPen1MAQt) Picture "@E 999,999,999,999"
      @nLin,066 PSAY "|"
      @nLin,080 PSAY  (nPdPen2QtMes + nPdPen2MAQt) Picture "@E 999,999,999,999"
      @nLin,096 PSAY "|"
      @nLin,111 PSAY  (nPdPen3QtMes + nPdPen3MAQt) Picture "@E 999,999,999,999"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif                  
      @nLin,000 PSAY "|VALOR TOTAL PEDIDOS PENDENTES      |"
      @nLin,051 PSAY (nPdPen1VlrM + nPdPMA1VlM) Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY (nPdPen2VlrM + nPdPMA2VlM) Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY (nPdPen3VlrM + nPdPMA3VlM) Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"

      //For�a Quebra de P�gina
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 8
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|COMERCIAL (continua��o ...)"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MERCADO INTERNO                     "
      @nLin,134 PSAY "|MERCADO EXTERNO                     "
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //                       10        20        30        40
      //              01234567890123456789012345678901234567890
      @nLin,000 PSAY "|MES                                |"
      @nLin,037 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,066 PSAY "|"
      @nLin,067 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,096 PSAY "|"
      @nLin,097 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,135 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,164 PSAY "|"
      @nLin,165 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,195 PSAY "|"
      @nLin,196 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|FATURAMENTO BRUTO                  |"
      @nLin,051 PSAY  (nFat1MesBru) Picture "@E 999,999,999.99" //Faturamento Bruto menos Faturamento de Sucata
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  (nFat2MesBru) Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  (nFat3MesBru) Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|CONDI��O DE FATURAMENTO"
      @nLin,219 PSAY "|"            
      
      For nFtBru := 1 To Len(aCFTList)
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|"+Substr(Posicione("SE4",1,xFilial("SE4")+aCFTList[nFtBru,1],"E4_DESCRI"),1,33)
         @nLin,036 PSAY "|"
         @nLin,051 PSAY  aCFTList[nFtBru,3] Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  aCFTList[nFtBru,4] Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  aCFTList[nFtBru,5] Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"            
      
      Next nFtBru
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif

      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|DEVOLU��O"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| *Valor                            |"
      @nLin,051 PSAY  nVl1MesDev Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nVl2MesDev Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nVl3MesDev Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| *Custo                            |"
      @nLin,051 PSAY  nCus1MesDev Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nCus2MesDev Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nCus3MesDev Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      n1MesFtLiq := (nFat1MesBru - nVl1MesDev)
      n2MesFtLiq := (nFat2MesBru - nVl2MesDev)
      n3MesFtLiq := (nFat3MesBru - nVl3MesDev)
      @nLin,000 PSAY "|FATURAMENTO L�QUIDO                |"
      @nLin,051 PSAY  n1MesFtLiq Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  n2MesFtLiq Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  n3MesFtLiq Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      nCMP1Mes := ((nCus1MesFat+nVlM1CusSUC) - nCus1MesDev)
      nCMP2Mes := ((nCus2MesFat+nVlM2CusSUC) - nCus2MesDev)
      nCMP3Mes := ((nCus3MesFat+nVlM3CusSUC) - nCus3MesDev)
      @nLin,000 PSAY "|CUSTO M.P.                         |"
      @nLin,051 PSAY  nCMP1Mes Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nCMP2Mes Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nCMP3Mes Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      nMg1Mes := (n1MesFtLiq / nCMP1Mes)
      nMg2Mes := (n2MesFtLiq / nCMP2Mes)
      nMg3Mes := (n3MesFtLiq / nCMP3Mes)
      @nLin,000 PSAY "|MARGEM                             |"
      @nLin,051 PSAY  ((nMg1Mes)-1)*100 Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  ((nMg2Mes)-1)*100 Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  ((nMg3Mes)-1)*100 Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      
      nProdTon1Mes := (nVlr1MesBx/nProd1Mes)      
      nProdTon2Mes := (nVlr2MesBx/nProd2Mes) 
      nProdTon3Mes := (nVlr3MesBx/nProd3Mes)
      
      n1MTonelVen  := (nQtTon1MVen+nQtTon1MAPd)
      n2MTonelVen  := (nQtTon2MVen+nQtTon2MAPd)
      n3MTonelVen  := (nQtTon3MVen+nQtTon3MAPd)
      
      //Five Solutions - 14/07/2008
      //Custo da Mat�ria Prima = Custo do Faturamento somado ao Custo de Faturamento Sucata dividido pela
      //quantidade de toneladas vendidas.
     
      nMatPrim1Mes :=((nCus1MesFat+nVlM1CusSUC)/n1MTonelVen)
      nMatPrim2Mes :=((nCus2MesFat+nVlM2CusSUC)/n2MTonelVen)
      nMatPrim3Mes :=((nCus3MesFat+nVlM3CusSUC)/n3MTonelVen)
      
      
      nCMTon1Mes := (nProdTon1Mes+nMatPrim1Mes)
      nCMTon2Mes := (nProdTon2Mes+nMatPrim2Mes) 
      nCMTon3Mes := (nProdTon3Mes+nMatPrim3Mes) 
      
      //For�a quebra de p�gina
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 8
      
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      @nLin,000 PSAY "|COMERCIAL (continua��o ...)"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MERCADO INTERNO                     "
      @nLin,134 PSAY "|MERCADO EXTERNO                     "
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //                       10        20        30        40
      //              01234567890123456789012345678901234567890
      @nLin,000 PSAY "|MES                                |"
      @nLin,037 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,066 PSAY "|"
      @nLin,067 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,096 PSAY "|"
      @nLin,097 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,135 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,164 PSAY "|"
      @nLin,165 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,195 PSAY "|"
      @nLin,196 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
     
      //PRE�O M�DIO (Venda) = FATURAMENTO BRUTO ( / ) DIVIDIDO PELA QTD. DE TONELADAS VENDIDAS 
      nPMVen1Mes := (nFat1MesBru/n1MTonelVen)
      nPMVen2Mes := (nFat2MesBru/n2MTonelVen)
      nPMVen3Mes := (nFat3MesBru/n3MTonelVen)
      
      @nLin,000 PSAY "|PRE�O M�DIO (Venda)                |"
      @nLin,051 PSAY nPMVen1Mes  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nPMVen2Mes Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nPMVen3Mes Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|CUSTO M�DIO (Ton)                  |"
      @nLin,051 PSAY  nCMTon1Mes Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nCMTon2Mes Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nCMTon3Mes Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "| * Produ��o (Ton)                  |"
      @nLin,051 PSAY  nProdTon1Mes Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nProdTon2Mes Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nProdTon3Mes Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      
      //Mat�ria Prima (Ton) = (CUSTO APRESENTADO NO RELAT�RIO DE FATURAMENTO "NELSON") / (QUANTIDADE DE TONELADAS VENDIDAS) 
      
      @nLin,000 PSAY "| * Mat�ria Prima (Ton)             |"
      @nLin,051 PSAY  nMatPrim1Mes Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nMatPrim2Mes Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nMatPrim3Mes Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      
      //MARGEM M�DIA = (PRE�O MEDIO) / (CUSTO M�DIO)
      nMgMed1Mes := (nPMVen1Mes/nCMTon1Mes) 
      nMgMed2Mes := (nPMVen2Mes/nCMTon2Mes) 
      nMgMed3Mes := (nPMVen3Mes/nCMTon3Mes) 
      
      @nLin,000 PSAY "|MARGEM M�DIA                       |"
      @nLin,051 PSAY ((nMgMed1Mes)-1)*100  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY ((nMgMed2Mes)-1)*100 Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY ((nMgMed3Mes)-1)*100 Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|COMISS�O"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "| * Valor R$                        |"  //VALOR DA COMISS�O '100203'
      @nLin,051 PSAY  nVl1MComis  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nVl2MComis Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nVl3MComis Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * Percentual (%)                 |"  //VALOR DA COMISSAO (/) DIVIDIDO PELO FATURAMENTO LIQUIDO
      @nLin,051 PSAY  (nVl1MComis/n1MesFtLiq)*100  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  (nVl2MComis/n2MesFtLiq)*100 Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  (nVl3MComis/n3MesFtLiq)*100 Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      
      //Five Solutions, 11/07/2008 - Tratamento para imprimir indicadores apenas se os mesmos tiverem valor.
      //Solicita��o: Gilmar Filho / Patr�cia.
      
      //If (nVl1MGGComl+nVl2MGGComl+nVl3MGGComl) > 0
            
         @nLin,000 PSAY "|GASTOS GERAIS COML - MI            |"  //TOTAL DE GASTOS DA �REA COMERCIAL
         @nLin,051 PSAY  nVl1MGGComl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  nVl2MGGComl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  nVl3MGGComl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                  
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      
      //EndIf
      
      If (nVl1MGGMO+nVl2MGGMO+nVl3MGGMO) > 0
      
         @nLin,000 PSAY "| * M�O DE OBRA                     |" //SOMAT�RIO DE TODAS AS NATUREZAS COM GASTOS REFERENTE A M�O DE OBRA
         @nLin,051 PSAY  nVl1MGGMO  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  nVl2MGGMO Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  nVl3MGGMO Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                  
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      
      EndIf
      
      If (nVl1MDFGG+nVl2MDFGG+nVl3MDFGG) > 0
      
         @nLin,000 PSAY "|   ** Desp. Fixa                   |" //
         @nLin,051 PSAY  nVl1MDFGG  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  nVl2MDFGG Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  nVl3MDFGG Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                  
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      
      EndIf
      
      If (nVl1MDPGG+nVl2MDPGG+nVl3MDPGG) > 0
      
         @nLin,000 PSAY "|   ** Desp. Period.                |" //
         @nLin,051 PSAY  nVl1MDPGG  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  nVl2MDPGG Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  nVl3MDPGG Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                  
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      
      EndIf
      
      If (nVl1MResGG+nVl2MResGG+nVl3MResGG) > 0
      
         @nLin,000 PSAY "|   ** Rescis�es                    |" //
         @nLin,051 PSAY  nVl1MResGG  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  nVl2MResGG Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  nVl3MResGG Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                  
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      
      EndIf
      
      If (nVl1MPrmGG+nVl2MPrmGG+nVl3MPrmGG) > 0
            
         @nLin,000 PSAY "|   ** Premia��es                   |" //
         @nLin,051 PSAY  nVl1MPrmGG  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  nVl2MPrmGG Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  nVl3MPrmGG Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                  
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      
      EndIf
      
      If (nVl1MComis+nVl2MComis+nVl3MComis) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      
         @nLin,000 PSAY "| * COMISS�O                        |" //
         @nLin,051 PSAY  nVl1MComis  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  nVl2MComis Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  nVl3MComis Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                  
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      
      EndIf
      
      If (nVl1MMktGG+nVl2MMktGG+nVl3MMktGG) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "| * MARKETING                       |" //
         @nLin,051 PSAY  nVl1MMktGG  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  nVl2MMktGG Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  nVl3MMktGG Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                  
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      
      EndIf
      
      If (nVl1MDVgGG+nVl2MDVgGG+nVl3MDVgGG) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "| * DESPESAS DE VIAGENS             |" //
         @nLin,051 PSAY  nVl1MDVgGG  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  nVl2MDVgGG Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  nVl3MDVgGG Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                  
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      
      EndIf
      
      If (nVl1MDFrtGG+nVl2MDFrtGG+nVl3MDFrtGG) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "| * FRETE DE VENDAS                 |" //
         @nLin,051 PSAY  nVl1MDFrtGG  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  nVl2MDFrtGG Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  nVl3MDFrtGG Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                  
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      
      EndIf
      
      If (nVl1MOutDGG+nVl2MOutDGG+nVl3MOutDGG) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "| * OUTRAS DESPESAS COM.            |" //
         @nLin,051 PSAY  nVl1MOutDGG  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  nVl2MOutDGG Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  nVl3MOutDGG Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                  
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      
      EndIf
      
      If (nVl1MONtGG+nVl2MONtGG+nVl3MONtGG) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "| * OUTRAS NATUREZAS                |" //
         @nLin,051 PSAY  nVl1MONtGG  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  nVl2MONtGG Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  nVl3MONtGG Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                  
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      
      EndIf
      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|VALOR META OR�AMENTO COML.         |" //
      @nLin,051 PSAY  nCom1MMtOrc  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nCom2MMtOrc Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nCom3MMtOrc Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|GASTOS META OR�AMENTO COML.        |" //
      @nLin,051 PSAY  nGst1MMtOrC  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nGst2MMtOrC Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nGst3MMtOrC Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      
      //For�a Quebra de P�gina
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 8
      
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|COMERCIAL (continua��o ...)"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MERCADO INTERNO                     "
      @nLin,134 PSAY "|MERCADO EXTERNO                     "
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MES                                |"
      @nLin,037 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,066 PSAY "|"
      @nLin,067 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,096 PSAY "|"
      @nLin,097 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,135 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,164 PSAY "|"
      @nLin,165 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,195 PSAY "|"
      @nLin,196 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|FATURAMENTO SUCATA                 |" //
      @nLin,051 PSAY  nVl1SUCMFat  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nVl2SUCMFat Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nVl3SUCMFat Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|CONDI��O DE FATURAMENTO            |" //
      
      For nFtSUC := 1 To Len(aTSUCList)
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|"+Substr(Posicione("SE4",1,xFilial("SE4")+aCFTList[nFtSUC,1],"E4_DESCRI"),1,33)
         @nLin,036 PSAY "|"
         @nLin,051 PSAY  aTSUCList[nFtSUC,3] Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  aTSUCList[nFtSUC,4] Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  aTSUCList[nFtSUC,5] Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"            
      Next nFtSUC
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|CUSTO SUCATA                       |" //
      @nLin,051 PSAY  nVlM1CusSUC  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nVlM2CusSUC Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nVlM3CusSUC Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      nMg1SUCMes := (nVl1SUCMFat / nVlM1CusSUC)
      nMg2SUCMes := (nVl2SUCMFat / nVlM2CusSUC)
      nMg3SUCMes := (nVl3SUCMFat / nVlM3CusSUC)
      @nLin,000 PSAY "|MARGEM SUCATA                      |"
      @nLin,051 PSAY  ((nMg1SUCMes)-1)*100 Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  ((nMg2SUCMes)-1)*100 Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  ((nMg3SUCMes)-1)*100 Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
     
      //PRE�O M�DIO) = FATURAMENTO ( / ) DIVIDIDO PELA QTD. DE TONELADAS
      nPMV1SUCM := (nVl1SUCMFat/nQt1MSUCVen)
      nPMV2SUCM := (nVl2SUCMFat/nQt2MSUCVen)
      nPMV3SUCM := (nVl3SUCMFat/nQt3MSUCVen)
      
      @nLin,000 PSAY "|PRE�O M�DIO                        |"
      @nLin,051 PSAY nPMV1SUCM  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nPMV2SUCM Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nPMV3SUCM Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|QTD. SUCATA VENDIDA                |"
      @nLin,050 PSAY nQt1MSUCVen Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nQt2MSUCVen Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nQt3MSUCVen Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999.99"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999.99"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999.99"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      
   EndIf //Fim da Condi��o para impress�o dos dados da An�lise Gerencial Comercial

   If lIndustria
   
      If lComercial
         //For�a Quebra de P�gina
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
         @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      EndIf
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|IND�STRIA"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MERCADO INTERNO                     "
      @nLin,134 PSAY "|MERCADO EXTERNO                     "
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //                       10        20        30        40
      //              01234567890123456789012345678901234567890
      @nLin,000 PSAY "|MES                                |"
      @nLin,037 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,066 PSAY "|"
      @nLin,067 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,096 PSAY "|"
      @nLin,097 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,135 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,164 PSAY "|"
      @nLin,165 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,195 PSAY "|"
      @nLin,196 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|QTD. TONELADAS"
      @nLin,219 PSAY"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Vendidas                        |"      
      @nLin,051 PSAY n1MTonelVen Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY n2MTonelVen Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY n3MTonelVen Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Produzidas                      |"      
      @nLin,051 PSAY nProd1Mes Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nProd2Mes Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nProd3Mes Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MAQUINA                            |PRODUCAO"      


      For nPrMaq := 1 To Len(aMaqPrdM)
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         //@nLin,000 PSAY "|"+Substr(Alltrim(aMaqPrdM[nPrMaq,1])+" "+Posicione("SH1",1,xFilial("SH1")+aMaqPrdM[nPrMaq,1],"H1_DESCRI"),1,29)
         @nLin,000 PSAY "|"+Substr(Alltrim(aMaqPrdM[nPrMaq,1])+" "+Posicione("CTT",1,xFilial("CTT")+aMaqPrdM[nPrMaq,1],"CTT_DESC01"),1,29)
         @nLin,036 PSAY "|"
         @nLin,051 PSAY  aMaqPrdM[nPrMaq,2] Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY  aMaqPrdM[nPrMaq,3] Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY  aMaqPrdM[nPrMaq,4] Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"            
      Next nPrMaq
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MAQUINA                            |PARADA"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MAQUINA                            |PERDA"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MATERIA PRIMA"
      @nLin,219 PSAY"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Qtd. Estoque                    |"      
//      Alert("Imprimindo o Valor nQtdFM1MEst "+Alltrim(Str(nQtdFM1MEst)))
      @nLin,050 PSAY nQtdFM1MEst Picture "@E 999,999,999,999" 
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nQtdFM2MEst Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nQtdFM3MEst Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"           
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Valor Estoque                   |"
      @nLin,051 PSAY nVlrFM1MEst  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nVlrFM2MEst Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nVlrFM3MEst Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      nVM1MMP := (nVlrFM1MEst/nQtdFM1MEst)      
      nVM2MMP := (nVlrFM2MEst/nQtdFM2MEst)      
      nVM3MMP := (nVlrFM3MEst/nQtdFM3MEst)      
      @nLin,000 PSAY "| * Valor M�dio MP (Ton)            |"
      @nLin,050 PSAY nVM1MMP  Picture "@E 999,999,999.999"
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nVM2MMP Picture "@E 999,999,999.999"
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nVM3MMP Picture "@E 999,999,999.999"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"


      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MATERIA PRIMA(SLITER)"
      @nLin,219 PSAY"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Qtd. Estoque                    |"      
//      Alert("Imprimindo o Valor nQtdFM1MEst "+Alltrim(Str(nQtdFM1MEst)))
         
      @nLin,050 PSAY nQt1MPSlit Picture "@E 999,999,999,999" 
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nQt2MPSlit Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nQt3MPSlit Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"           
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Valor Estoque                   |"
      @nLin,051 PSAY nVl1MCSlit  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nVl2MCSlit Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nVl3MCSlit Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      
      nSL1TVM := (nVl1MCSlit/nQt1MPSlit)      
      nSL2TVM := (nVl2MCSlit/nQt2MPSlit)      
      nSL3TVM := (nVl3MCSlit/nQt3MPSlit)      
      
      @nLin,000 PSAY "| * Valor M�dio MP (Ton)            |"
      @nLin,050 PSAY nSL1TVM  Picture "@E 999,999,999.999"
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nSL2TVM Picture "@E 999,999,999.999"
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nSL3TVM Picture "@E 999,999,999.999"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"

      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|PRODUTO ACABADO"
      @nLin,219 PSAY"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Qtd. Estoque                    |"      
      @nLin,050 PSAY nQt1PAM Picture "@E 999,999,999,999" 
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nQt2PAM Picture "@E 999,999,999,999" 
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nQt3PAM Picture "@E 999,999,999,999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,148 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,164 PSAY "|"
      @nLin,179 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,195 PSAY "|"
      @nLin,202 PSAY  (0) Picture "@E 999,999,999,999"
      @nLin,219 PSAY "|"           
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Valor Estoque                   |"
      @nLin,051 PSAY nCs1PAM  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nCs2PAM Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nCs3PAM Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      nVM1MPA := (nCs1PAM/nQt1PAM)      
      nVM2MPA := (nCs2PAM/nQt2PAM)
      nVM3MPA := (nCs3PAM/nQt3PAM)
      @nLin,000 PSAY "| * Valor M�dio PA (Ton)            |"
      @nLin,050 PSAY nVM1MPA  Picture "@E 999,999,999.999"
      @nLin,066 PSAY "|"
      @nLin,080 PSAY nVM2MPA Picture "@E 999,999,999.999"
      @nLin,096 PSAY "|"
      @nLin,111 PSAY nVM3MPA Picture "@E 999,999,999.999"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      
      If lComercial
         //For�a Quebra de P�gina
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
         @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      EndIf
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|IND�STRIA (Continua��o)"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MERCADO INTERNO                     "
      @nLin,134 PSAY "|MERCADO EXTERNO                     "
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //                       10        20        30        40
      //              01234567890123456789012345678901234567890
      @nLin,000 PSAY "|MES                                |"
      @nLin,037 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,066 PSAY "|"
      @nLin,067 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,096 PSAY "|"
      @nLin,097 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,135 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,164 PSAY "|"
      @nLin,165 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,195 PSAY "|"
      @nLin,196 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|PRODUTIVIDADE                      |"
      @nLin,051 PSAY nProvd1M  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nProvd2M Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nProvd3M Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|CUSTO DE PRODU��O                  |"
      @nLin,051 PSAY nVlr1CusPro  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nVlr2CusPro Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nVlr3CusPro Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|GASTOS GERAIS IND.                 |"
      @nLin,051 PSAY nVGG1MInd  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nVGG2MInd Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nVGG3MInd Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|IMOBILIZADO(TOTAL)                 |"
      @nLin,051 PSAY nVImob1MTt  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nVImob2MTt Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nVImob3MTt Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //Five Solutions, 11/07/2008 - Tratamento para imprimir indicadores apenas se os mesmos tiverem valor.
      //Solicita��o: Gilmar Filho / Patr�cia.
      If (nVMqMt1MTt+nVMqMt2MTt+nVMqMt3MTt) > 0
      
         @nLin,000 PSAY "| * Maquinas - Motores              |"
         @nLin,051 PSAY nVMqMt1MTt  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVMqMt2MTt Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVMqMt3MTt Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      
      EndIf
      
      If (nVlM1ImpInd+nVlM2ImpInd+nVlM3ImpInd) > 0
      
         @nLin,000 PSAY "| * Importado                       |"
         @nLin,051 PSAY nVlM1ImpInd  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVlM2ImpInd Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVlM3ImpInd Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      EndIf
      
      If (nVlInfM1Ind+nVlInfM2Ind+nVlInfM3Ind) > 0
      
         @nLin,000 PSAY "| * Informatica                     |"
         @nLin,051 PSAY nVlInfM1Ind  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVlInfM2Ind Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVlInfM3Ind Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      EndIf
      
      If (nVlMU1MInd+nVlMU2MInd+nVlMU3MInd) > 0
      
         @nLin,000 PSAY "| * Moveis e Utens�lios             |"
         @nLin,051 PSAY nVlMU1MInd  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVlMU2MInd Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVlMU3MInd Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      
      nVl1MInvTt := (nVl1MEmpInv+nVlMq1MInv+nVlFin1MInv+nVlLea1MInv+nVlBenf1MI+nVlISO1MI)
      nVl2MInvTt := (nVl2MEmpInv+nVlMq2MInv+nVlFin2MInv+nVlLea2MInv+nVlBenf2MI+nVlISO2MI)
      nVl3MInvTt := (nVl3MEmpInv+nVlMq3MInv+nVlFin3MInv+nVlLea3MInv+nVlBenf3MI+nVlISO3MI)
      
      @nLin,000 PSAY "|  INVESTIMENTOS (Total)            |"
      @nLin,051 PSAY nVl1MInvTt  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nVl2MInvTt Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nVl3MInvTt Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      
      If (nVl1MEmpInv+nVl2MEmpInv+nVl3MEmpInv) > 0
      
         @nLin,000 PSAY "|   * Empresas                      |"
         @nLin,051 PSAY nVl1MEmpInv  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVl2MEmpInv Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVl3MEmpInv Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      
      EndIf
      
      If (nVlMq1MInv+nVlMq2MInv+nVlMq3MInv) > 0
      
         @nLin,000 PSAY "|   * M�quinas                      |"
         @nLin,051 PSAY nVlMq1MInv  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVlMq2MInv Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVlMq3MInv Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      EndIf
      
      If (nVlFin1MInv+nVlFin2MInv+nVlFin3MInv) > 0
      
         @nLin,000 PSAY "|   * Finame                        |"
         @nLin,051 PSAY nVlFin1MInv  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVlFin2MInv Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVlFin3MInv Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      EndIf
      
      If (nVlLea1MInv+nVlLea2MInv+nVlLea3MInv) > 0
      
         @nLin,000 PSAY "|   * Leasing                       |"
         @nLin,051 PSAY nVlLea1MInv  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVlLea2MInv Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVlLea3MInv Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      EndIf
      
      
      If (nVlBenf1MI+nVlBenf2MI+nVlBenf3MI) > 0
      
         @nLin,000 PSAY "|   * Benfeitorias                  |"
         @nLin,051 PSAY nVlBenf1MI  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVlBenf2MI Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVlBenf3MI Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      EndIf
      
      If (nVlISO1MI+nVlISO2MI+nVlISO3MI) > 0
      
         @nLin,000 PSAY "|   * ISO 9001 - 2000               |"
         @nLin,051 PSAY nVlISO1MI  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVlISO2MI Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVlISO3MI Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|  MATERIA PRIMA (Total)            |"
      @nLin,051 PSAY nMP1MVlTt  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nMP2MVlTt Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nMP3MVlTt Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      
      If (nNacMP1MVl+nNacMP2MVl+nNacMP3MVl) > 0
      
         @nLin,000 PSAY "|    ** M.P. Nacional               |"
         @nLin,051 PSAY nNacMP1MVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nNacMP2MVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nNacMP3MVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
           nLin := 8
         Endif       
      EndIf
      
      If (nImpMP1MVl+nImpMP2MVl+nImpMP3MVl) > 0
      
         @nLin,000 PSAY "|    ** M.P. Importada              |"
         @nLin,051 PSAY nImpMP1MVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nImpMP2MVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nImpMP3MVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      EndIf
      
      If (nTrbImp1MMP+nTrbImp2MMP+nTrbImp3MMP) > 0
      
         @nLin,000 PSAY "|    ** Tributos Importa��o         |"
         @nLin,051 PSAY nTrbImp1MMP  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nTrbImp2MMP Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nTrbImp3MMP Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      EndIf
      
      If (nDspImp1MMP+nDspImp2MMP+nDspImp3MMP) > 0
      
         @nLin,000 PSAY "|    ** Despesas Importa��o         |"
         @nLin,051 PSAY nDspImp1MMP  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nDspImp2MMP Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nDspImp3MMP Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      EndIf
      
      If (nFrtImp1MMP+nFrtImp2MMP+nFrtImp3MMP) > 0
      
         @nLin,000 PSAY "|    ** Frete Importa��o            |"
         @nLin,051 PSAY nFrtImp1MMP  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nFrtImp2MMP Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nFrtImp3MMP Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|  INSUMOS                          |"
      @nLin,051 PSAY nInsVl1MInd  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nInsVl2MInd Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nInsVl3MInd Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //Itacolomy Mariano - 11/03/2008
      //Ao inv�s de usar query para calcular soma da m�o de obra, somarei as vari�veis
      nMOInd1MVl := (nFixMO1MInd+nPerMO1MInd+nRes1MIndVl+nPrm1IndVl)
      nMOInd2MVl := (nFixMO2MInd+nPerMO2MInd+nRes2MIndVl+nPrm2IndVl)
      nMOInd3MVl := (nFixMO3MInd+nPerMO3MInd+nRes3MIndVl+nPrm3IndVl)
      
      @nLin,000 PSAY "|  M�O DE OBRA                      |"
      @nLin,051 PSAY nMOInd1MVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nMOInd2MVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nMOInd3MVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      
      If (nFixMO1MInd+nFixMO2MInd+nFixMO3MInd) > 0
      
         @nLin,000 PSAY "|    * Desp. Fixa                   |"
         @nLin,051 PSAY nFixMO1MInd  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nFixMO2MInd Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nFixMO3MInd Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      EndIf
      
      If (nPerMO1MInd+nPerMO2MInd+nPerMO3MInd) > 0
      
         @nLin,000 PSAY "|    * Desp. Period.                |"
         @nLin,051 PSAY nPerMO1MInd  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nPerMO2MInd Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nPerMO3MInd Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      EndIf
      
      If (nRes1MIndVl+nRes2MIndVl+nRes3MIndVl) > 0
      
         @nLin,000 PSAY "|    * Rescis�es                    |"
         @nLin,051 PSAY nRes1MIndVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nRes2MIndVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nRes3MIndVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      EndIf
      
      If (nPrm1IndVl+nPrm2IndVl+nPrm3IndVl) > 0
      
         @nLin,000 PSAY "|    * Premia��es                   |"
         @nLin,051 PSAY nPrm1IndVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nPrm2IndVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nPrm3IndVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"

      //For�a Quebra de P�gina
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 8
      
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|IND�STRIA (continua��o ...)"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MERCADO INTERNO                     "
      @nLin,134 PSAY "|MERCADO EXTERNO                     "
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MES                                |"
      @nLin,037 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,066 PSAY "|"
      @nLin,067 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,096 PSAY "|"
      @nLin,097 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,135 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,164 PSAY "|"
      @nLin,165 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,195 PSAY "|"
      @nLin,196 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      
      If (nBnf1IndVl+nBnf2IndVl+nBnf3IndVl) > 0
      
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|  BENEFICIAMENTO                   |"
         @nLin,051 PSAY nBnf1IndVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nBnf2IndVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nBnf3IndVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"            
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nMan1IndVl+nMan2IndVl+nMan3IndVl) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|  MANUTEN��O                       |"
         @nLin,051 PSAY nMan1IndVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nMan2IndVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nMan3IndVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nSTc1IndVl+nSTc2IndVl+nSTc3IndVl) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|  SERVI�OS TERCEIRIZADOS           |"
         @nLin,051 PSAY nSTc1IndVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nSTc2IndVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nSTc3IndVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"            
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nEPI1IndVl+nEPI2IndVl+nEPI3IndVl) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|  EPI                              |"
         @nLin,051 PSAY nEPI1IndVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nEPI2IndVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nEPI3IndVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"            
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nEner1IndVl+nEner2IndVl+nEner3IndVl) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|  ENERGIA                          |"
         @nLin,051 PSAY nEner1IndVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nEner2IndVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nEner3IndVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"            
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nAgua1IndVl+nAgua2IndVl+nAgua3IndVl) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|  �GUA                             |"
         @nLin,051 PSAY nAgua1IndVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nAgua2IndVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nAgua3IndVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                                                      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nAlg1IndVl+nAlg2IndVl+nAlg3IndVl) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|  ALUGUEL                          |"
         @nLin,051 PSAY nAlg1IndVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nAlg2IndVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nAlg3IndVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"            
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nCPr1IndVl+nCPr2IndVl+nCPr3IndVl) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|  CONSERV. PR�DIO                  |"
         @nLin,051 PSAY nCPr1IndVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nCPr2IndVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nCPr3IndVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"            
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nDVg1IndVl+nDVg2IndVl+nDVg3IndVl) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|  DESPESAS DE VIAGENS              |"
         @nLin,051 PSAY nDVg1IndVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nDVg2IndVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nDVg3IndVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"            
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nODp1IndVl+nODp2IndVl+nODp3IndVl) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|  OUTRAS DESPESAS IND.             |"
         @nLin,051 PSAY nODp1IndVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nODp2IndVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nODp3IndVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"            
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nONt1IndVl+nONt2IndVl+nONt3IndVl) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
         @nLin,000 PSAY "|  OUTRAS NATUREZAS                 |"
         @nLin,051 PSAY nONt1IndVl  Picture "@E 999,999,999.99"
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nONt2IndVl Picture "@E 999,999,999.99"
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nONt3IndVl Picture "@E 999,999,999.99"
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"            
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|  VALOR META OR�AMENTO IND.        |"
      @nLin,051 PSAY nInd1MMtOrc  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nInd2MMtOrc Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nInd3MMtOrc Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|  GASTOS METAS OR�AMENTO IND.      |"
      @nLin,051 PSAY nGst1MtOrInd  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nGst2MtOrInd Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nGst3MtOrInd Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"      

   EndIf //Fim da Condi��o para impress�o dos dados da An�lise Gerencial Industrial
   
   If lAdmFin
      
      If lComercial .Or. lIndustria
         //For�a Quebra de P�gina
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
         @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      EndIf
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|ADMINISTRATIVO - FINANCEIRO"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MERCADO INTERNO                     "
      @nLin,134 PSAY "|MERCADO EXTERNO                     "
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //                       10        20        30        40
      //              01234567890123456789012345678901234567890
      @nLin,000 PSAY "|MES                                |"
      @nLin,037 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,066 PSAY "|"
      @nLin,067 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,096 PSAY "|"
      @nLin,097 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,135 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,164 PSAY "|"
      @nLin,165 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,195 PSAY "|"
      @nLin,196 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      nPd1MTrvTt := (nPdTrv1MVlr+nPdTrv1AntM)                  
      nPd2MTrvTt := (nPdTrv2MVlr+nPdTrv2AntM)                  
      nPd3MTrvTt := (nPdTrv3MVlr+nPdTrv3AntM)                  
      
      @nLin,000 PSAY "|PEDIDOS NA TRAVA (Total)           |"
      @nLin,051 PSAY nPd1MTrvTt Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nPd2MTrvTt Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nPd3MTrvTt Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * M�s Atual                      |"
      @nLin,051 PSAY nPdTrv1MVlr Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nPdTrv2MVlr Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nPdTrv3MVlr Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * M�ses Anteriores               |"
      @nLin,051 PSAY nPdTrv1AntM Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nPdTrv2AntM Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nPdTrv3AntM Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      nTit1MTrvTt := (nTit1MTrvVlr+nAntTit1MTrv)                  
      nTit2MTrvTt := (nTit2MTrvVlr+nAntTit2MTrv)
      nTit3MTrvTt := (nTit3MTrvVlr+nAntTit3MTrv)
      
      @nLin,000 PSAY "|TITULOS NA TRAVA (Total)           |"
      @nLin,051 PSAY nTit1MTrvTt Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nTit2MTrvTt Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nTit3MTrvTt Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * M�s Atual                      |"
      @nLin,051 PSAY nTit1MTrvVlr Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nTit2MTrvVlr Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nTit3MTrvVlr Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * M�ses Anteriores               |"
      @nLin,051 PSAY nAntTit1MTrv Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nAntTit2MTrv Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nAntTit3MTrv Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|TITULOS A PAGAR (M�s)              |"
      @nLin,051 PSAY nAPg1MTit Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nAPg2MTit Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nAPg3MTit Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //Ita 03/03/2008 - Novo Conceito p/Pagamentos Realizados
      nPg1MRealiz := (nBx1MComp+nDBT1MCC+nBx1MNorm+nBx1MDevol)
      nPg2MRealiz := (nBx2MComp+nDBT2MCC+nBx2MNorm+nBx2MDevol)
      nPg3MRealiz := (nBx3MComp+nDBT3MCC+nBx3MNorm+nBx3MDevol)
      
      @nLin,000 PSAY "|PAGAMENTOS REALIZADOS (M�s)        |"
      @nLin,051 PSAY nPg1MRealiz Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nPg2MRealiz Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nPg3MRealiz Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|COMPENSA��O                        |"
      @nLin,051 PSAY nBx1MComp Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nBx2MComp Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nBx3MComp Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|DEBITO CC                          |"
      @nLin,051 PSAY nDBT1MCC Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nDBT2MCC Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nDBT3MCC Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|NORMAL                             |"
      @nLin,051 PSAY nBx1MNorm Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nBx2MNorm Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nBx3MNorm Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|DEVOLU��O                          |"
      @nLin,051 PSAY nBx1MDevol Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nBx2MDevol Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nBx3MDevol Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      nPgAb1MVlr := (nSld1MPag+nSld1MAnt)
      nPgAb2MVlr := (nSld2MPag+nSld2MAnt)
      nPgAb3MVlr := (nSld3MPag+nSld3MAnt)
      
      @nLin,000 PSAY "|PAGAMENTOS EM ABERTO (Total)       |"
      @nLin,051 PSAY nPgAb1MVlr Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nPgAb2MVlr Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nPgAb3MVlr Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * M�s Atual                      |"
      @nLin,051 PSAY nSld1MPag Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nSld2MPag Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nSld3MPag Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * M�ses Anteriores               |"
      @nLin,051 PSAY nSld1MAnt Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nSld2MAnt Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nSld3MAnt Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      
      //Titulos A Pagar Total
      
      /* Five Solutions Consultoria
         20 de agosto de 2008
      nTt1MAPG := (nAPg1M1a30+nAPg1M31a60+nAPg1M61a90+nAPg1M90M)
      nTt2MAPG := (nAPg2M1a30+nAPg2M31a60+nAPg2M61a90+nAPg2M90M)
      nTt3MAPG := (nAPg3M1a30+nAPg3M31a60+nAPg3M61a90+nAPg3M90M)
      */
      nTt1MAPG := (nAPg1M1a30+nAPg1M31a60+nAcm1M60dd)
      nTt2MAPG := (nAPg2M1a30+nAPg2M31a60+nAcm2M60dd)
      nTt3MAPG := (nAPg3M1a30+nAPg3M31a60+nAcm3M60dd)
      
      @nLin,000 PSAY "|TITULOS A PAGAR (Total)            |"
      @nLin,051 PSAY nTt1MAPG Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nTt2MAPG Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nTt3MAPG Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * 01 a 30 dias                   |"
      @nLin,051 PSAY nAPg1M1a30 Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nAPg2M1a30 Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nAPg3M1a30 Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * 31 a 60 dias                   |"
      @nLin,051 PSAY nAPg1M31a60 Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nAPg2M31a60 Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nAPg3M31a60 Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
//AkiIta
      @nLin,000 PSAY "|  * acima de 60 dias               |"
      @nLin,051 PSAY nAcm1M60dd  Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nAcm2M60dd Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nAcm3M60dd Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      

/*
      @nLin,000 PSAY "|  * 61 a 90 dias                   |"
      @nLin,051 PSAY nAPg1M61a90 Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nAPg2M61a90 Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nAPg3M61a90 Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * acima de 90 dias               |"
      @nLin,051 PSAY nAPg1M90M Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nAPg2M90M Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nAPg3M90M Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
*/
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|TITULOS A RECEBER (M�s)            |"
      @nLin,051 PSAY nARe1MTit Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nARe2MTit Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nARe3MTit Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      
      //Ita 03/03/2008 - Novo Conceito p/ Calcular Titulos Recebidos
      
      nRe1MRealiz := (nBxRE1MComp+nBx1MDvRE+nBxRE1MNorm)
      nRe2MRealiz := (nBxRE2MComp+nBx2MDvRE+nBxRE2MNorm)
      nRe3MRealiz := (nBxRE3MComp+nBx3MDvRE+nBxRE3MNorm)
      
      @nLin,000 PSAY "|TITULOS RECEBIDOS (M�s)            |"
      @nLin,051 PSAY nRe1MRealiz Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nRe2MRealiz Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nRe3MRealiz Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|COMPENSA��O                        |"
      @nLin,051 PSAY nBxRE1MComp Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nBxRE2MComp Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nBxRE3MComp Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|DEVOLU��O                          |"
      @nLin,051 PSAY nBx1MDvRE Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nBx1MDvRE Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nBx1MDvRE Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|NORMAL                             |"
      @nLin,051 PSAY nBxRE1MNorm Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nBxRE2MNorm Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nBxRE3MNorm Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|TITULOS EM ABERTO (M�s)            |"
      @nLin,051 PSAY nSld1MRec Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nSld2MRec Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nSld3MRec Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      
      nPer1MInadpl := (nSld1MRec/nRe1MRealiz) * 100
      nPer2MInadpl := (nSld2MRec/nRe2MRealiz) * 100
      nPer3MInadpl := (nSld3MRec/nRe3MRealiz) * 100
      
      @nLin,000 PSAY "|% INADIPL�NCIA (M�s)               |"
      @nLin,051 PSAY nPer1MInadpl Picture "@E 99,999,999.999" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nPer2MInadpl Picture "@E 99,999,999.999" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nPer3MInadpl Picture "@E 99,999,999.999" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 99,999,999.999" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 99,999,999.999" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 99,999,999.999" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      
      nTt1MVc := (nAAtTt1MVc+nA1ATt1MVc+nA2ATt1MVc+nAAntTt1MVc)
      nTt2MVc := (nAAtTt2MVc+nA1ATt2MVc+nA2ATt2MVc+nAAntTt2MVc)
      nTt3MVc := (nAAtTt3MVc+nA1ATt3MVc+nA2ATt3MVc+nAAntTt3MVc) 
      
      @nLin,000 PSAY "|TITULOS VENCIDOS (Total)           |"
      @nLin,051 PSAY nTt1MVc Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nTt2MVc Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nTt3MVc Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * "+cAnoAtu+"                           |"
      @nLin,051 PSAY nAAtTt1MVc Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nAAtTt2MVc Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nAAtTt3MVc Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * "+cAno1Menos+"                           |"
      @nLin,051 PSAY nA1ATt1MVc Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nA1ATt2MVc Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nA1ATt3MVc Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * "+cAno2Menos+"                           |"
      @nLin,051 PSAY nA2ATt1MVc Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nA2ATt2MVc Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nA2ATt3MVc Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * Anteriores                    |"
      @nLin,051 PSAY nAAntTt1MVc Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nAAntTt2MVc Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nAAntTt3MVc Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      
      //For�a Quebra de P�gina
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 8
      
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|ADMINISTRATIVO - FINANCEIRO (Continua��o)"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MERCADO INTERNO                     "
      @nLin,134 PSAY "|MERCADO EXTERNO                     "
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MES                                |"
      @nLin,037 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,066 PSAY "|"
      @nLin,067 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,096 PSAY "|"
      @nLin,097 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,135 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,164 PSAY "|"
      @nLin,165 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,195 PSAY "|"
      @nLin,196 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //Titulos A Receber Fluxo
      
      /* Five Solutions Consultoria
         20 de agosto de 2008

      nTtFlx1MARE := (nARe1M1a30+nARe1M31a60+nARe1M61a90+nARe1M90M)
      nTtFlx2MARE := (nARe2M1a30+nARe2M31a60+nARe2M61a90+nARe2M90M)
      nTtFlx3MARE := (nARe3M1a30+nARe3M31a60+nARe3M61a90+nARe3M90M)
      */

      nTtFlx1MARE := (nARe1M1a30+nARe1M31a60+nARAc160dd)
      nTtFlx2MARE := (nARe2M1a30+nARe2M31a60+nARAc260dd)
      nTtFlx3MARE := (nARe3M1a30+nARe3M31a60+nARAc360dd)
            
      @nLin,000 PSAY "|TITULOS A RECEBER (FLUXO)          |"
      @nLin,051 PSAY nTtFlx1MARE Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nTtFlx2MARE Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nTtFlx3MARE Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * 01 a 30 dias                   |"
      @nLin,051 PSAY nARe1M1a30 Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nARe2M1a30 Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nARe3M1a30 Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * 31 a 60 dias                   |"
      @nLin,051 PSAY nARe1M31a60 Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nARe2M31a60 Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nARe3M31a60 Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * acima de 60 dias               |"
      @nLin,051 PSAY nARAc160dd  Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nARAc260dd Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nARAc360dd Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
/* AkiIta

      @nLin,000 PSAY "|  * 61 a 90 dias                   |"
      @nLin,051 PSAY nARe1M61a90 Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nARe2M61a90 Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nARe3M61a90 Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * acima de 90 dias               |"
      @nLin,051 PSAY nARe1M90M Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nARe2M90M Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nARe3M90M Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
*/
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif

      nImp1MCTt := (nV1MIRImp+nV1MICMS+nV1MIPI+nPIS1MCOF+n1MTribut+nV1MISS+nV1MIPTU+nV1MCPMF)
      nImp2MCTt := (nV2MIRImp+nV2MICMS+nV2MIPI+nPIS2MCOF+n2MTribut+nV2MISS+nV2MIPTU+nV2MCPMF)
      nImp3MCTt := (nV2MIRImp+nV3MICMS+nV3MIPI+nPIS3MCOF+n3MTribut+nV3MISS+nV3MIPTU+nV3MCPMF)

      nVGG1MMOADM := (nVGG1MDFADM+nVGG1MDPAD+nVRes1MAdm+nVPrm1MAdm) 
      nVGG2MMOADM := (nVGG2MDFADM+nVGG2MDPAD+nVRes2MAdm+nVPrm2MAdm)
      nVGG3MMOADM := (nVGG3MDFADM+nVGG3MDPAD+nVRes3MAdm+nVPrm3MAdm)

      
      nVl1MAdmGG := (nVGG1MMOADM+nVDVia1MAdm+nVSvT1MAdm+nHonor1MAdv+nDsp1MJudc+nRes1MAcAd+nEmp1MAdmEE+nFnd1MFixo+nImp1MCTt+nVOut1MAdm+nOut1MAdm)
      nVl2MAdmGG := (nVGG2MMOADM+nVDVia2MAdm+nVSvT2MAdm+nHonor2MAdv+nDsp2MJudc+nRes2MAcAd+nEmp2MAdmEE+nFnd2MFixo+nImp2MCTt+nVOut2MAdm+nOut2MAdm)
      nVl3MAdmGG := (nVGG3MMOADM+nVDVia3MAdm+nVSvT3MAdm+nHonor3MAdv+nDsp3MJudc+nRes3MAcAd+nEmp3MAdmEE+nFnd3MFixo+nImp3MCTt+nVOut3MAdm+nOut3MAdm)
      
      @nLin,000 PSAY "|GASTOS GERAIS ADM.                 |"
      @nLin,051 PSAY nVl1MAdmGG Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nVl2MAdmGG Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nVl3MAdmGG Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      /*
      nVGG1MMOADM := (nVGG1MDFADM+nVGG1MDPAD+nVRes1MAdm+nVPrm1MAdm) 
      nVGG2MMOADM := (nVGG2MDFADM+nVGG2MDPAD+nVRes2MAdm+nVPrm2MAdm)
      nVGG3MMOADM := (nVGG3MDFADM+nVGG3MDPAD+nVRes3MAdm+nVPrm3MAdm)
      */
      
      @nLin,000 PSAY "|M�O DE OBRA                        |"
      @nLin,051 PSAY nVGG1MMOADM Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nVGG2MMOADM Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nVGG3MMOADM Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //Five Solutions, 11/07/2008 - Tratamento para imprimir indicadores apenas se os mesmos tiverem valor.
      //Solicita��o: Gilmar Filho / Patr�cia.      
      
      If (nVGG1MDFADM+nVGG2MDFADM+nVGG3MDFADM) > 0
      
         @nLin,000 PSAY "|  * Desp. Fixa                     |"
         @nLin,051 PSAY nVGG1MDFADM Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVGG2MDFADM Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVGG3MDFADM Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nVGG1MDPAD+nVGG2MDPAD+nVGG3MDPAD) > 0
      
         @nLin,000 PSAY "|  * Desp. Period.                  |"
         @nLin,051 PSAY nVGG1MDPAD Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVGG2MDPAD Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVGG3MDPAD Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                                          
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nVRes1MAdm+nVRes2MAdm+nVRes3MAdm) > 0
      
         @nLin,000 PSAY "|  * Rescis�o                       |"
         @nLin,051 PSAY nVRes1MAdm Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVRes2MAdm Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVRes3MAdm Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                                          
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nVPrm1MAdm+nVPrm2MAdm+nVPrm3MAdm) > 0
      
         @nLin,000 PSAY "|  * Premia��es                     |"
         @nLin,051 PSAY nVPrm1MAdm Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVPrm2MAdm Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVPrm3MAdm Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"                                          
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nVDVia1MAdm+nVDVia2MAdm+nVDVia3MAdm) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "|  DESPESAS DE VIAGENS              |"
         @nLin,051 PSAY nVDVia1MAdm Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVDVia2MAdm Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVDVia3MAdm Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         
      EndIf
      
      If (nVSvT1MAdm+nVSvT2MAdm+nVSvT3MAdm) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "|  SERVI�OS TERCEIRIZADOS           |"
         @nLin,051 PSAY nVSvT1MAdm Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVSvT2MAdm Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVSvT3MAdm Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif

      EndIf
      
      If (nHonor1MAdv+nHonor2MAdv+nHonor3MAdv) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "|  HONOR�RIOS ADVOCAT�CIOS          |"
         @nLin,051 PSAY nHonor1MAdv Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nHonor2MAdv Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nHonor3MAdv Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nDsp1MJudc+nDsp3MJudc+nDsp3MJudc) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "|  DESPESAS JUDICIAIS               |"
         @nLin,051 PSAY nDsp1MJudc Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nDsp2MJudc Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nDsp3MJudc Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nRes1MAcAd+nRes2MAcAd+nRes3MAcAd) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "|  RESCIS�ES ACORDOS/ADV            |"
         @nLin,051 PSAY nRes1MAcAd Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nRes2MAcAd Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nRes3MAcAd Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nEmp1MAdmEE+nEmp2MAdmEE+nEmp3MAdmEE) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "|EMPR�STIMOS ENTRE EMPRESAS         |"
         @nLin,051 PSAY nEmp1MAdmEE Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nEmp2MAdmEE Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nEmp3MAdmEE Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif      
      EndIf
      
      If (nFnd1MFixo+nFnd2MFixo+nFnd3MFixo) > 0
      
         @nLin,000 PSAY "|  FUNDO FIXO                       |"
         @nLin,051 PSAY nFnd1MFixo Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nFnd2MFixo Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nFnd3MFixo Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIF
      
      //If (nImp1MCTt+nImp2MCTt+nImp3MCTt) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"      
         //Impxxx
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         /*
         nImp1MCTt := (nV1MIRImp+nV1MICMS+nV1MIPI+nPIS1MCOF+n1MTribut+nV1MISS+nV1MIPTU+nV1MCPMF)
         nImp2MCTt := (nV2MIRImp+nV2MICMS+nV2MIPI+nPIS2MCOF+n2MTribut+nV2MISS+nV2MIPTU+nV2MCPMF)
         nImp3MCTt := (nV2MIRImp+nV3MICMS+nV3MIPI+nPIS3MCOF+n3MTribut+nV3MISS+nV3MIPTU+nV3MCPMF)
         */
      
         @nLin,000 PSAY "|IMPOSTOS                           |"
         @nLin,051 PSAY nImp1MCTt Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nImp2MCTt Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nImp3MCTt Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      //EndIf
      
      If (nV1MIRImp+nV2MIRImp+nV3MIRImp) > 0
      
         @nLin,000 PSAY "|  * IR                             |"
         @nLin,051 PSAY nV1MIRImp Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nV2MIRImp Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nV3MIRImp Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nV1MICMS+nV2MICMS+nV3MICMS) > 0
      
         @nLin,000 PSAY "|  * ICMS                           |"
         @nLin,051 PSAY nV1MICMS Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nV2MICMS Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nV3MICMS Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"            
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nV1MIPI+nV2MIPI+nV3MIPI) > 0
      
         @nLin,000 PSAY "|  * IPI                            |"
         @nLin,051 PSAY nV1MIPI Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nV2MIPI Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nV3MIPI Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nPIS1MCOF+nPIS2MCOF+nPIS3MCOF) > 0
      
         @nLin,000 PSAY "|  * PIS/COFINS                     |"
         @nLin,051 PSAY nPIS1MCOF Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nPIS2MCOF Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nPIS3MCOF Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (n1MTribut+n2MTribut+n3MTribut) > 0
      
         @nLin,000 PSAY "|  * TRIBUTOS                       |"
         @nLin,051 PSAY n1MTribut Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY n2MTribut Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY n3MTribut Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nV1MISS+nV2MISS+nV3MISS) > 0
      
         @nLin,000 PSAY "|  * ISS                            |"
         @nLin,051 PSAY nV1MISS Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nV2MISS Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nV3MISS Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nV1MIPTU+nV2MIPTU+nV3MIPTU) > 0
      
         @nLin,000 PSAY "|  * IPTU                           |"
         @nLin,051 PSAY nV1MIPTU Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nV2MIPTU Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nV3MIPTU Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nV1MCPMF+nV2MCPMF+nV3MCPMF) > 0
      
         @nLin,000 PSAY "|  * CPMF                           |"
         @nLin,051 PSAY nV1MCPMF Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nV2MCPMF Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nV3MCPMF Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nVOut1MAdm+nVOut2MAdm+nVOut3MAdm) > 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "|  OUTRAS DESPESAS ADM.             |"
         @nLin,051 PSAY nVOut1MAdm Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nVOut2MAdm Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nVOut3MAdm Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      EndIf
      
      If (nOut1MAdm+nOut2MAdm+nOut3MAdm)> 0
      
         @nLin,000 PSAY "|"
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         @nLin,000 PSAY "|  OUTRAS NATUREZAS                 |"
         @nLin,051 PSAY nOut1MAdm Picture "@E 999,999,999.99" 
         @nLin,066 PSAY "|"
         @nLin,081 PSAY nOut2MAdm Picture "@E 999,999,999.99" 
         @nLin,096 PSAY "|"
         @nLin,112 PSAY nOut3MAdm Picture "@E 999,999,999.99" 
         @nLin,127 PSAY "|"
         @nLin,134 PSAY "|"
         @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,164 PSAY "|"
         @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,195 PSAY "|"
         @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
         @nLin,219 PSAY "|"      
         nLin ++
         If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      
      EndIf
      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|VALOR META OR�AMENTO ADM           |"
      @nLin,051 PSAY nAdm1MMtOrc Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nAdm2MMtOrc Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nAdm3MMtOrc Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|GASTOS METAS OR�AMENTO ADM         |"
      @nLin,051 PSAY nGMt1MAdm Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nGMt2MAdm Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nGMt3MAdm Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      
      /* Retirado P�gina conforme solicita��o de Patr�cia - 15/02/2008
      
      //For�a Quebra de P�gina
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 8
      
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|ADMINISTRATIVO - FINANCEIRO (Continua��o)"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MERCADO INTERNO                     "
      @nLin,134 PSAY "|MERCADO EXTERNO                     "
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //                       10        20        30        40
      //              01234567890123456789012345678901234567890
      @nLin,000 PSAY "|MES                                |"
      @nLin,037 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,066 PSAY "|"
      @nLin,067 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,096 PSAY "|"
      @nLin,097 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,135 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,164 PSAY "|"
      @nLin,165 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,195 PSAY "|"
      @nLin,196 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|IMPOSTOS CONTRIB. (Total)          |"
      @nLin,051 PSAY nImp1MCTt Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nImp2MCTt Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nImp3MCTt Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      
      
      //nLin ++
      //If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
      //   Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      //   nLin := 8
      //Endif
      //@nLin,000 PSAY "|  * INSS                           |"
      //@nLin,051 PSAY nV1MINSS Picture "@E 999,999,999.99" 
      //@nLin,066 PSAY "|"
      //@nLin,081 PSAY nV2MINSS Picture "@E 999,999,999.99" 
      //@nLin,096 PSAY "|"
      //@nLin,112 PSAY nV3MINSS Picture "@E 999,999,999.99" 
      //@nLin,127 PSAY "|"
      //@nLin,134 PSAY "|"
      //@nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      //@nLin,164 PSAY "|"
      //@nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      //@nLin,195 PSAY "|"
      //@nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      //@nLin,219 PSAY "|"      
      //nLin ++
      //If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
      //   Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      //   nLin := 8
      //Endif
      //@nLin,000 PSAY "|  * FGTS                           |"
      //@nLin,051 PSAY nV1MFGTS Picture "@E 999,999,999.99" 
      //@nLin,066 PSAY "|"
      //@nLin,081 PSAY nV2MFGTS Picture "@E 999,999,999.99" 
      //@nLin,096 PSAY "|"
      //@nLin,112 PSAY nV3MFGTS Picture "@E 999,999,999.99" 
      //@nLin,127 PSAY "|"
      //@nLin,134 PSAY "|"
      //@nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      //@nLin,164 PSAY "|"
      //@nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      //@nLin,195 PSAY "|"
      //@nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      //@nLin,219 PSAY "|"      
      
      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * CPMF                           |"
      @nLin,051 PSAY nV1MCPMF Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nV2MCPMF Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nV3MCPMF Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * IPTU                           |"
      @nLin,051 PSAY nV1MIPTU Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nV2MIPTU Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nV3MIPTU Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * IPI                            |"
      @nLin,051 PSAY nV1MIPI Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nV2MIPI Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nV3MIPI Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * ICMS                           |"
      @nLin,051 PSAY nV1MICMS Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nV2MICMS Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nV3MICMS Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * ISS                            |"
      @nLin,051 PSAY nV1MISS Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nV2MISS Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nV3MISS Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * IR                             |"
      @nLin,051 PSAY nV1MIRImp Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nV2MIRImp Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nV3MIRImp Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      
*******************************************************************************************************************
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * PIS/COFINS                     |"
      @nLin,051 PSAY nPIS1MCOF Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nPIS2MCOF Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nPIS3MCOF Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * TRIBUTOS                       |"
      @nLin,051 PSAY n1MTribut Picture "@E 999,999,999.99" 
      @nLin,066 PSAY "|"
      @nLin,081 PSAY n2MTribut Picture "@E 999,999,999.99" 
      @nLin,096 PSAY "|"
      @nLin,112 PSAY n3MTribut Picture "@E 999,999,999.99" 
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      

*******************************************************************************************************************
      
      //nLin ++
      //If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
      //   Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      //   nLin := 8
      //Endif
      //@nLin,000 PSAY "|  * TRIBUTOS IMPORTA��O            |"
      //@nLin,051 PSAY nV1MImpImport Picture "@E 999,999,999.99" 
      //@nLin,066 PSAY "|"
      //@nLin,081 PSAY nV2MImpImport Picture "@E 999,999,999.99" 
      //@nLin,096 PSAY "|"
      //@nLin,112 PSAY nV3MImpImport Picture "@E 999,999,999.99" 
      //@nLin,127 PSAY "|"
      //@nLin,134 PSAY "|"
      //@nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      //@nLin,164 PSAY "|"
      //@nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      //@nLin,195 PSAY "|"
      //@nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      //@nLin,219 PSAY "|"      
      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      
      */  // Fim - Retirado P�gina conforme solicita��o de Patr�cia - 15/02/2008
      
      //For�a Quebra de P�gina
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 8
      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|LOGISTICA"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MERCADO INTERNO                     "
      @nLin,134 PSAY "|MERCADO EXTERNO                     "
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|MES                                |"
      @nLin,037 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,066 PSAY "|"
      @nLin,067 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,096 PSAY "|"
      @nLin,097 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,135 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,164 PSAY "|"
      @nLin,165 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,195 PSAY "|"
      @nLin,196 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|FOB - Faturamento                  |" //
      @nLin,051 PSAY  nFOB1MFt  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nFOB2MFt Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nFOB3MFt Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|CIF"
      @nLin,219 PSAY "|" //
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * Faturamento                    |" //
      @nLin,051 PSAY  nFt1MCIF  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nFt2MCIF Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nFt3MCIF Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * Frete                          |" //
      @nLin,051 PSAY  nVl1MsFrete  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nVl2MsFrete Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nVl3MsFrete Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      
      nPer1MCIF := (nVl1MsFrete/nFt1MCIF) * 100
      nPer2MCIF := (nVl2MsFrete/nFt2MCIF) * 100
      nPer3MCIF := (nVl3MsFrete/nFt3MCIF) * 100
      
      @nLin,000 PSAY "|  * Percentual (%) Frete           |" //
      @nLin,051 PSAY  nPer1MCIF  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nPer2MCIF Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nPer3MCIF Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|FATURAMENTO S/ DEF. FRETE          |" //
      @nLin,051 PSAY  nFt1MSDEF  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nFt2MSDEF Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nFt3MSDEF Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      
      n1MTotFrtPg := (n1FrtVeMVl+n1FrtImpMVl+n1FrtExpMVl+n1FrtBenMVl+n1FrtTrfMVl+n1FrtComMVl+n1FrtIndMVl+n1FrtAdmMVl)
      n2MTotFrtPg := (n2FrtVeMVl+n2FrtImpMVl+n2FrtExpMVl+n2FrtBenMVl+n2FrtTrfMVl+n2FrtComMVl+n2FrtIndMVl+n2FrtAdmMVl)
      n3MTotFrtPg := (n3FrtVeMVl+n3FrtImpMVl+n3FrtExpMVl+n3FrtBenMVl+n3FrtTrfMVl+n3FrtComMVl+n3FrtIndMVl+n3FrtAdmMVl)

      @nLin,000 PSAY "|FRETES (Total)                     |" //
      @nLin,051 PSAY  n1MTotFrtPg  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  n2MTotFrtPg Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  n3MTotFrtPg Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * Vendas                         |" //
      @nLin,051 PSAY  n1FrtVeMVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  n2FrtVeMVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  n3FrtVeMVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * Importa��o                     |" //
      @nLin,051 PSAY  n1FrtImpMVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  n2FrtImpMVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  n3FrtImpMVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * Exporta��o                     |" //
      @nLin,051 PSAY  n1FrtExpMVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  n2FrtExpMVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  n3FrtExpMVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                        
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * Beneficiamento                 |" //
      @nLin,051 PSAY  n1FrtBenMVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  n2FrtBenMVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  n3FrtBenMVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * Transfer�ncia                  |" //
      @nLin,051 PSAY  n1FrtTrfMVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  n2FrtTrfMVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  n3FrtTrfMVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * Comercio                       |" //
      @nLin,051 PSAY  n1FrtComMVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  n2FrtComMVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  n3FrtComMVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * Ind�stria                      |" //
      @nLin,051 PSAY  n1FrtIndMVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  n2FrtIndMVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  n3FrtIndMVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|  * Administrativo                 |" //
      @nLin,051 PSAY  n1FrtAdmMVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  n2FrtAdmMVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  n3FrtAdmMVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                  
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"

   EndIf //Fim da Condi��o para impress�o dos dados da An�lise Gerencial Administrativo/Financeira
   
   If lDemonsRes
      
      //For�a Quebra de P�gina
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 8
      
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|DEMONSTRATIVO DE RESULTADO"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //SOMAT�RIO DAS VENDAS DO MI + SUCATA + ME
      @nLin,000 PSAY "|FATURAMENTO BRUTO                  |" //
      @nLin,051 PSAY  (nFat1MesBru + nVl1SUCMFat)  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  (nFat2MesBru + nVl2SUCMFat) Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  (nFat3MesBru + nVl3SUCMFat) Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //SOMAT�RIO DAS VENDAS DO MI + SUCATA
      @nLin,000 PSAY "|  Mercado Interno (Total)          |" //
      @nLin,051 PSAY  (nFat1MesBru + nVl1SUCMFat)  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  (nFat2MesBru + nVl2SUCMFat) Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  (nFat3MesBru + nVl3SUCMFat) Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //TODAS AS VENDAS DO MI QUE POSSUEM A CONDI��O DE FATURAMENTO "ANTECIPADO"
      @nLin,000 PSAY "|    * Vendas a Vista               |" //
      @nLin,051 PSAY  nFat1AVista  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nFat2AVista Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nFat3AVista Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //TODAS AS CONDI��ES DE VENDAS DO MI COM EXCESS�O "ANTECIPADO"
      @nLin,000 PSAY "|    * Vendas a Prazo               |" //
      @nLin,051 PSAY  nFat1APrazo  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nFat2APrazo Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nFat3APrazo Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //TODAS AS CONDI��ES DE VENDA DA SUCATA
      @nLin,000 PSAY "|    * Sucata                       |" //
      @nLin,051 PSAY  nVl1SUCMFat  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nVl2SUCMFat Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nVl3SUCMFat Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //SOMAT�RIO DAS VENDAS DO MI + SUCATA
      @nLin,000 PSAY "|  Mercado Externo (Total)          |" //
      @nLin,051 PSAY  (0)  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  (0) Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  (0) Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //TODAS AS VENDAS DO MI QUE POSSUEM A CONDI��O DE FATURAMENTO "ANTECIPADO"
      @nLin,000 PSAY "|    * Vendas a Vista               |" //
      @nLin,051 PSAY  0  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  (0) Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  (0) Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //TODAS AS CONDI��ES DE VENDAS DO MI COM EXCESS�O "ANTECIPADO"
      @nLin,000 PSAY "|    * Vendas a Prazo               |" //
      @nLin,051 PSAY  0  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  (0) Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  (0) Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //TODAS AS DEDU��ES DE VENDAS (IMPOSTOS,  DEVOLU��ES, OUTRAS DEDU��ES )

      /* Manuten��o efetuada pouco antes da impress�o em S�o Paulo/SP - Solicita��o: Patr�cia
      Zerado os impostos para evitar subtra��o na materia prima.
      
      nImp1SVend := (nIPI1Valor+nICM1Vlor+nPIS1Valor+nCOF1Valor)
      nImp2SVend := (nIPI2Valor+nICM2Vlor+nPIS2Valor+nCOF2Valor)
      nImp3SVend := (nIPI3Valor+nICM3Vlor+nPIS3Valor+nCOF3Valor)
      */

      nIPI1Valor := 0
      nICM1Vlor  := 0
      nPIS1Valor := 0
      nCOF1Valor := 0
      
      nIPI2Valor := 0
      nICM2Vlor  := 0
      nPIS2Valor := 0
      nCOF2Valor := 0
      
      nIPI3Valor := 0
      nICM3Vlor  := 0
      nPIS3Valor := 0
      nCOF3Valor := 0
      
      /*Fim da Manuten��o dos Impostos - Valores Zerados*/
      
      nImp1SVend := (nIPI1Valor+nICM1Vlor+nPIS1Valor+nCOF1Valor)
      nImp2SVend := (nIPI2Valor+nICM2Vlor+nPIS2Valor+nCOF2Valor)
      nImp3SVend := (nIPI3Valor+nICM3Vlor+nPIS3Valor+nCOF3Valor)
      
      nDevol1M  := (nVl1MesDev)
      nDevol2M  := (nVl2MesDev)
      nDevol3M  := (nVl3MesDev)
      
      nDed1MVend := (nImp1SVend+nDevol1M)
      nDed2MVend := (nImp2SVend+nDevol2M)
      nDed3MVend := (nImp3SVend+nDevol3M)

      @nLin,000 PSAY "|(-)DEDU��ES DE VENDAS              |" //
      @nLin,051 PSAY  nDed1MVend  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nDed2MVend Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nDed3MVend Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      /*
      nImp1SVend := (nIPI1Valor+nICM1Vlor+nPIS1Valor+nCOF1Valor)
      nImp2SVend := (nIPI2Valor+nICM2Vlor+nPIS2Valor+nCOF2Valor)
      nImp3SVend := (nIPI3Valor+nICM3Vlor+nPIS3Valor+nCOF3Valor)
      */
      
      //SOMAT�RIO DOS IMPOSTOS QUE INCIDEM NAS N.F.'S DE VENDAS
      @nLin,000 PSAY "|  Impostos                         |" //
      @nLin,051 PSAY  nImp1SVend  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nImp2SVend Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nImp3SVend Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //VALOR DO ICMS LAN�ADO NAS N.F.'S DE VENDAS (LIVRO FISCAL DE SA�DA) 
      @nLin,000 PSAY "|    * ICMS                         |" //
      @nLin,051 PSAY  nICM1Vlor  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nICM2Vlor Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nICM3Vlor Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //VALOR DO IPI LAN�ADO NAS N.F.'S DE VENDAS (LIVRO FISCAL DE SA�DA)
      @nLin,000 PSAY "|    * IPI                          |" //
      @nLin,051 PSAY  nIPI1Valor  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nIPI2Valor Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nIPI3Valor Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //CALCULAR PIS (1,65%) E COFINS (7,60%) SOBRE O FATURAMENTO TOTAL DE VENDAS (LIVRO FISCAL DE SA�DA)
      @nLin,000 PSAY "|    * PIS+COFINS                   |" //
      @nLin,051 PSAY  (nPIS1Valor+nCOF1Valor)  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  (nPIS2Valor+nCOF2Valor) Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  (nPIS3Valor+nCOF3Valor) Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //DEVOLU��ES DO MI E ME REGISTRADAS PELO LIVRO FISCAL DE ENTRADA 
      /*
      nDevol1M  := (nVl1MesDev)
      nDevol2M  := (nVl2MesDev)
      nDevol3M  := (nVl3MesDev)
      */
      
      @nLin,000 PSAY "|  Devolu��es                       |" //
      @nLin,051 PSAY  nDevol1M  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nDevol2M Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nDevol3M Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //DEVOLU��ES DO MI REGISTRADAS PELO LIVRO FISCAL DE ENTRADA 
      @nLin,000 PSAY "|    * Mercado Interno              |" //
      @nLin,051 PSAY  nVl1MesDev  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nVl2MesDev Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nVl3MesDev Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //DEVOLU��ES DO ME REGISTRADAS PELO LIVRO FISCAL DE ENTRADA 
      @nLin,000 PSAY "|    * Mercado Externo              |" //
      @nLin,051 PSAY  (0)  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  (0) Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  (0) Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //OUTRAS DEDU��ES...
      @nLin,000 PSAY "|  Outras Dedu��es                  |" //
      @nLin,051 PSAY  0  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  0 Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  0 Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //RECEITA LIQUIDA = RECEITA BRUTA (-) DEDU��ES DE VENDAS
      NRecLi1Mquida := ((nFat1MesBru + nVl1SUCMFat) - nDed1MVend)
      NRecLi2Mquida := ((nFat2MesBru + nVl2SUCMFat) - nDed2MVend)
      NRecLi3Mquida := ((nFat3MesBru + nVl3SUCMFat) - nDed3MVend)
      @nLin,000 PSAY "|RECEITA L�QUIDA DE VENDAS          |" //
      @nLin,051 PSAY  NRecLi1Mquida  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  NRecLi2Mquida Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  NRecLi3Mquida Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //SOMAT�RIO DOS CUSTOS DIRETOS E INDIRETOS 
      nMPCst1M := (nCMP1Mes+nVlM1CusSUC+nCus1MesDev) - nImp1SVend
      nMPCst2M := (nCMP2Mes+nVlM2CusSUC+nCus2MesDev) - nImp2SVend
      nMPCst3M := (nCMP3Mes+nVlM3CusSUC+nCus3MesDev) - nImp3SVend

      nCust1oDir := (nMPCst1M+nMPMO1IndAdm+nBnf1IndVl+nInsVl1MInd)
      nCust2oDir := (nMPCst2M+nMPMO2IndAdm+nBnf2IndVl+nInsVl2MInd)
      nCust3oDir := (nMPCst3M+nMPMO3IndAdm+nBnf3IndVl+nInsVl3MInd)
      
      nInd1MCusto := (nSalInd1M+nMan1IndVl+nSTc1IndVl+nEPI1IndVl+nEner1IndVl+nAgua1IndVl+nAlg1IndVl+nODp1IndVl)
      nInd2MCusto := (nSalInd2M+nMan2IndVl+nSTc2IndVl+nEPI2IndVl+nEner2IndVl+nAgua2IndVl+nAlg2IndVl+nODp2IndVl)
      nInd3MCusto := (nSalInd3M+nMan3IndVl+nSTc3IndVl+nEPI3IndVl+nEner3IndVl+nAgua3IndVl+nAlg3IndVl+nODp3IndVl)
      
      nCPV1M := (nCust1oDir+nInd1MCusto) 
      nCPV2M := (nCust2oDir+nInd2MCusto) 
      nCPV3M := (nCust3oDir+nInd3MCusto) 
      
      @nLin,000 PSAY "|(-)CUSTO PRODUTOS VENDIDOS         |" //
      @nLin,051 PSAY  nCPV1M  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nCPV2M Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nCPV3M Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif

      //Novo tratamento p/Custo da Materia Prima - Solicita��o de Patr�cia - 18/03/2008(Total de Custo(-)Impostos)
      /*
      nMPCst1M := (nCMP1Mes+nVlM1CusSUC) - nImp1SVend
      nMPCst2M := (nCMP2Mes+nVlM2CusSUC) - nImp2SVend
      nMPCst3M := (nCMP3Mes+nVlM3CusSUC) - nImp3SVend
      */
      
      //TOTAL DO CUSTO DIRETO
      
      /*
      nCust1oDir := (nMP1MVlTt+nMPMO1IndAdm+nBnf1IndVl+nInsVl1MInd)
      nCust2oDir := (nMP2MVlTt+nMPMO2IndAdm+nBnf2IndVl+nInsVl2MInd)
      nCust3oDir := (nMP3MVlTt+nMPMO3IndAdm+nBnf3IndVl+nInsVl3MInd)
      */
      
      /*
      nCust1oDir := (nMPCst1M+nMPMO1IndAdm+nBnf1IndVl+nInsVl1MInd)
      nCust2oDir := (nMPCst2M+nMPMO2IndAdm+nBnf2IndVl+nInsVl2MInd)
      nCust3oDir := (nMPCst3M+nMPMO3IndAdm+nBnf3IndVl+nInsVl3MInd)
      */
      
      @nLin,000 PSAY "|  Custo Direto                     |" //
      @nLin,051 PSAY  nCust1oDir  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nCust2oDir Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nCust3oDir Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //CUSTO LAN�ADO NO SISTEMA "NELSON"
      //Novo tratamento p/Custo da Materia Prima - Solicita��o de Patr�cia - 18/03/2008(Total de Custo(-)Impostos)
      /*
      nMPCst1M := (nCMP1Mes+nVlM1CusSUC) - nImp1SVend
      nMPCst2M := (nCMP2Mes+nVlM2CusSUC) - nImp2SVend
      nMPCst3M := (nCMP3Mes+nVlM3CusSUC) - nImp3SVend
      */
      
      @nLin,000 PSAY "|   * Materia Prima                 |" //
      //@nLin,051 PSAY  nMP1MVlTt  Picture "@E 999,999,999.99"
      @nLin,051 PSAY  nMPCst1M  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      //@nLin,081 PSAY  nMP2MVlTt Picture "@E 999,999,999.99"
      @nLin,081 PSAY  nMPCst2M Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      //@nLin,112 PSAY  nMP3MVlTt Picture "@E 999,999,999.99"
      @nLin,112 PSAY  nMPCst3M Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //SALARIOS PRODU��O: TODAS AS DESPESAS DA M.O. DA IND�STRIA INCLUINDO A SUA AREA ADMINISTRATIVA '100301','100302','100303','100304','100305','100306','100307',
      //'100309','100310','100311','100312','100313','100317','100401','100402','100403','100404','100405','100406','100407','100409','100410','100411','100412','100417'
      @nLin,000 PSAY "|   * Salarios Produ��o             |" //
      @nLin,051 PSAY  nMPMO1IndAdm  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nMPMO2IndAdm Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nMPMO3IndAdm Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //BENEFICIAMENTO '200314','600601'
      @nLin,000 PSAY "|   * Beneficiamento                |" //
      @nLin,051 PSAY  nBnf1IndVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nBnf2IndVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nBnf3IndVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //INSUMOS '200308','200408','200312','200412'
      @nLin,000 PSAY "|   * Insumos                       |" //
      @nLin,051 PSAY  nInsVl1MInd  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nInsVl2MInd Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nInsVl3MInd Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //TOTAL DO CUSTO INDIRETO
      /*
      nInd1MCusto := (nSalInd1M+nMan1IndVl+nSTc1IndVl+nEPI1IndVl+nEner1IndVl+nAgua1IndVl+nAlg1IndVl+nODp1IndVl)
      nInd2MCusto := (nSalInd2M+nMan2IndVl+nSTc2IndVl+nEPI2IndVl+nEner2IndVl+nAgua2IndVl+nAlg2IndVl+nODp2IndVl)
      nInd3MCusto := (nSalInd3M+nMan3IndVl+nSTc3IndVl+nEPI3IndVl+nEner3IndVl+nAgua3IndVl+nAlg3IndVl+nODp3IndVl)
      */
      
      @nLin,000 PSAY "|  Custo Indireto                   |" //
      @nLin,051 PSAY nInd1MCusto  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nInd2MCusto Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nInd3MCusto Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //Salarios Adm Industria - COMO N�O TINHA COMO SEPARAR, J� FOI LAN�ADO JUNTAMENTE COM O CUSTO DO SAL�RIO DA PRODU��O
      @nLin,000 PSAY "|   * Salarios Adm Industria        |" //
      //Conforme Solicitacao de Patricia, o valor foi zerado momentaneamente.
      nSalInd1M := 0
      nSalInd2M := 0
      nSalInd3M := 0
      
      //@nLin,051 PSAY  nMPMO1IndAdm  Picture "@E 999,999,999.99"
      @nLin,051 PSAY  nSalInd1M  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      //@nLin,081 PSAY  nMPMO2IndAdm Picture "@E 999,999,999.99"
      @nLin,081 PSAY  nSalInd2M Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      //@nLin,112 PSAY  nMPMO3IndAdm Picture "@E 999,999,999.99"
      @nLin,112 PSAY  nSalInd3M Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //MANUTEN��O '200317','200318','200319','200417','200418','200419'
      @nLin,000 PSAY "|   * Manuten��o                    |" //
      @nLin,051 PSAY  nMan1IndVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nMan2IndVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nMan3IndVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //SERVI�OS TERCEIRIZADOS '200303','200403'
      @nLin,000 PSAY "|   * Servi�os Terceirizados        |" //
      @nLin,051 PSAY  nSTc1IndVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nSTc2IndVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nSTc3IndVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //EPI '200320','200420'
      @nLin,000 PSAY "|   * EPI                           |" //
      @nLin,051 PSAY  nEPI1IndVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nEPI2IndVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nEPI3IndVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //ENERGIA '200311','200411'
      @nLin,000 PSAY "|   * Energia El�trica              |" //
      @nLin,051 PSAY  nEner1IndVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nEner2IndVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nEner3IndVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //AGUA '200310','200410'
      @nLin,000 PSAY "|   * �gua                          |" //
      @nLin,051 PSAY  nAgua1IndVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nAgua2IndVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nAgua3IndVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"                                                      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //ALUGUEL '200309'
      @nLin,000 PSAY "|   * Alugu�l                       |" //
      @nLin,051 PSAY  nAlg1IndVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nAlg2IndVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nAlg3IndVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //AINDA NAO POSSUIMOS ESSE CONTROLE NO SISTEMA
      @nLin,000 PSAY "|   * Deprecia��o                   |" //
      @nLin,051 PSAY  0  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY 0 Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY 0 Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //OUTRAS DESPESAS DAS IND�STRIA
      //'200321','200421','200304','200404','200316','200416','200301','200302','200305','200307','200315','200401','200402','200405','200407','600301','600501','700301','700304'
      @nLin,000 PSAY "|   * Outras Despesas Ind�stria     |" //
      @nLin,051 PSAY  nODp1IndVl  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nODp2IndVl Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nODp3IndVl Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //LUCRO BRUTO = RECEITA LIQUIDA (-) MENOS CUSTRO PRODUTO VENDIDO
      nLcr1MBrut := (NRecLi1Mquida - nCPV1M)
      nLcr2MBrut := (NRecLi2Mquida - nCPV2M)
      nLcr3MBrut := (NRecLi3Mquida - nCPV3M)
      
      @nLin,000 PSAY "|LUCRO BRUTO                        |" //
      @nLin,051 PSAY nLcr1MBrut  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nLcr2MBrut Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nLcr3MBrut Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //DESPESAS FINANCEIRAS = DESPESAS (-) MENOS RECEITAS
      nFin1anc := (nDspFi1nsDER-nRecDR1EFin)
      nFin2anc := (nDspFi2nsDER-nRecDR2EFin)
      nFin3anc := (nDspFi3nsDER-nRecDR3EFin)


      //(-)DESPESAS(Receitas) OPERACIONAIS: DESPESAS/RECEITAS = (COMERCIAIS) + (ADMINISTRATIVAS) + (FINANCEIRAS) + (OUTRAS) 
      nOper1MDesp := (nDesp1ComDRE+nDsp1ADMDER+nFin1anc+nOutras)
      nOper2MDesp := (nDesp2ComDRE+nDsp2ADMDER+nFin2anc+nOutras)
      nOper3MDesp := (nDesp3ComDRE+nDsp3ADMDER+nFin3anc+nOutras)
      
      @nLin,000 PSAY "|(-)DESPESAS(Receitas) OPERACIONAIS |" //
      @nLin,051 PSAY nOper1MDesp  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nOper2MDesp Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nOper3MDesp Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //TODAS AS DESPESAS DO COMERCIAL (EXCLUI A NATUREZA DE EMPR�STIMO E REEMBOLSO A CLIENTE) '100201','100202','100203','100204','100205','100206',
      //'100207','100208','100209','100210','100211','100213','100214','100215','100216','100222','200201','200202','200203','200204','200205','200206','200207','600101','600102','600303','600401'
      @nLin,000 PSAY "|  Comerciais                       |" //
      @nLin,051 PSAY  nDesp1ComDRE  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nDesp2ComDRE Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nDesp3ComDRE Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //TOTAL DE DESPESAS ADMINISTRATIVAS (EXCLUI EMPR�STIMOS, IMPOSTOS(Fica o INSS da Diretoria, FGTS e GRFC) E OUTRAS...)
      //'100101','100102','100103','100104','100105','100106','100107','100108','100109','100110','100111','100112','100114','100116','100117','100118','100119','100122','200101','200102','200103','200104',
      //'200105','200106','200107','200108','200112','200113','200114','200115','200116','200117','200118','200119','200120','200121','200122','200123','200124','200127','200129','200130','200131',
      //'200133','200134','200136','200501','200502','200503','200504','200505','200506','200507','200508','200509','200510','500101','500102','500103','500104','600302',
      @nLin,000 PSAY "|  Administrativas                  |" //
      @nLin,051 PSAY  nDsp1ADMDER  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nDsp2ADMDER Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nDsp3ADMDER Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            

      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif            
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"      
      
      //For�a Quebra de P�gina
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 8
      
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif                  
      //DESPESAS FINANCEIRAS = DESPESAS (-) MENOS RECEITAS
      /*
      nFin1anc := (nDspFi1nsDER-nRecDR1EFin)
      nFin2anc := (nDspFi2nsDER-nRecDR2EFin)
      nFin3anc := (nDspFi3nsDER-nRecDR3EFin)
      */
      
      @nLin,000 PSAY "|  Financeiras                      |" //
      @nLin,051 PSAY  nFin1anc  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY nFin2anc Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY nFin3anc Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //DESPESAS FINANCEIRAS '200141','200126','D.ESCONT'
      @nLin,000 PSAY "|    *Despesas                      |" //
      @nLin,051 PSAY  nDspFi1nsDER  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nDspFi2nsDER Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nDspFi3nsDER Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //RECEITAS FINANCEIRAS '200140'
      @nLin,000 PSAY "|    *Receitas                      |" //
      @nLin,051 PSAY  nRecDR1EFin  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nRecDR2EFin Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nRecDR3EFin Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      
      //OUTRAS DESPESAS/RECEITAS QUE NAO DESCRIMINADAS ACIMA
      nOutras := (0)
      
      @nLin,000 PSAY "|  Outras                           |" //
      @nLin,051 PSAY  0  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  0 Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  0 Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //OUTRAS DESPESAS
      @nLin,000 PSAY "|    *Despesas                      |" //
      @nLin,051 PSAY  0  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  0 Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  0 Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //OUTRAS RECEITAS
      @nLin,000 PSAY "|    *Receitas                      |" //
      @nLin,051 PSAY  0  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  0 Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  0 Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //LUCRO OPERACIONAL = LUCRO BRUTO (-) MENOS DESPESAS/RECEITAS OPERACIONAIS
      nLcr1MOper := (nLcr1MBrut - nOper1MDesp)
      nLcr2MOper := (nLcr2MBrut - nOper2MDesp)
      nLcr3MOper := (nLcr3MBrut - nOper3MDesp)
      
      @nLin,000 PSAY "|LUCRO OPERACIONAL                  |" //
      @nLin,051 PSAY  nLcr1MOper  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nLcr2MOper Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nLcr3MOper Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //NAO EXISTE NENHUM TIPO DE CONTROLE NO SISTEMA
      //(TUDO QUE � VENDIDO QUE NAO ESTA LIGADO A ATIVIDADE PRINCIPAL DA EMPRESA - EX.: VENDA DE IMOBILIZADO, A��ES�)
      nRes1MNOLiq := (0)
      nRes2MNOLiq := (0)
      nRes3MNOLiq := (0)
      
      @nLin,000 PSAY "|Resultado N�o Operacional L�quido  |" //
      @nLin,051 PSAY  nRes1MNOLiq  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nRes2MNOLiq Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nRes3MNOLiq Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //LUCRO ANTES DO IR E CSLL = LUCRO OPERACIONAL (+/-) RESULTADO NAO OPERACIONAL
      nL1MAIRCSL := (nLcr1MOper-nRes1MNOLiq)
      nL2MAIRCSL := (nLcr2MOper-nRes2MNOLiq)
      nL3MAIRCSL := (nLcr3MOper-nRes3MNOLiq)
      
      @nLin,000 PSAY "|LUCRO ANTES DO IR E CSLL           |" //
      @nLin,051 PSAY  nL1MAIRCSL  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nL2MAIRCSL Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nL3MAIRCSL Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //PROVISAO PARA CSLL = QDO. O "LUCRO ANTES DO IR E CSLL" DER POSITIVO CALCULAR 9% (NOVE POR CENTO)
      nPrV1MCSLL := 0
      nPrV2MCSLL := 0
      nPrV3MCSLL := 0
      
      If nL1MAIRCSL > 0
         nPrV1MCSLL := (nL1MAIRCSL * (0.09))
      EndIf

      If nL2MAIRCSL > 0
         nPrV2MCSLL := (nL2MAIRCSL * (0.09))
      EndIf

      If nL3MAIRCSL > 0
         nPrV3MCSLL := (nL3MAIRCSL * (0.09))
      EndIf
      
      @nLin,000 PSAY "|PROVIS�O PARA O CSLL               |" //
      @nLin,051 PSAY  nPrV1MCSLL  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nPrV2MCSLL Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nPrV3MCSLL Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //PROVISAO PARA IR = "LUCRO ANTES DO IR E CSLL"  (-) MENOS "PROVISAO PARA CSLL" QUANDO DER POSITIVO CALCULAR 12% (DOZE POR CENTO)
      
      nPrvIR1M := 0
      nPrvIR2M := 0
      nPrvIR3M := 0
      
      If nL1MAIRCSL > 0
         nPrvIR1M := (nL1MAIRCSL * (0.15)) + ((nL1MAIRCSL - (20000))*(0.10))
      EndIf
     
      If nL2MAIRCSL > 0
         nPrvIR2M := (nL2MAIRCSL * (0.15)) + ((nL2MAIRCSL - (20000))*(0.10))
      EndIf      

      If nL3MAIRCSL > 0
         nPrvIR3M := (nL3MAIRCSL * (0.15)) + ((nL3MAIRCSL - (20000))*(0.10))
      EndIf
            
      @nLin,000 PSAY "|PROVIS�O PARA O IR                 |" //
      @nLin,051 PSAY  nPrvIR1M  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nPrvIR2M Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nPrvIR3M Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //LUCRO LIQUIDO DO EXERCICIO = "LUCRO ANTES DO IR E CSLL" (-) MENOS "PROVISAO PARA CSLL" E (-) MENOS "PROVISAO PARA IR"
      
      nLL1MExerc := nL1MAIRCSL - (nPrV1MCSLL+nPrvIR1M)
      nLL2MExerc := nL2MAIRCSL - (nPrV2MCSLL+nPrvIR2M)
      nLL3MExerc := nL3MAIRCSL - (nPrV3MCSLL+nPrvIR3M)
      
      @nLin,000 PSAY "|LUCRO L�QUIDO DO EXERC�CIO         |" //
      @nLin,051 PSAY  nLL1MExerc  Picture "@E 999,999,999.99"
      @nLin,066 PSAY "|"
      @nLin,081 PSAY  nLL2MExerc Picture "@E 999,999,999.99"
      @nLin,096 PSAY "|"
      @nLin,112 PSAY  nLL3MExerc Picture "@E 999,999,999.99"
      @nLin,127 PSAY "|"
      @nLin,134 PSAY "|"
      @nLin,149 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,164 PSAY "|"
      @nLin,180 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,195 PSAY "|"
      @nLin,203 PSAY  (0) Picture "@E 999,999,999.99" 
      @nLin,219 PSAY "|"            
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|"
      @nLin,219 PSAY "|"      
      nLin ++
      If nLin > nLRel //55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      //xxaki
            
   EndIf
   
   nLin ++ // Avanca a linha de impressao

   //dbSkip() // Avanca o ponteiro do registro no arquivo
   lRoda := .F. //For�a sa�da do La�o de Impress�o.
   
EndDo

//���������������������������������������������������������������������Ŀ
//� Finaliza a execucao do relatorio...                                 �
//�����������������������������������������������������������������������

SET DEVICE TO SCREEN

//���������������������������������������������������������������������Ŀ
//� Se impressao em disco, chama o gerenciador de impressao...          �
//�����������������������������������������������������������������������

If aReturn[5]==1
   dbCommitAll()
   SET PRINTER TO
   OurSpool(wnrel)
Endif

MS_FLUSH()
Return


//�������������������������������������������������������������������������������������������������������������Ŀ
//�                                     F U N � � E S   G E N � R I C A S                                                        �
//���������������������������������������������������������������������������������������������������������������


Static Function AjustaSx1

Local aArea := GetArea()

//PutSx1(cGrupo,cOrdem,cPergunt  ,"","",cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
  PutSx1(cPerg ,"01","Data de ?" ,"","","mv_ch1","D"  ,8       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR01",""     ,"","",""    ,""           ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"02","Data ate ?","","","mv_ch2","D"  ,8       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR02",""     ,"","",""    ,""           ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"03","Custo ?"   ,"","","mv_ch3","C"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR03","Medio","","",""    ,"Reposicao"  ,"","","Calculado","","",""          ,"","","","","")
  PutSx1(cPerg ,"04","Analitico?","","","mv_ch4","C"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR04","Sim"  ,"","",""    ,"Nao"        ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"05","Depto.?"   ,"","","mv_ch5","C"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria","","","Financeiro","","","Gerencial","","")

RestArea(aArea)

Return(.T.)

Static Function fNomeMes(cMes)
   cNomeMes := ""
   If cMes == "01"
      cNomeMes := "JANEIRO"
   ElseIf cMes == "02"
      cNomeMes := "FEVEREIRO"
   ElseIf cMes == "03"
      cNomeMes := "MAR�O"
   ElseIf cMes == "04"
      cNomeMes := "ABRIL"
   ElseIf cMes == "05"
      cNomeMes := "MAIO"
   ElseIf cMes == "06"
      cNomeMes := "JUNHO"
   ElseIf cMes == "07"
      cNomeMes := "JULHO"
   ElseIf cMes == "08"
      cNomeMes := "AGOSTO"
   ElseIf cMes == "09"
      cNomeMes := "SETEMBBRO"
   ElseIf cMes == "10"
      cNomeMes := "OUTUBRO"
   ElseIf cMes == "11"
      cNomeMes := "NOVEMBRO"
   ElseIf cMes == "12"
      cNomeMes := "DEZEMBRO"
   EndIf
Return(cNomeMes)

Static Function ItaSldPagar(dData,nMoeda,lDtAnterior,lMovSe5)

Local aArea     := { Alias() , IndexOrd() , Recno() }
Local aAreaSE2  := { SE2->(IndexOrd()), SE2->(Recno()) }
Local bCondSE2
Local nSaldo    := 0

#IFDEF TOP
	Local cFiltro   := ""
#ENDIF

//Alert("Estou no ItaSldPagar ")
// ������������������������������������������������������Ŀ
// � Testa os parametros vindos do Excel                  �
// ��������������������������������������������������������
nMoeda      := If(Empty(nMoeda),1,nMoeda)
dData       := If(Empty(dData),dDataBase,dData)
lDtAnterior := BoolWindow(lDtAnterior)
lMovSe5     := BoolWindow(lMovSe5)
If ( ValType(nMoeda) == "C" )
	nMoeda      := Val(nMoeda)
EndIf
dData       := DataWindow(dData)
// Quando eh chamada do Excel, estas variaveis estao em branco
IF Empty(MVABATIM) .Or.;
	Empty(MV_CPNEG) .Or.;
	Empty(MVPAGANT) .Or.;
	Empty(MVPROVIS)
	CriaTipos()
Endif
dbSelectArea("SE2")
dbSetOrder(3)
//Ita - 11/03/2008
cQrySE2 := "SELECT SE2.E2_PREFIXO,SE2.E2_NUM,SE2.E2_PARCELA,SE2.E2_TIPO,SE2.E2_NATUREZ,SE2.E2_FORNECE,SE2.E2_LOJA,"
cQrySE2 += "SE2.E2_SALDO,SE2.E2_MOEDA,SE2.E2_TIPO,SE2.E2_NATUREZ"
cQrySE2 += " FROM "+RetSQLName("SE2")+" SE2"
cQrySE2 += " WHERE "
cQrySE2 += " SE2.E2_EMIS1 <= '"+DTOS(dData)+"' AND " //SE2->E2_EMIS1 <= dData .And. ;
cQrySE2 += " SE2.E2_TIPO NOT IN "+FormatIn(MVPROVIS+MVABATIM,"|")+" AND "//!SE2->E2_TIPO $MVPROVIS+"/"+MVABATIM .AND. ;
cQrySE2 += " ((SUBSTRING(SE2.E2_FATURA,1,3) <> '   ' AND "//((!Empty(SE2->E2_FATURA).And.;
cQrySE2 += " SUBSTRING(SE2.E2_FATURA,1,6) = 'NOTFAT' ) OR "//Substr(SE2->E2_FATURA,1,6)=="NOTFAT" ) .Or.;
cQrySE2 += " (SUBSTRING(SE2.E2_FATURA,1,3) <> '   ' AND "//(!Empty(SE2->E2_FATURA) .And.;
cQrySE2 += " SUBSTRING(SE2.E2_FATURA,1,6) <> 'NOTFAT' AND "//Substr(SE2->E2_FATURA,1,6)!="NOTFAT" .And.;
cQrySE2 += " SE2.E2_DTFATUR > '"+DTOS(dData)+"' ) OR "
cQrySE2 += " SUBSTRING(SE2.E2_FATURA,1,3) = '   ')  AND SE2.E2_FLUXO <> 'N'" //Empty(SE2->E2_FATURA)) ) .And. SE2->E2_FLUXO != "N"
cQrySE2 += " AND SE2.E2_SALDO > 0 "
cQrySE2 += " AND SE2.D_E_L_E_T_ <> '*'"

TCQuery cQrySE2 NEW ALIAS "TSE2"

TCSetField("TSE2","E2_SALDO","N",17,02)

DbSelectArea("TSE2")
While !Eof()
			//Alert("Entrei no If do While do ItaSldPagar") 
		If ( !TSE2->E2_TIPO $ MVPAGANT+"/"+MV_CPNEG )
			If ( !lMovSE5 )
				nSaldo += xMoeda(TSE2->E2_SALDO,TSE2->E2_MOEDA,1,dData)
				nSaldo -= SomaAbat(TSE2->E2_PREFIXO,TSE2->E2_NUM,TSE2->E2_PARCELA,"P",1,dData,TSE2->E2_FORNECE)
			Else
				nSaldo += TitSaldo(TSE2->E2_PREFIXO,TSE2->E2_NUM,TSE2->E2_PARCELA,TSE2->E2_TIPO,TSE2->E2_NATUREZ,"P",TSE2->E2_FORNECE,nMoeda,,dData,TSE2->E2_LOJA)
				nSaldo -= SomaAbat(TSE2->E2_PREFIXO,TSE2->E2_NUM,TSE2->E2_PARCELA,"P",1,dData,TSE2->E2_FORNECE)
			EndIf
		Else
			If ( !lMovSE5 .And. !TSE2->E2_TIPO $ MVABATIM )
				nSaldo -= TitSaldo(TSE2->E2_PREFIXO,TSE2->E2_NUM,TSE2->E2_PARCELA,TSE2->E2_TIPO,TSE2->E2_NATUREZ,"P",TSE2->E2_FORNECE,nMoeda,,dData,TSE2->E2_LOJA)
			Else
				nSaldo -= xMoeda(TSE2->E2_SALDO,TSE2->E2_MOEDA,1,dData)
			EndIf
		EndIf
	dbSelectArea("TSE2")
	dbSkip()
EndDo
DbSelectArea("TSE2")
DbCloseArea()

/*
bCondSE2  := {|| !Eof() .And. xFilial() == SE2->E2_FILIAL .And.;
	SE2->E2_VENCREA <= dData }
	*/
/*
bCondSE2  := {|| !Eof() .And. xFilial() == SE2->E2_FILIAL .And.;
	SE2->E2_VENCREA <= dData .And. SE2->E2_SALDO > 0}
If ( lDtAnterior )
	dbSeek(xFilial(),.T.)
Else
	dbSeek(xFilial()+Dtos(dData))
EndIf
While ( Eval(bCondSe2) )
    
	If ( SE2->E2_EMIS1 <= dData .And. ;
			!SE2->E2_TIPO $MVPROVIS+"/"+MVABATIM .AND. ;
			((!Empty(SE2->E2_FATURA).And.;
			Substr(SE2->E2_FATURA,1,6)=="NOTFAT" ) .Or.;
			(!Empty(SE2->E2_FATURA) .And.;
			Substr(SE2->E2_FATURA,1,6)!="NOTFAT" .And.;
			SE2->E2_DTFATUR > dData ) .Or.;
			Empty(SE2->E2_FATURA)) ) .And. SE2->E2_FLUXO != "N"
			Alert("Entrei no If do While do ItaSldPagar") 
		If ( !SE2->E2_TIPO $ MVPAGANT+"/"+MV_CPNEG )
			If ( !lMovSE5 )
				nSaldo += xMoeda(SE2->E2_SALDO,SE2->E2_MOEDA,1,dData)
				nSaldo -= SomaAbat(SE2->E2_PREFIXO,SE2->E2_NUM,SE2->E2_PARCELA,"P",1,dData,SE2->E2_FORNECE)
			Else
				nSaldo += TitSaldo(SE2->E2_PREFIXO,SE2->E2_NUM,SE2->E2_PARCELA,SE2->E2_TIPO,SE2->E2_NATUREZ,"P",SE2->E2_FORNECE,nMoeda,,dData,SE2->E2_LOJA)
				nSaldo -= SomaAbat(SE2->E2_PREFIXO,SE2->E2_NUM,SE2->E2_PARCELA,"P",1,dData,SE2->E2_FORNECE)
			EndIf
		Else
			If ( !lMovSE5 .And. !SE2->E2_TIPO $ MVABATIM )
				nSaldo -= TitSaldo(SE2->E2_PREFIXO,SE2->E2_NUM,SE2->E2_PARCELA,SE2->E2_TIPO,SE2->E2_NATUREZ,"P",SE2->E2_FORNECE,nMoeda,,dData,SE2->E2_LOJA)
			Else
				nSaldo -= xMoeda(SE2->E2_SALDO,SE2->E2_MOEDA,1,dData)
			EndIf
		EndIf
	EndIf
	dbSelectArea("SE2")
	dbSkip()
EndDo
dbSelectArea("SE2")
RetIndex("SE2")
dbClearFilter()
dbSetOrder(aAreaSE2[1])
dbGoto(aAreaSE2[2])
*/
/*
dbSelectArea(aArea[1])
dbSetOrder(aArea[2])
dbGoto(aArea[3])
*/
Return(nSaldo)
/*
Static Function RunSldPg
/*
�������������������������������������������������������������������������Ĵ��
���Parametros� dData   : Data do Movimento a Receber - Default dDataBase  ���
���          � nMoeda  : Moeda do Saldo Bancario - Defa 1                 ���
���          � lDtAnterior: Se .T. Ate a Data,.F. Somente Data - Defa .T. ���
���          � lMovSE5 : Se .T. considera o saldo do SE5 - Defa .T.       ���
�������������������������������������������������������������������������Ĵ��
��� Uso      � SIGAFIN                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������


Function SldPagar(dData,nMoeda,lDtAnterior,lMovSe5)
*/
/*
   ProcRegua(6)
   nCt := 1
   WhIle nCt <= 6
      IncProc("Calc. Saldo A Pagar, Proc. "+Alltrim(Str(nCt))+" / 6")
      If nCt == 1
         nSld1MPag := ItaSldPagar(dDataDe,1,.T.,.T.)
      ElseIf nCt == 2
         nSld2MPag := ItaSldPagar(LastDay(dData02),1,.T.,.T.) 
      ElseIf nCt == 3
         nSld3MPag := ItaSldPagar(LastDay(dData03),1,.T.,.T.) 
      ElseIf nCt == 4
         nSld1MAnt := ItaSldPagar(FirstDay(dDataDe)-1,1,.T.,.T.)
      ElseIf nCt == 5
         nSld2MAnt := ItaSldPagar(FirstDay(dData02)-1,1,.T.,.T.)
      ElseIf nCt == 6
         nSld3MAnt := ItaSldPagar(FirstDay(dData03)-1,1,.T.,.T.)
      EndIf
      nCt ++
   EndDo
Return
*/

/*
Static Function RunSldRe
   ProcRegua(3)
   nCt := 1
   WhIle nCt <= 3
      IncProc("Calc. Saldo A Receber, Proc. "+Alltrim(Str(nCt))+" / 3")
      If nCt == 1
         nSld1MRec := ItaSldReceb(dDataDe,1,.T.,.T.)
      ElseIf nCt == 2
         nSld2MRec := ItaSldReceb(LastDay(dData02),1,.T.,.T.) 
      ElseIf nCt == 3
         nSld3MRec := ItaSldReceb(LastDay(dData03),1,.T.,.T.) 
      EndIf
      nCt ++
   EndDo
Return
*/

Static Function ItaSldReceb(dData,nMoeda,lDtAnterior,lMovSE5)

Local aArea     := { Alias() , IndexOrd() , Recno() }
Local aAreaSE1  := { SE1->(IndexOrd()), SE1->(Recno()) }
Local bCondSE1
Local nSaldo    := 0
#IFDEF TOP
	Local cFiltro   := ""
#ENDIF

//Alert("Estou no ItaSldReceb")
// ������������������������������������������������������Ŀ
// � Testa os parametros vindos do Excel                  �
// ��������������������������������������������������������
nMoeda      := If(Empty(nMoeda),1,nMoeda)
dData       := If(Empty(dData),dDataBase,dData)
lDtAnterior := BoolWindow(lDtAnterior)
lMovSe5     := BoolWindow(lMovSe5)
If ( ValType(nMoeda) == "C" )
	nMoeda      := Val(nMoeda)
EndIf
dData       := DataWindow(dData)

// Quando eh chamada do Excel, estas variaveis estao em branco
IF Empty(MVABATIM) .Or.;
	Empty(MV_CRNEG) .Or.;
	Empty(MVRECANT) .Or.;
	Empty(MVPROVIS)
	CriaTipos()
Endif

dbSelectArea("SE1")
dbSetOrder(7)
//Ita - 11/03/2008
/*
bCondSE1  := {|| !Eof() .And. xFilial() == SE1->E1_FILIAL .And.;
	SE1->E1_VENCREA <= dData }
*/
cQrySE1 := "SELECT SE1.E1_PREFIXO,SE1.E1_NUM,SE1.E1_PARCELA,SE1.E1_CLIENTE,SE1.E1_LOJA,"
cQrySE1 += "SE1.E1_SALDO,SE1.E1_MOEDA,SE1.E1_TIPO,SE1.E1_NATUREZ"
cQrySE1 += " FROM "+RetSQLName("SE1")+" SE1 "
cQrySE1 += " WHERE "
cQrySE1 += " SE1.E1_EMISSAO <= '"+DTOS(dData)+"' AND " 
cQrySE1 += " SE1.E1_TIPO NOT IN "+FormatIn(MVPROVIS+MVABATIM,"|")+" AND " //!SE1->E1_TIPO $ MVPROVIS+"/"+MVABATIM .AND. 
cQrySE1 += " ((SUBSTRING(SE1.E1_FATURA,1,3) <> '   ' AND "			//((!Empty(SE1->E1_FATURA).And.
cQrySE1 += " SUBSTRING(SE1.E1_FATURA,1,6)= 'NOTFAT' ) OR " 			//Substr(SE1->E1_FATURA,1,6)=="NOTFAT" ) .Or.
cQrySE1 += " (SUBSTRING(SE1.E1_FATURA,1,3) <> '   ' AND " //(!Empty(SE1->E1_FATURA) .And.
cQrySE1 += " SUBSTRING(SE1.E1_FATURA,1,6) <> 'NOTFAT' AND "//Substr(SE1->E1_FATURA,1,6)!="NOTFAT" .And.
cQrySE1 += " SE1.E1_DTFATUR > '"+DTOS(dData)+"' ) OR " //SE1->E1_DTFATUR > dData ) .Or.
cQrySE1 += " SUBSTRING(SE1.E1_FATURA,1,3) = '   ' )  AND  SE1.E1_FLUXO <> 'N'" //Empty(SE1->E1_FATURA)) ) .And. SE1->E1_FLUXO != "N"
cQrySE1 += " AND SE1.E1_SALDO > 0 "
cQrySE1 += " AND SE1.D_E_L_E_T_ <> '*'"

TCQuery cQrySE1 NEW ALIAS "TSE1" 

TCSetField("TSE1","E1_SALDO","N",17,02)

DbSelectArea("TSE1")
While !Eof()
			//Alert("Entei no If do While do ItaSldReceb")
		If ( !TSE1->E1_TIPO $ MVRECANT+"/"+MV_CRNEG )
			If ( !lMovSE5 )
				nSaldo += xMoeda(TSE1->E1_SALDO,TSE1->E1_MOEDA,1,dData )
				nSaldo -= SomaAbat(TSE1->E1_PREFIXO,TSE1->E1_NUM,TSE1->E1_PARCELA,"R",1,dData,TSE1->E1_CLIENTE)
			Else
				nSaldo += TitSaldo(TSE1->E1_PREFIXO,TSE1->E1_NUM,TSE1->E1_PARCELA,TSE1->E1_TIPO,TSE1->E1_NATUREZ,"R",TSE1->E1_CLIENTE,nMoeda,,dData,TSE1->E1_LOJA)
				nSaldo -= SomaAbat(TSE1->E1_PREFIXO,TSE1->E1_NUM,TSE1->E1_PARCELA,"R",1,dData,TSE1->E1_CLIENTE)
			EndIf
		Else
			If ( !lMovSE5 )
				nSaldo -= TitSaldo(TSE1->E1_PREFIXO,TSE1->E1_NUM,TSE1->E1_PARCELA,TSE1->E1_TIPO,TSE1->E1_NATUREZ,"R",TSE1->E1_CLIENTE,nMoeda,,dData,TSE1->E1_LOJA)
			Else
				nSaldo -= xMoeda(TSE1->E1_SALDO,TSE1->E1_MOEDA,1,dData)
			EndIf
		EndIf
	dbSelectArea("TSE1")
	dbSkip()
EndDo
dbSelectArea("TSE1")
DbCloseArea()

/*
bCondSE1  := {|| !Eof() .And. xFilial() == SE1->E1_FILIAL .And.;
	SE1->E1_VENCREA <= dData .And. SE1->E1_SALDO > 0}
If ( lDtAnterior )
	dbSeek(xFilial(),.T.)
Else
	dbSeek(xFilial()+Dtos(dData))
EndIf
While ( Eval(bCondSe1) )
	If ( SE1->E1_EMISSAO <= dData .And. ;
			!SE1->E1_TIPO $ MVPROVIS+"/"+MVABATIM .AND. ;
			((!Empty(SE1->E1_FATURA).And.;
			Substr(SE1->E1_FATURA,1,6)=="NOTFAT" ) .Or.;
			(!Empty(SE1->E1_FATURA) .And.;
			Substr(SE1->E1_FATURA,1,6)!="NOTFAT" .And.;
			SE1->E1_DTFATUR > dData ) .Or.;
			Empty(SE1->E1_FATURA)) ) .And. SE1->E1_FLUXO != "N"
			Alert("Entei no If do While do ItaSldReceb")
		If ( !SE1->E1_TIPO $ MVRECANT+"/"+MV_CRNEG )
			If ( !lMovSE5 )
				nSaldo += xMoeda(SE1->E1_SALDO,SE1->E1_MOEDA,1,dData )
				nSaldo -= SomaAbat(SE1->E1_PREFIXO,SE1->E1_NUM,SE1->E1_PARCELA,"R",1,dData,SE1->E1_CLIENTE)
			Else
				nSaldo += TitSaldo(SE1->E1_PREFIXO,SE1->E1_NUM,SE1->E1_PARCELA,SE1->E1_TIPO,SE1->E1_NATUREZ,"R",SE1->E1_CLIENTE,nMoeda,,dData,SE1->E1_LOJA)
				nSaldo -= SomaAbat(SE1->E1_PREFIXO,SE1->E1_NUM,SE1->E1_PARCELA,"R",1,dData,SE1->E1_CLIENTE)
			EndIf
		Else
			If ( !lMovSE5 )
				nSaldo -= TitSaldo(SE1->E1_PREFIXO,SE1->E1_NUM,SE1->E1_PARCELA,SE1->E1_TIPO,SE1->E1_NATUREZ,"R",SE1->E1_CLIENTE,nMoeda,,dData,SE1->E1_LOJA)
			Else
				nSaldo -= xMoeda(SE1->E1_SALDO,SE1->E1_MOEDA,1,dData)
			EndIf
		EndIf
	Endif
	dbSelectArea("SE1")
	dbSkip()
EndDo

dbSelectArea("SE1")
RetIndex("SE1")
dbClearFilter()
dbSetOrder(aAreaSE1[1])
dbGoto(aAreaSE1[2])
*/
/*
dbSelectArea(aArea[1])
dbSetOrder(aArea[2])
dbGoto(aArea[3])
*/
Return(nSaldo)

Static Function RunSelRG

//���������������������������������������������������������������������Ŀ
//� �rea de Instru��es SQL - Querys.                                    �
//�����������������������������������������������������������������������

If lComercial .Or. lIndustria

   /**************
   * Calculando Clientes (Comercial)
   * Conceitos COMAFAL
   * Clientes Novos: S�o aqueles cujo a data de cadastro(A1_DTCAD-Customizado) refere-se ao m�s analizado.
   * Clientes Ativos: S�o os clientes novos que nunca compraram, somados aos clientes que compraram nos �ltimos 
   *                  tr�s(3) meses.
   * Clientes Inativos: S�o aqueles que n�o compraram nos �ltimos 3 meses.
   * Clientes Atendidos: S�o aqueles que compraram no m�s analizado.
   ***************/
   
   ProcRegua(6) 
   
   For nMes := 1 To 6   // De 1 a 3(Clientes Novos) - De 4 a 6(Clientes Atendidos)
   //A partir do discutido durante a reuni�o do dia 08/04/2008 na COMAFAL/SP esses clientes,Clientes Novos,
   //ter�o novo conceito, ser�o os clientes, cuja primeira compra seja dento do m�s em an�lise.
   //A decis�o foi tomada pelo Sr. Gilmar Filho em conjunto com os presentes na reuni�o.

      
      IncProc("Calculando Clientes Novos/Atendidos, Processo "+Alltrim(Str(nMes))+" / 6 ")//9 ")
      
      If nMes <= 3
         cQrySD2 := "SELECT COUNT(DISTINCT SA1.A1_COD+SA1.A1_LOJA) AS CLIENTES"
      ElseIf nMes >= 4
         cQrySD2 := "SELECT COUNT(DISTINCT SD2.D2_CLIENTE+SD2.D2_LOJA) AS CLIENTES"
      EndIf
      
      If nMes <= 3
         cQrySD2 += " FROM "+RetSQLName("SA1")+" SA1"
         cQrySD2 += " WHERE SA1.A1_FILIAL = '"+xFilial("SA1")+"'"
      Else
         cQrySD2 += " FROM "+RetSQLName("SD2")+" SD2"
         cQrySD2 += " WHERE SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
      EndIf
      
      /*Comentado para adotar o novo conceito, logo abaixo, adotado na reuni�o do dia 08/04/2008
      If nMes <= 3 
         If nMes == 1
            cQrySD2 += "   AND SUBSTRING(SA1.A1_DTCAD,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         ElseIf nMes == 2
            cQrySD2 += "   AND SUBSTRING(SA1.A1_DTCAD,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
         ElseIf nMes == 3
            cQrySD2 += "   AND SUBSTRING(SA1.A1_DTCAD,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
         EndIf
      EndIf
      */
      If nMes <= 3 
         If nMes == 1
            cQrySD2 += "   AND SUBSTRING(SA1.A1_PRICOM,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         ElseIf nMes == 2
            cQrySD2 += "   AND SUBSTRING(SA1.A1_PRICOM,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
         ElseIf nMes == 3
            cQrySD2 += "   AND SUBSTRING(SA1.A1_PRICOM,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
         EndIf
      EndIf      
      
      
      If nMes >= 4
         If nMes == 4
            cQrySD2 += "   AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         ElseIf nMes == 5
            cQrySD2 += "   AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
         ElseIf nMes == 6
            cQrySD2 += "   AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
         EndIf
      EndIf
      
      
      If nMes <= 3
         cQrySD2 += "   AND SA1.D_E_L_E_T_ <> '*'"
      Else
         //Filtra CFOPs de Vendas - Conforme Solicita��o de Patr�cia(Auditoria COMAFAL)
         cQrySD2 += "   AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
         cQrySD2 += "   AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
         
         cQrySD2 += "   AND SD2.D_E_L_E_T_ <> '*'"
      EndIf
      


         MemoWrite("C:\TEMP\CliNovos_Atendidos"+Alltrim(Str(nMes))+".SQL",cQrySD2)
      
      TCQuery cQrySD2 NEW ALIAS "TCLI"
   
      If nMes <= 3
         If nMes == 1
            DbSelectArea("TCLI")
            nCliN01Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf

         If nMes == 2
            DbSelectArea("TCLI")
            nCliN02Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf

         If nMes == 3
            DbSelectArea("TCLI")
            nCliN03Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("TCLI")
            nClAtd1Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf

         If nMes == 5
            DbSelectArea("TCLI")
            nClAtd2Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf

         If nMes == 6
            DbSelectArea("TCLI")
            nClAtd3Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf   
      EndIf

   Next nMes
   
   ProcRegua(3)
   
   For nMes := 1 To 3 //Clientes Ativos
      
      IncProc("Calc. Clientes Ativos, Processo "+Alltrim(Str(nMes))+" / 3")
      
      cQrySA1 := "SELECT COUNT(*) AS CLIENTES"
      cQrySA1 += " FROM "+RetSQLName("SA1")+" SA1"
      cQrySA1 += " WHERE SA1.A1_FILIAL = '"+xFilial("SA1")+"'"
      cQrySA1 += " AND (SELECT COUNT(*)"
      cQrySA1 += "      FROM "+RetSQLName("SD2")+" SD2"
      cQrySA1 += "      WHERE "//Five Solutions 14/07/2008 - SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
      If nMes == 1
         cQrySA1 += "     SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dDataDe) - 90)+"' AND '"+DTOS(LastDay(dDataDe))+"'"
      ElseIf nMes == 2
         cQrySA1 += "     SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dData02) - 90)+"' AND '"+DTOS(LastDay(dData02))+"'"
      ElseIf nMes == 3
         cQrySA1 += "     SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dData03) - 90)+"' AND '"+DTOS(LastDay(dData03))+"'"
      EndIf
      //Filtra CFOPs de Vendas - Conforme Solicita��o de Patr�cia(Auditoria COMAFAL)
      cQrySA1 += "        AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQrySA1 += "        AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      cQrySA1 += "        AND SD2.D2_CLIENTE = SA1.A1_COD"
      cQrySA1 += "        AND SD2.D2_LOJA = SA1.A1_LOJA"
      cQrySA1 += "        AND SD2.D_E_L_E_T_ <> '*') > 0 "
      
      If nMes == 1
         cQrySA1 += "     AND SUBSTRING(SA1.A1_DTCAD,1,6) < '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQrySA1 += "     AND SUBSTRING(SA1.A1_DTCAD,1,6) < '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQrySA1 += "     AND SUBSTRING(SA1.A1_DTCAD,1,6) < '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf

      
      cQrySA1 += " AND SA1.D_E_L_E_T_ <> '*'"
      
      MemoWrite("C:\TEMP\CliAtivos"+Alltrim(Str(nMes))+".SQL",cQrySA1)
      
      TCQuery cQrySA1 NEW ALIAS "TCLI"
      
         If nMes == 1
            DbSelectArea("TCLI")
            nClAtv1Mes := TCLI->CLIENTES
            //Acrescentando Clientes Novos aos Clientes Ativos
            //Ita - Comentado em decorrencia do novo conceito(08/04/08) - nClAtv1Mes += nCliN01Mes
            DbCloseArea()
         EndIf

         If nMes == 2
            DbSelectArea("TCLI")
            nClAtv2Mes := TCLI->CLIENTES
            //Acrescentando Clientes Novos aos Clientes Ativos
            //Ita - Comentado em decorrencia do novo conceito(08/04/08) - nClAtv2Mes += nCliN02Mes
            DbCloseArea()
         EndIf

         If nMes == 3
            DbSelectArea("TCLI")
            nClAtv3Mes := TCLI->CLIENTES
            //Acrescentando Clientes Novos aos Clientes Ativos
            //Ita - Comentado em decorrencia do novo conceito(08/04/08) - nClAtv3Mes += nCliN03Mes
            DbCloseArea()
         EndIf   

   Next nMes  

      ProcRegua(3)
   
   For nMes := 1 To 3 //Clientes Ativos Sem Compras
   //A partir do discutido durante a reuni�o do dia 08/04/2008 na COMAFAL/SP esses clientes,Ativos Sem Compras,
   //ser�o denominados "Prospects" e ter�o uma linha exclusiva de apresenta��o de sua quantidade.
   //A decis�o foi tomada pelo Sr. Gilmar Filho em conjunto com os presentes na reuni�o.
      
      //IncProc("Calc.Cli.Ativos s/Compras, Processo "+Alltrim(Str(nMes))+" / 3")
      IncProc("Calc. Prospects, Processo "+Alltrim(Str(nMes))+" / 3")
      
      cQrySA1 := "SELECT COUNT(*) AS CLIENTES"
      cQrySA1 += " FROM "+RetSQLName("SA1")+" SA1"
      cQrySA1 += " WHERE SA1.A1_FILIAL = '"+xFilial("SA1")+"'"
      /*
      
      cQrySA1 += " AND (SELECT COUNT(*)"
      cQrySA1 += "      FROM "+RetSQLName("SD2")+" SD2"
      cQrySA1 += "      WHERE SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
      
      //If nMes == 1
      //   cQrySA1 += "     AND SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dDataDe) - 90)+"' AND '"+DTOS(LastDay(dDataDe))+"'"
      //ElseIf nMes == 2
      //   cQrySA1 += "     AND SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dData02) - 90)+"' AND '"+DTOS(LastDay(dData02))+"'"
      //ElseIf nMes == 3
      //   cQrySA1 += "     AND SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dData03) - 90)+"' AND '"+DTOS(LastDay(dData03))+"'"
      //EndIf
      

      If nMes == 1
         cQrySA1 += "     AND SD2.D2_EMISSAO BETWEEN '19800101' AND '"+DTOS(LastDay(dDataDe))+"'"
      ElseIf nMes == 2
         cQrySA1 += "     AND SD2.D2_EMISSAO BETWEEN '19800101' AND '"+DTOS(LastDay(dData02))+"'"
      ElseIf nMes == 3
         cQrySA1 += "     AND SD2.D2_EMISSAO BETWEEN '19800101' AND '"+DTOS(LastDay(dData03))+"'"
      EndIf
            
      //Filtra CFOPs de Vendas - Conforme Solicita��o de Patr�cia(Auditoria COMAFAL)
      cQrySA1 += "        AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQrySA1 += "        AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      cQrySA1 += "        AND SD2.D2_CLIENTE = SA1.A1_COD"
      cQrySA1 += "        AND SD2.D2_LOJA = SA1.A1_LOJA"
      cQrySA1 += "        AND SD2.D_E_L_E_T_ <> '*') = 0 "
      
      */

      /*
      If nMes == 1
         cQrySA1 += " AND SA1.A1_DTCAD BETWEEN '"+DTOS(LastDay(dDataDe) - 90)+"' AND '"+DTOS(FirstDay(dDataDe)-1)+"'"
      ElseIf nMes == 2
         cQrySA1 += " AND SA1.A1_DTCAD BETWEEN '"+DTOS(LastDay(dData02) - 90)+"' AND '"+DTOS(FirstDay(dData02)-1)+"'"
      ElseIf nMes == 3
         cQrySA1 += " AND SA1.A1_DTCAD BETWEEN '"+DTOS(LastDay(dData03) - 90)+"' AND '"+DTOS(FirstDay(dData03)-1)+"'"
      EndIf
      */
      If nMes == 1
         cQrySA1 += " AND SA1.A1_DTCAD BETWEEN '19800101' AND '"+DTOS(LastDay(dDataDe))+"'"
      ElseIf nMes == 2
         cQrySA1 += " AND SA1.A1_DTCAD BETWEEN '19800101' AND '"+DTOS(LastDay(dData02))+"'"
      ElseIf nMes == 3
         cQrySA1 += " AND SA1.A1_DTCAD BETWEEN '19800101' AND '"+DTOS(LastDay(dData03))+"'"
      EndIf
      
      cQrySA1 += " AND SUBSTRING(SA1.A1_PRICOM,1,3) = '   ' AND SUBSTRING(SA1.A1_ULTCOM,1,3) = '   '" //Nunca Comprou (Teste para evitar considerar SD2)
      
      cQrySA1 += " AND SA1.D_E_L_E_T_ <> '*'"
      
      MemoWrite("C:\TEMP\CliASC"+Alltrim(Str(nMes))+".SQL",cQrySA1)
      
      TCQuery cQrySA1 NEW ALIAS "TCLI"
      
      If nMes == 1
         DbSelectArea("TCLI")
         nASC1Mes := TCLI->CLIENTES
         DbCloseArea()
      EndIf

      If nMes == 2
         DbSelectArea("TCLI")
         nASC2Mes := TCLI->CLIENTES
         DbCloseArea()
      EndIf

      If nMes == 3
         DbSelectArea("TCLI")
         nASC3Mes := TCLI->CLIENTES
         DbCloseArea()
         
      EndIf   
   Next nMes  
   
   //Acrescentando Clientes Ativos Sem Compras aos Clientes Ativos
   //Ita - Comentado em decorrencia do novo conceito(08/04/08)
   //nClAtv1Mes += nASC1Mes
   //nClAtv2Mes += nASC2Mes
   //nClAtv3Mes += nASC3Mes

   ProcRegua(3)
   
   For nMes := 1 To 3 //Clientes Inativos
      
      IncProc("Calc. Clientes Inativos, Processo "+Alltrim(Str(nMes))+" / 3")
      
      cQrySA1 := "SELECT COUNT(*) AS CLIENTES"
      cQrySA1 += " FROM "+RetSQLName("SA1")+" SA1"
      cQrySA1 += " WHERE SA1.A1_FILIAL = '"+xFilial("SA1")+"'"
      cQrySA1 += " AND (SELECT COUNT(*)"
      cQrySA1 += "      FROM "+RetSQLName("SD2")+" SD2"
      cQrySA1 += "      WHERE "//Five Solutions - 14/07/2008 - SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
      If nMes == 1
         cQrySA1 += "     SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dDataDe) - 90)+"' AND '"+DTOS(LastDay(dDataDe))+"'"
      ElseIf nMes == 2
         cQrySA1 += "     SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dData02) - 90)+"' AND '"+DTOS(LastDay(dData02))+"'"
      ElseIf nMes == 3
         cQrySA1 += "     SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dData03) - 90)+"' AND '"+DTOS(LastDay(dData03))+"'"
      EndIf
      //Filtra CFOPs de Vendas - Conforme Solicita��o de Patr�cia(Auditoria COMAFAL)
      cQrySA1 += "        AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQrySA1 += "        AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      cQrySA1 += "        AND SD2.D2_CLIENTE = SA1.A1_COD"
      cQrySA1 += "        AND SD2.D2_LOJA = SA1.A1_LOJA"
      cQrySA1 += "        AND SD2.D_E_L_E_T_ <> '*') = 0 "
      
      If nMes == 1
         cQrySA1 += " AND SA1.A1_DTCAD < '"+DTOS(LastDay(dDataDe) - 90)+"'"
      ElseIf nMes == 2
         cQrySA1 += " AND SA1.A1_DTCAD < '"+DTOS(LastDay(dData02) - 90)+"'"
      ElseIf nMes == 3
         cQrySA1 += " AND SA1.A1_DTCAD < '"+DTOS(LastDay(dData03) - 90)+"'"
      EndIf
      
      cQrySA1 += " AND SA1.D_E_L_E_T_ <> '*'"
      
      MemoWrite("C:\TEMP\CliInativos"+Alltrim(Str(nMes))+".SQL",cQrySA1)
      
      TCQuery cQrySA1 NEW ALIAS "TCLI"
      
      If nMes == 1
         DbSelectArea("TCLI")
         nClIna1Mes := TCLI->CLIENTES
         DbCloseArea()
      EndIf

      If nMes == 2
         DbSelectArea("TCLI")
         nClIna2Mes := TCLI->CLIENTES
         DbCloseArea()
      EndIf

      If nMes == 3
         DbSelectArea("TCLI")
         nClIna3Mes := TCLI->CLIENTES
         DbCloseArea()
      EndIf   
   Next nMes  

   /**************
   * Calculando Pedidos (Quantidade/Valor)
   ***************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. Pedidos, Processo "+Alltrim(Str(nMes))+" / 3")
      
      //Pedidos Emitidos - Quantidade
      cQrySC5 := "SELECT COUNT(DISTINCT SC5.C5_NUM) AS PEDIDOS" 
      cQrySC5 += " FROM "+RetSQLName("SC5")+" SC5,"+RetSQLName("SC6")+" SC6"
      cQrySC5 += " WHERE SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
      cQrySC5 += " AND SC6.C6_FILIAL = '"+xFilial("SC6")+"'"
      cQrySC5 += " AND SC5.C5_NUM = SC6.C6_NUM"
      cQrySC5 += " AND SC6.C6_CF IN "+FormatIn(cCFOPVenda,",")
      If nMes == 1
         cQrySC5 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQrySC5 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQrySC5 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQrySC5 += " AND SC5.D_E_L_E_T_ <> '*'"
      cQrySC5 += " AND SC6.D_E_L_E_T_ <> '*'"
      
      MemoWrite("C:\TEMP\PedEmitidos.SQL",cQrySC5)
      
      TCQuery cQrySC5 NEW ALIAS "TSC5"
      
      //Pedidos Emitidos - Valor
      cQrySC6 := "SELECT SUM(C6_VALOR) AS VLRPED" 
      cQrySC6 += " FROM "+RetSQLName("SC5")+" SC5,"+RetSQLName("SC6")+" SC6"
      cQrySC6 += " WHERE SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
      cQrySC6 += "   AND SC6.C6_FILIAL = '"+xFilial("SC6")+"'"
      cQrySC6 += "   AND SC5.C5_NUM = SC6.C6_NUM"
      cQrySC6 += "   AND SC6.C6_CF IN "+FormatIn(cCFOPVenda,",")
      If nMes == 1
         cQrySC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQrySC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQrySC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQrySC6 += " AND SC5.D_E_L_E_T_ <> '*'"
      cQrySC6 += " AND SC6.D_E_L_E_T_ <> '*'"
      
      TCQuery cQrySC6 NEW ALIAS "TSC6"

      //Pedidos Faturados(M�s) - Quantidade 
      cQuery := "SELECT SUBSTRING(SD2.D2_EMISSAO,1,6),COUNT(DISTINCT D2_PEDIDO) AS PEDFAT FROM " 
      cQuery += RetSqlName('SD2') + " SD2,"+RetSQLName("SB1")+" SB1,"+RetSQLName("SF2")+" SF2,"+RetSQLName("SC5")+" SC5"
      cQuery += " WHERE SD2.D2_FILIAL  = '"+xFilial("SD2")+"' AND  SD2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SB1.B1_FILIAL  = '"+xFilial("SB1")+"' AND  SB1.D_E_L_E_T_ = ' ' "
      cQuery += " AND SF2.F2_FILIAL  = '"+xFilial("SF2")+"' AND  SF2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SC5.C5_FILIAL  = '"+xFilial("SC5")+"' AND  SC5.D_E_L_E_T_ = ' ' "
      cQuery += " AND SD2.D2_COD = SB1.B1_COD"
      //cQuery += " AND SUBSTRING(SB1.B1_GRUPO,1,2) <> '80'" //Faturamento sem Sucatas
      cQuery += " AND SF2.F2_DOC = SD2.D2_DOC"
      cQuery += " AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQuery += " AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQuery += " AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      cQuery += " AND SD2.D2_PEDIDO = SC5.C5_NUM"
      
      If nMes == 1
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
         cQuery += " AND  SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
      ElseIf nMes == 2
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
         cQuery += " AND  SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
      ElseIf nMes == 3
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
         cQuery += " AND  SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
      EndIf
      
      cQuery += " GROUP BY SUBSTRING(SD2.D2_EMISSAO,1,6) "
      
      MemoWrite("C:\TEMP\PedFaturaMesQtd.SQL",cQrySD2)
      
      TCQuery cQuery NEW ALIAS "TSD2"
      
      //Pedidos Faturados(M�s) - Valor
      cQuery := " SELECT SUBSTRING(SD2.D2_EMISSAO,1,6) AS EMISSAO,SUM(D2_TOTAL+CASE WHEN SD2.D2_TIPO <> 'P' THEN D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA ELSE 0 END) AS VLRFATPED,SUM(D2_QUANT) AS QTDFATPED FROM "
      cQuery += RetSqlName('SD2') + " SD2,"+RetSQLName("SB1")+" SB1,"+RetSQLName("SF2")+" SF2,"+RetSQLName("SC5")+" SC5"
      cQuery += " WHERE SD2.D2_FILIAL  = '"+xFilial("SD2")+"' AND  SD2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SB1.B1_FILIAL  = '"+xFilial("SB1")+"' AND  SB1.D_E_L_E_T_ = ' ' "
      cQuery += " AND SF2.F2_FILIAL  = '"+xFilial("SF2")+"' AND  SF2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SC5.C5_FILIAL  = '"+xFilial("SC5")+"' AND  SC5.D_E_L_E_T_ = ' ' "
      cQuery += " AND SD2.D2_COD = SB1.B1_COD"
      //cQuery += " AND SUBSTRING(SB1.B1_GRUPO,1,2) <> '80'" //Faturamento sem Sucatas
      cQuery += " AND SF2.F2_DOC = SD2.D2_DOC"
      cQuery += " AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQuery += " AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQuery += " AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      cQuery += " AND SD2.D2_PEDIDO = SC5.C5_NUM"
      
      If nMes == 1
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
         cQuery += " AND  SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
      ElseIf nMes == 2
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
         cQuery += " AND  SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
      ElseIf nMes == 3
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
         cQuery += " AND  SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
      EndIf
      
      cQuery += " GROUP BY SUBSTRING(SD2.D2_EMISSAO,1,6) "
      cQuery += " ORDER BY SUBSTRING(SD2.D2_EMISSAO,1,6) ASC"
      
      MemoWrite("C:\TEMP\PedFaturaMesVlr"+Alltrim(Str(nMes))+".SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "T2SD2"
      
      //Pedidos Pendentes(M�s) - Quantidade 
      cQry2SC6 := "SELECT COUNT(DISTINCT SC6.C6_NUM) AS PEDPEND" 
      cQry2SC6 += " FROM "+RetSQLName("SC5")+" SC5,"+RetSQLName("SC6")+" SC6"
      cQry2SC6 += " WHERE SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
      cQry2SC6 += "   AND SC6.C6_FILIAL = '"+xFilial("SC6")+"'"
      cQry2SC6 += "   AND SC5.C5_NUM = SC6.C6_NUM"
      cQry2SC6 += "   AND SC6.C6_CF IN "+FormatIn(cCFOPVenda,",")
      cQry2SC6 += "   AND (SUBSTRING(SC6.C6_NOTA,1,3) = '   ' OR (SUBSTRING(SC6.C6_NOTA,1,3) <> '   ' AND (SELECT SUBSTRING(SF2.F2_EMISSAO,1,6)"
      cQry2SC6 += "                                                                                         FROM "+RetSQLName("SF2")+" SF2"
      cQry2SC6 += "                                                                                        WHERE SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      cQry2SC6 += "                                                                                          AND SF2.F2_DOC = SC6.C6_NOTA"
      cQry2SC6 += "                                                                                          AND SF2.F2_SERIE = SC6.C6_SERIE"
      If nMes == 1
         cQry2SC6 += "                                                                                          AND SF2.D_E_L_E_T_ <> '*') > '"+Substr(DTOS(dDataDe),1,6)+"'))"
         cQry2SC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQry2SC6 += "                                                                                          AND SF2.D_E_L_E_T_ <> '*') > '"+Substr(DTOS(dData02),1,6)+"'))" 
         cQry2SC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQry2SC6 += "                                                                                          AND SF2.D_E_L_E_T_ <> '*') > '"+Substr(DTOS(dData03),1,6)+"'))"  
         cQry2SC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQry2SC6 += " AND SC5.D_E_L_E_T_ <> '*'"
      cQry2SC6 += " AND SC6.D_E_L_E_T_ <> '*'"
      
      MemoWrite("C:\TEMP\PedPenQtd.SQL",cQry2SC6)
      
      TCQuery cQry2SC6 NEW ALIAS "T2SC6"

      //Pedidos Pendentes(M�s) - Valor 
      cQry3SC6 := "SELECT SUM(SC6.C6_VALOR) AS VLRPEDPEND" 
      cQry3SC6 += " FROM "+RetSQLName("SC5")+" SC5,"+RetSQLName("SC6")+" SC6"
      cQry3SC6 += " WHERE SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
      cQry3SC6 += "   AND SC6.C6_FILIAL = '"+xFilial("SC6")+"'"
      cQry3SC6 += "   AND SC5.C5_NUM = SC6.C6_NUM"
      cQry3SC6 += "   AND SC6.C6_CF IN "+FormatIn(cCFOPVenda,",")
      cQry3SC6 += "   AND (SUBSTRING(SC6.C6_NOTA,1,3) = '   ' OR (SUBSTRING(SC6.C6_NOTA,1,3) <> '   ' AND (SELECT SUBSTRING(SF2.F2_EMISSAO,1,6)"
      cQry3SC6 += "                                                                                         FROM "+RetSQLName("SF2")+" SF2"
      cQry3SC6 += "                                                                                        WHERE SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      cQry3SC6 += "                                                                                          AND SF2.F2_DOC = SC6.C6_NOTA"
      cQry3SC6 += "                                                                                          AND SF2.F2_SERIE = SC6.C6_SERIE"
      If nMes == 1
         cQry3SC6 += "                                                                                          AND SF2.D_E_L_E_T_ <> '*') > '"+Substr(DTOS(dDataDe),1,6)+"'))" 
         cQry3SC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQry3SC6 += "                                                                                          AND SF2.D_E_L_E_T_ <> '*') > '"+Substr(DTOS(dData02),1,6)+"'))"  
         cQry3SC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQry3SC6 += "                                                                                          AND SF2.D_E_L_E_T_ <> '*') > '"+Substr(DTOS(dData03),1,6)+"'))"    
         cQry3SC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQry3SC6 += " AND SC5.D_E_L_E_T_ <> '*'"
      cQry3SC6 += " AND SC6.D_E_L_E_T_ <> '*'"
      
      TCQuery cQry3SC6 NEW ALIAS "T3SC6"


      //Pedidos Faturados - Quantidade(Meses Anteriores)
      cQuery := "SELECT  SUBSTRING(SD2.D2_EMISSAO,1,6),COUNT(DISTINCT D2_PEDIDO) AS PEDMAFAT FROM " 
      cQuery += RetSqlName('SD2') + " SD2,"+RetSQLName("SB1")+" SB1,"+RetSQLName("SF2")+" SF2,"+RetSQLName("SC5")+" SC5"
      cQuery += " WHERE SD2.D2_FILIAL  = '"+xFilial("SD2")+"' AND  SD2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SB1.B1_FILIAL  = '"+xFilial("SB1")+"' AND  SB1.D_E_L_E_T_ = ' ' "
      cQuery += " AND SF2.F2_FILIAL  = '"+xFilial("SF2")+"' AND  SF2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SC5.C5_FILIAL  = '"+xFilial("SC5")+"' AND  SC5.D_E_L_E_T_ = ' ' "
      cQuery += " AND SD2.D2_COD = SB1.B1_COD"
      //cQuery += " AND SUBSTRING(SB1.B1_GRUPO,1,2) <> '80'" //Faturamento sem Sucatas
      cQuery += " AND SF2.F2_DOC = SD2.D2_DOC"
      cQuery += " AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQuery += " AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQuery += " AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      cQuery += " AND SD2.D2_PEDIDO = SC5.C5_NUM"
      
      
      If nMes == 1
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
         cQuery += " AND  SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dDataDe),1,6)+"'" 
      ElseIf nMes == 2
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
         cQuery += " AND  SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dData02),1,6)+"'" 
      ElseIf nMes == 3
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
         cQuery += " AND  SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dData03),1,6)+"'" 
      EndIf
      
      cQuery += " GROUP BY SUBSTRING(SD2.D2_EMISSAO,1,6) "
      
      MemoWrite("C:\TEMP\PedFatQtdMesAnt"+Alltrim(Str(nMes))+".SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "TXSD2"
      
      //Pedidos Faturados - Valor(Meses Anteriores)
      cQuery := " SELECT SUBSTRING(SD2.D2_EMISSAO,1,6) AS EMISSAO,SUM(D2_TOTAL+CASE WHEN SD2.D2_TIPO <> 'P' THEN D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA ELSE 0 END) AS VLRMAFATPED,SUM(D2_QUANT) AS QTDMAFTPED FROM "
      cQuery += RetSqlName('SD2') + " SD2,"+RetSQLName("SB1")+" SB1,"+RetSQLName("SF2")+" SF2,"+RetSQLName("SC5")+" SC5"
      cQuery += " WHERE SD2.D2_FILIAL  = '"+xFilial("SD2")+"' AND  SD2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SB1.B1_FILIAL  = '"+xFilial("SB1")+"' AND  SB1.D_E_L_E_T_ = ' ' "
      cQuery += " AND SF2.F2_FILIAL  = '"+xFilial("SF2")+"' AND  SF2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SC5.C5_FILIAL  = '"+xFilial("SC5")+"' AND  SC5.D_E_L_E_T_ = ' ' "
      cQuery += " AND SD2.D2_COD = SB1.B1_COD"
      //cQuery += " AND SUBSTRING(SB1.B1_GRUPO,1,2) <> '80'" //Faturamento sem Sucatas
      cQuery += " AND SF2.F2_DOC = SD2.D2_DOC"
      cQuery += " AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQuery += " AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQuery += " AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      cQuery += " AND SD2.D2_PEDIDO = SC5.C5_NUM"
      
      If nMes == 1
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
         cQuery += " AND  SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dDataDe),1,6)+"'" 
      ElseIf nMes == 2
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
         cQuery += " AND  SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dData02),1,6)+"'" 
      ElseIf nMes == 3
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
         cQuery += " AND  SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dData03),1,6)+"'" 
      EndIf
      
      cQuery += " GROUP BY SUBSTRING(SD2.D2_EMISSAO,1,6) "
      cQuery += " ORDER BY SUBSTRING(SD2.D2_EMISSAO,1,6) ASC"
      
      MemoWrite("C:\TEMP\PedFaturaMesAntVlr"+Alltrim(Str(nMes))+".SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "T4SD2"

      //Pedidos Pendentes - Quantidade(Meses Anteriores) 
      cQry4SC6 := "SELECT COUNT(DISTINCT SC6.C6_NUM) AS PEDMAPEND" 
      cQry4SC6 += " FROM "+RetSQLName("SC5")+" SC5,"+RetSQLName("SC6")+" SC6"
      cQry4SC6 += " WHERE SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
      cQry4SC6 += "   AND SC6.C6_FILIAL = '"+xFilial("SC6")+"'"
      cQry4SC6 += "   AND SC5.C5_NUM = SC6.C6_NUM"
      cQry4SC6 += "   AND SC6.C6_CF IN "+FormatIn(cCFOPVenda,",")
      cQry4SC6 += "   AND SUBSTRING(SC5.C5_NOTA,1,3) <> 'XXX'" 
      //Ita - 03/03/2008 - Novo Conceito p/Calcular Quantidade dos Pedidos Pendentes
      //cQry4SC6 += "   AND SUBSTRING(SC6.C6_NOTA,1,3) = '   '"
      cQry4SC6 += "   AND (SUBSTRING(SC6.C6_NOTA,1,3) = '   ' OR (SUBSTRING(SC6.C6_NOTA,1,3) <> '   ' AND SUBSTRING(SC6.C6_NOTA,1,3) <> 'XXX' AND (SELECT SUBSTRING(SF2.F2_EMISSAO,1,6)"
      cQry4SC6 += "                                                                                         FROM "+RetSQLName("SF2")+" SF2"
      cQry4SC6 += "                                                                                        WHERE SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      cQry4SC6 += "                                                                                          AND SF2.F2_DOC = SC6.C6_NOTA"
      cQry4SC6 += "                                                                                          AND SF2.F2_SERIE = SC6.C6_SERIE"
      If nMes == 1
         cQry4SC6 += "                                                                                          AND SF2.D_E_L_E_T_ <> '*') > '"+Substr(DTOS(dDataDe),1,6)+"'))"  
         cQry4SC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQry4SC6 += "                                                                                          AND SF2.D_E_L_E_T_ <> '*') > '"+Substr(DTOS(dData02),1,6)+"'))"   
         cQry4SC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQry4SC6 += "                                                                                          AND SF2.D_E_L_E_T_ <> '*') > '"+Substr(DTOS(dData03),1,6)+"'))"    
         cQry4SC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQry4SC6 += " AND SC5.D_E_L_E_T_ <> '*'"
      cQry4SC6 += " AND SC6.D_E_L_E_T_ <> '*'"
      
      TCQuery cQry4SC6 NEW ALIAS "T4SC6"

      //Pedidos Pendentes - Valor(Meses Anteriores) 
      cQry5SC6 := "SELECT SUM(SC6.C6_VALOR) AS VLRPMAPEND" 
      cQry5SC6 += " FROM "+RetSQLName("SC5")+" SC5,"+RetSQLName("SC6")+" SC6"
      cQry5SC6 += " WHERE SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
      cQry5SC6 += "   AND SC6.C6_FILIAL = '"+xFilial("SC6")+"'"
      cQry5SC6 += "   AND SC5.C5_NUM = SC6.C6_NUM"
      cQry5SC6 += "   AND SC6.C6_CF IN "+FormatIn(cCFOPVenda,",")
      cQry5SC6 += "   AND SUBSTRING(SC5.C5_NOTA,1,3) <> 'XXX'" 
      
      //Ita - 03/03/2008 - Novo Conceito p/Calcular Valor dos Pedidos Pendentes
      //cQry5SC6 += "   AND SUBSTRING(SC6.C6_NOTA,1,3) = '   '"
      cQry5SC6 += "   AND (SUBSTRING(SC6.C6_NOTA,1,3) = '   ' OR (SUBSTRING(SC6.C6_NOTA,1,3) <> '   ' AND SUBSTRING(SC6.C6_NOTA,1,3) <> 'XXX' AND (SELECT SUBSTRING(SF2.F2_EMISSAO,1,6)"
      cQry5SC6 += "                                                                                         FROM "+RetSQLName("SF2")+" SF2"
      cQry5SC6 += "                                                                                        WHERE SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      cQry5SC6 += "                                                                                          AND SF2.F2_DOC = SC6.C6_NOTA"
      cQry5SC6 += "                                                                                          AND SF2.F2_SERIE = SC6.C6_SERIE"
      If nMes == 1
         cQry5SC6 += "                                                                                          AND SF2.D_E_L_E_T_ <> '*') > '"+Substr(DTOS(dDataDe),1,6)+"'))" 
         cQry5SC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQry5SC6 += "                                                                                          AND SF2.D_E_L_E_T_ <> '*') > '"+Substr(DTOS(dData02),1,6)+"'))"    
         cQry5SC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQry5SC6 += "                                                                                          AND SF2.D_E_L_E_T_ <> '*') > '"+Substr(DTOS(dData03),1,6)+"'))"     
         cQry5SC6 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      
      cQry5SC6 += " AND SC5.D_E_L_E_T_ <> '*'"
      cQry5SC6 += " AND SC6.D_E_L_E_T_ <> '*'"
      
      MemoWrite("C:\TEMP\PedPenMesesAnt"+Alltrim(Str(nMes))+".SQL",cQry5SC6)
      
      TCQuery cQry5SC6 NEW ALIAS "T5SC6"

      If nMes == 1
         DbSelectArea("TSC5")
         nPed1MesQt := TSC5->PEDIDOS
         DbCloseArea()
         DbSelectArea("TSC6")
         nPed1VlMes := TSC6->VLRPED
         DbCloseArea()
         DbSelectArea("TSD2")
         nFatPd1QtMes := TSD2->PEDFAT //Pedidos Faturados - Quantidade
         DbCloseArea()
         DbSelectArea("T2SD2") //Pedidos Faturados - Valor
         nFtPd1VlrMes := T2SD2->VLRFATPED
         nQtTon1MVen  := T2SD2->QTDFATPED
         DbCloseArea()
         DbSelectArea("T2SC6")
         nPdPen1QtMes := T2SC6->PEDPEND //Pedidos Pendentes - Quantidade
         DbCloseArea()
         DbSelectArea("T3SC6")
         nPdPen1VlrM := T3SC6->VLRPEDPEND //Pedidos Pendentes - Valor
         DbCloseArea()
         DbSelectArea("TXSD2")
         nFtMAPd1Qt := TXSD2->PEDMAFAT //Pedidos Faturados - Quantidade(Meses Anteriores)
         DbCloseArea()
         DbSelectArea("T4SD2")
         nFtPd1MAVl := T4SD2->VLRMAFATPED //Pedidos Faturados - Valor(Meses Anteriores)
         nQtTon1MAPd:= T4SD2->QTDMAFTPED
         DbCloseArea()
         DbSelectArea("T4SC6")
         nPdPen1MAQt := T4SC6->PEDMAPEND //Pedidos Pendentes - Quantidade(Meses Anteriores)
         DbCloseArea()
         DbSelectArea("T5SC6")
         nPdPMA1VlM := T5SC6->VLRPMAPEND //Pedidos Pendentes - Valor(Meses Anteriores)
         DbCloseArea()
      EndIf

      If nMes == 2
         DbSelectArea("TSC5")
         nPed2MesQt := TSC5->PEDIDOS
         DbCloseArea()
         DbSelectArea("TSC6")
         nPed2VlMes := TSC6->VLRPED
         DbCloseArea()
         DbSelectArea("TSD2")
         nFatPd2QtMes := TSD2->PEDFAT
         DbCloseArea()
         DbSelectArea("T2SD2")
         nFtPd2VlrMes := T2SD2->VLRFATPED
         nQtTon2MVen  := T2SD2->QTDFATPED
         DbCloseArea()
         DbSelectArea("T2SC6")
         nPdPen2QtMes := T2SC6->PEDPEND
         DbCloseArea()
         DbSelectArea("T3SC6")
         nPdPen2VlrM := T3SC6->VLRPEDPEND
         DbCloseArea()
         DbSelectArea("TXSD2")
         nFtMAPd2Qt := TXSD2->PEDMAFAT //Pedidos Faturados - Quantidade(Meses Anteriores)
         DbCloseArea()
         DbSelectArea("T4SD2")
         nFtPd2MAVl := T4SD2->VLRMAFATPED
         nQtTon2MAPd:= T4SD2->QTDMAFTPED
         DbCloseArea()
         DbSelectArea("T4SC6")
         nPdPen2MAQt := T4SC6->PEDMAPEND
         DbCloseArea()
         DbSelectArea("T5SC6")
         nPdPMA2VlM := T5SC6->VLRPMAPEND
         DbCloseArea()
      EndIf

      If nMes == 3
         DbSelectArea("TSC5")
         nPed3MesQt := TSC5->PEDIDOS
         DbCloseArea()
         DbSelectArea("TSC6")
         nPed3VlMes := TSC6->VLRPED
         DbCloseArea()
         DbSelectArea("TSD2")
         nFatPd3QtMes := TSD2->PEDFAT
         DbCloseArea()
         DbSelectArea("T2SD2")
         nFtPd3VlrMes := T2SD2->VLRFATPED
         nQtTon3MVen  := T2SD2->QTDFATPED
         DbCloseArea()
         DbSelectArea("T2SC6")
         nPdPen3QtMes := T2SC6->PEDPEND
         DbCloseArea()
         DbSelectArea("T3SC6")
         nPdPen3VlrM := T3SC6->VLRPEDPEND
         DbCloseArea()
         DbSelectArea("TXSD2")
         nFtMAPd3Qt := TXSD2->PEDMAFAT //Pedidos Faturados - Quantidade(Meses Anteriores)
         DbCloseArea()
         DbSelectArea("T4SD2")
         nFtPd3MAVl := T4SD2->VLRMAFATPED
         nQtTon3MAPd:= T4SD2->QTDMAFTPED
         DbCloseArea()
         DbSelectArea("T4SC6")
         nPdPen3MAQt := T4SC6->PEDMAPEND
         DbCloseArea()
         DbSelectArea("T5SC6")
         nPdPMA3VlM := T5SC6->VLRPMAPEND
         DbCloseArea()
      EndIf   
   
   Next nMes
   
   /********************************
   * FATURAMENTO BRUTO
   *********************************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. Faturamento Bruto, Processo "+Alltrim(Str(nMes))+" / 3")

      cQuery := " SELECT SUBSTRING(SD2.D2_EMISSAO,1,6) AS EMISSAO,SUM(D2_TOTAL+CASE WHEN SD2.D2_TIPO <> 'P' THEN D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA ELSE 0 END) AS TVALBRUTO,SUM(D2_X_CST2) AS CUSTOMP FROM "
      cQuery += RetSqlName('SD2') + " SD2,"+RetSQLName("SB1")+" SB1,"+RetSQLName("SF2")+" SF2"
      cQuery += " WHERE SD2.D2_FILIAL  = '"+xFilial("SD2")+"' AND  SD2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SB1.B1_FILIAL  = '"+xFilial("SB1")+"' AND  SB1.D_E_L_E_T_ = ' ' "
      cQuery += " AND SF2.F2_FILIAL  = '"+xFilial("SF2")+"' AND  SF2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SD2.D2_COD = SB1.B1_COD"
      cQuery += " AND SUBSTRING(SB1.B1_GRUPO,1,2) <> '80'" //Faturamento sem Sucatas
      cQuery += " AND SF2.F2_DOC = SD2.D2_DOC"
      cQuery += " AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQuery += " AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQuery += " AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      
     
      If nMes == 1
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
      ElseIf nMes == 2
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
      ElseIf nMes == 3
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
      EndIf
      
      cQuery += " GROUP BY SUBSTRING(SD2.D2_EMISSAO,1,6) "
      cQuery += " ORDER BY SUBSTRING(SD2.D2_EMISSAO,1,6) ASC"
            
      MemoWrite("C:\TEMP\FatBrutTot"+Alltrim(Str(nMes))+".SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "CFT1"
      
      If nMes == 1
         DbSelectArea("CFT1")
            nFat1MesBru := CFT1->TVALBRUTO
            nCus1MesFat := CFT1->CUSTOMP
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("CFT1")
            nFat2MesBru := CFT1->TVALBRUTO
            nCus2MesFat := CFT1->CUSTOMP
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("CFT1")
            nFat3MesBru := CFT1->TVALBRUTO
            nCus3MesFat := CFT1->CUSTOMP
         DbCloseArea()
      EndIf
      
   Next nMes
   
   aCFTList := {}
   
   ProcRegua(3)
   
   For nMes := 1 To 3  //Faturamento Bruto por Condi��o de Faturamento(Pagamento)
      
      IncProc("Calc. Fat.Bruto p/Cond.Pg, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      //Exemlo do Faturamento Bruto - cQuery := " SELECT SUBSTRING(SD2.D2_EMISSAO,1,6) AS EMISSAO,SUM(D2_TOTAL+CASE WHEN SD2.D2_TIPO <> 'P' THEN D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA ELSE 0 END) AS TVALBRUTO,SUM(D2_X_CST2) AS CUSTOMP FROM "
      //Faturamento po Condi��o, Antes - cQuery := " SELECT SF2.F2_COND, SUM(SD2.D2_TOTAL+SD2.D2_VALIPI+SD2.D2_ICMSRET+SD2.D2_VALFRE+SD2.D2_SEGURO+SD2.D2_DESPESA) AS VALBRUTO FROM "
      cQuery := " SELECT SF2.F2_COND, SUM(D2_TOTAL+CASE WHEN SD2.D2_TIPO <> 'P' THEN D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA ELSE 0 END) AS VALBRUTO FROM "
      cQuery += RetSqlName('SD2') + " SD2,"+RetSQLName("SB1")+" SB1,"+RetSQLName("SF2")+" SF2"
      cQuery += " WHERE SD2.D2_FILIAL  = '"+xFilial("SD2")+"' AND  SD2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SB1.B1_FILIAL  = '"+xFilial("SB1")+"' AND  SB1.D_E_L_E_T_ = ' ' "
      cQuery += " AND SF2.F2_FILIAL  = '"+xFilial("SF2")+"' AND  SF2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SD2.D2_COD = SB1.B1_COD"
      cQuery += " AND SUBSTRING(SB1.B1_GRUPO,1,2) <> '80'" //Faturamento sem Sucatas
      cQuery += " AND SF2.F2_DOC = SD2.D2_DOC"
      cQuery += " AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQuery += " AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQuery += " AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      
      If nMes == 1
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
      ElseIf nMes == 2
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
      ElseIf nMes == 3
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
      EndIf
            
      cQuery += " GROUP BY SF2.F2_COND"
      cQuery += " ORDER BY SF2.F2_COND ASC"
            
      MemoWrite("C:\TEMP\FatBrutCndPg"+Alltrim(Str(nMes))+".SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "CFT"
      
      If nMes == 1
         TCSetField("CFT","VALBRUTO","N",14,02)
      EndIf
      
      DbSelectArea("CFT")
      While !Eof()
         nPos := aScan(aCFTList, {|x| x[1] == CFT->F2_COND})
         If nPos == 0
            Aadd(aCFTList, {CFT->F2_COND,"X",IIF(nMes==1,CFT->VALBRUTO,0),IIF(nMes==2,CFT->VALBRUTO,0),IIF(nMes==3,CFT->VALBRUTO,0)})
         Else
            aCFTList[nPos,(nMes+2)] := CFT->VALBRUTO
         EndIf
         DbSelectArea("CFT")
         DbSkip()
      EndDo
      DbSelectArea("CFT")
      DbCloseArea()
      
   Next nMes
   

   /**************************
   * Devolu��o de Clientes
   *
   **************************/
   /* Ita - 07/02/2008
   ProcRegua(6)
   
   For nMes := 1 To 6
      
      IncProc("Calc. Devolu��es de Clientes, Proc. "+Alltrim(Str(nMes))+" / 6")
      
      If nMes <= 3
         cQrySF3 := " SELECT SUM(SD1.D1_TOTAL) AS VLRDEVOL"
      ElseIf nMes >= 4
         cQrySF3 := " SELECT SUM(SD1.D1_QUANT * SB1.B1_X_CST2) AS CUSDEVOL"
      EndIf
      cQrySF3 += " FROM "+RetSQLName("SD1")+" SD1,"+RetSQLName("SF3")+" SF3,"+RetSQLName("SB1")+" SB1"
      cQrySF3 += " WHERE SF3.F3_FILIAL = '"+xFilial("SF3")+"'"
      cQrySF3 += " AND SD1.D1_FILIAL = '"+xFilial("SD1")+"'"
      cQrySF3 += " AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"
      cQrySF3 += " AND SD1.D1_COD = SB1.B1_COD"
      cQrySF3 += " AND SD1.D1_DOC = SF3.F3_NFISCAL"
      cQrySF3 += " AND SD1.D1_SERIE = SF3.F3_SERIE"
      cQrySF3 += " AND SD1.D1_FORNECE = SF3.F3_CLIEFOR"
      cQrySF3 += " AND SD1.D1_LOJA = SF3.F3_LOJA"
      If nMes == 1 .Or. nMes == 4
         cQrySF3 += " AND SUBSTRING(SD1.D1_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2 .Or. nMes == 5
         cQrySF3 += " AND SUBSTRING(SD1.D1_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3 .Or. nMes == 6
         cQrySF3 += " AND SUBSTRING(SD1.D1_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQrySF3 += " AND SD1.D1_CF IN "+FormatIn(cCFOPDvVd,",")
      cQrySF3 += " AND SF3.D_E_L_E_T_ <> '*'"
      cQrySF3 += " AND SD1.D_E_L_E_T_ <> '*'"
      cQrySF3 += " AND SB1.D_E_L_E_T_ <> '*'"
      
      If nMes == 3
         MemoWrite("DevolucaoVlr.SQL",cQrySF3)
      ElseIf nMes == 6
         MemoWrite("DevolucaoCus.SQL",cQrySF3)
      EndIf
      
      TCQuery cQrySF3 NEW ALIAS "NFDEV"
      
      If nMes == 1
         DbSelectArea("NFDEV")
         nVl1MesDev := NFDEV->VLRDEVOL
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("NFDEV")
         nVl2MesDev := NFDEV->VLRDEVOL
         DbCloseArea()         
      ElseIf nMes == 3
         DbSelectArea("NFDEV")
         nVl3MesDev := NFDEV->VLRDEVOL
         DbCloseArea()         
      ElseIf nMes == 4
         DbSelectArea("NFDEV")
         nCus1MesDev := NFDEV->CUSDEVOL
         DbCloseArea()         
      ElseIf nMes == 5
         DbSelectArea("NFDEV")
         nCus2MesDev := NFDEV->CUSDEVOL
         DbCloseArea()         
      ElseIf nMes == 6
         DbSelectArea("NFDEV")
         nCus3MesDev := NFDEV->CUSDEVOL
         DbCloseArea()         
      EndIf      
      
   Next nMes
   */
   
   /**************************
   * Devolu��o de Clientes
   *
   **************************/   
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. Devolu��es de Clientes, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := "SELECT SUBSTRING(SD1.D1_DTDIGIT,1,6) AS EMISSAO, SUM(SD1.D1_TOTAL+SD1.D1_VALIPI+SD1.D1_DESPESA) AS VLRDEVOL, SUM(SD2.D2_X_CST2) AS CUSDEVOL "
      cQuery += "FROM "+RetSqlName('SD1')+" SD1 LEFT OUTER JOIN "+RetSQLName("SD2")+" SD2 ON"
      cQuery += " SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
      cQuery += " AND SD1.D1_NFORI = SD2.D2_DOC"
      cQuery += " AND SD1.D1_SERIORI = SD2.D2_SERIE"
      cQuery += " AND SD1.D1_COD = SD2.D2_COD"
      cQuery += " AND SD1.D1_ITEMORI = SD2.D2_ITEM"
      cQuery += " AND SD2.D_E_L_E_T_ <> '*' "
      cQuery += " WHERE "
      cQuery += " SD1.D1_FILIAL = '"+xFilial("SD1")+"'  "
      
      If nMes == 1
         cQuery += " AND SUBSTRING(SD1.D1_DTDIGIT,1,6)  = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQuery += " AND SUBSTRING(SD1.D1_DTDIGIT,1,6)  = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQuery += " AND SUBSTRING(SD1.D1_DTDIGIT,1,6)  = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf

      cQuery += " AND SD1.D1_TIPO = 'D' AND "
      cQuery += " SD1.D1_CF IN ('1201','2201','1202','2202') AND "
      cQuery += " SD1.D1_NFORI <> ' ' AND "
      cQuery += " SD1.D_E_L_E_T_ <> '*' "
      
      cQuery += " GROUP BY SUBSTRING(SD1.D1_DTDIGIT,1,6) "
      cQuery += " ORDER BY SUBSTRING(SD1.D1_DTDIGIT,1,6) ASC "
   
      //If nMes == 1
         MemoWrite("C:\TEMP\DevolucaoVlr"+Alltrim(Str(nMes))+".SQL",cQuery)
      //EndIf
      
      TCQuery cQuery NEW ALIAS "NFDEV"
      
      If nMes == 1
         DbSelectArea("NFDEV")
         nVl1MesDev := NFDEV->VLRDEVOL
         nCus1MesDev := NFDEV->CUSDEVOL
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("NFDEV")
         nVl2MesDev := NFDEV->VLRDEVOL
         nCus2MesDev := NFDEV->CUSDEVOL
         DbCloseArea()         
      ElseIf nMes == 3
         DbSelectArea("NFDEV")
         nVl3MesDev := NFDEV->VLRDEVOL
         nCus3MesDev := NFDEV->CUSDEVOL
         DbCloseArea()         
      EndIf      
      
   Next nMes

   /**************************
   * Devolu��o de Clientes
   * Exclusivo Frete
   **************************/   

   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. Devol.Clientes Fretes, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := "SELECT SUBSTRING(SF1.F1_DTDIGIT,1,6) AS EMISSAO, SUM(SF1.F1_FRETE)AS VLRDVFRT "
      cQuery += "FROM "+RetSqlName('SF1')+" SF1" 
      cQuery += " WHERE " 
      cQuery += " SF1.F1_FILIAL = '"+xFilial("SF1")+"'  "
      cQuery += " AND SF1.F1_TIPO = 'D' "
      cQuery += " AND (SELECT COUNT(*)
      cQuery += "        FROM "+RetSqlName('SD1')+" SD1 LEFT OUTER JOIN "+RetSQLName("SD2")+" SD2 ON "
      cQuery += "        SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
      cQuery += "         AND SD1.D1_NFORI = SD2.D2_DOC"
      cQuery += "         AND SD1.D1_SERIORI = SD2.D2_SERIE"
      cQuery += "         AND SD1.D1_COD = SD2.D2_COD"
      cQuery += "         AND SD1.D1_ITEMORI = SD2.D2_ITEM"
      cQuery += "         AND SD2.D_E_L_E_T_ <> '*'"

      cQuery += "       WHERE SD1.D1_FILIAL = '"+xFilial("SD1")+"'  "

      cQuery += "         AND SD1.D1_DOC = SF1.F1_DOC"
      cQuery += "         AND SD1.D1_SERIE = SF1.F1_SERIE"
      cQuery += "         AND SD1.D1_TIPO = 'D' AND "
      cQuery += "         SD1.D1_CF IN ('1201','2201','1202','2202') AND "
      cQuery += "         SD1.D1_NFORI <> ' '"
      If nMes == 1
         cQuery += " AND SUBSTRING(SD1.D1_DTDIGIT,1,6)  = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQuery += " AND SUBSTRING(SD1.D1_DTDIGIT,1,6)  = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQuery += " AND SUBSTRING(SD1.D1_DTDIGIT,1,6)  = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQuery += "         AND SD1.D_E_L_E_T_ <> '*' ) > 0"

   
      If nMes == 1
         cQuery += " AND SUBSTRING(SF1.F1_DTDIGIT,1,6)  = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQuery += " AND SUBSTRING(SF1.F1_DTDIGIT,1,6)  = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQuery += " AND SUBSTRING(SF1.F1_DTDIGIT,1,6)  = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf

      cQuery += " AND SF1.D_E_L_E_T_ <> '*' "
      cQuery += " GROUP BY SUBSTRING(SF1.F1_DTDIGIT,1,6) "
      cQuery += " ORDER BY SUBSTRING(SF1.F1_DTDIGIT,1,6) ASC "
   
      If nMes == 1
         MemoWrite("C:\TEMP\DevolucaoVlr.SQL",cQuery)
      EndIf
      
      TCQuery cQuery NEW ALIAS "NFDEV"
      
      If nMes == 1
         DbSelectArea("NFDEV")
         nVl1MesDev += NFDEV->VLRDVFRT
         //nCus1MesDev := NFDEV->CUSDEVOL
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("NFDEV")
         nVl2MesDev += NFDEV->VLRDVFRT
         //nCus2MesDev := NFDEV->CUSDEVOL
         DbCloseArea()         
      ElseIf nMes == 3
         DbSelectArea("NFDEV")
         nVl3MesDev += NFDEV->VLRDVFRT
         //nCus3MesDev := NFDEV->CUSDEVOL
         DbCloseArea()         
      EndIf      
      
   Next nMes


   /**************************
   *
   * Produ��o
   *
   ***************************/

  
   /*******************************
   * Produ��o por M�s(Quantidade)
   ********************************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. Produ��o, Processo "+Alltrim(Str(nMes))+" / 3")
      
      cQrySD3 := " SELECT SUBSTRING(D3_EMISSAO,1,6) AS EMISSAO, SUM(D3_QUANT) AS PRODUZIDO"
      cQrySD3 += " FROM "+RetSqlName('SD3')+" SD3" //,"+RetSQLName("SH6")+" SH6"
      cQrySD3 += " WHERE SD3.D3_FILIAL = '"+xFilial("SD3")+"'  "
      //cQrySD3 += "   AND SH6.H6_FILIAL = '"+xFilial("SH6")+"'  "
      If nMes == 1
         cQrySD3 += " AND SUBSTRING(SD3.D3_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQrySD3 += " AND SUBSTRING(SD3.D3_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQrySD3 += " AND SUBSTRING(SD3.D3_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      //cQrySD3 += " AND SD3.D3_CF = 'PR0'"
      cQrySD3 += " AND SD3.D3_TM = '300'"
      cQrySD3 += " AND SD3.D3_LOCAL = '01'"
      cQrySD3 += " AND SD3.D3_ESTORNO = ' '"
      
      If SM0->M0_CODFIL == "06"
         cQrySD3 += " AND SD3.D3_GRUPO IN ('40','52','60','70') "
      Else
         cQrySD3 += " AND SD3.D3_GRUPO IN ('60','70') "
      EndIf

      //cQrySD3 += " AND SD3.D3_OP = SH6.H6_OP"
      //cQrySD3 += " AND SD3.D3_COD = SH6.H6_PRODUTO"
      //cQrySD3 += " AND SD3.D3_LOTECTL = SH6.H6_LOTECTL"
      
      cQrySD3 += " AND SD3.D_E_L_E_T_ = ' '  "
      //cQrySD3 += " AND SH6.D_E_L_E_T_ = ' '  "
      cQrySD3 += " GROUP BY SUBSTRING(SD3.D3_EMISSAO,1,6) "
      cQrySD3 += " ORDER BY SUBSTRING(SD3.D3_EMISSAO,1,6) ASC "
      
      TCQuery cQrySD3 NEW ALIAS "PSD3"
      
      If nMes == 1
         DbSelectArea("PSD3")
         nProd1Mes := PSD3->PRODUZIDO
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("PSD3")
         nProd2Mes := PSD3->PRODUZIDO
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("PSD3")
         nProd3Mes := PSD3->PRODUZIDO
         DbCloseArea()
      EndIf
      
   Next nMes

   /*************************
   *
   * Baixas
   *
   *************************/
   // Obtem os registros a serem processados
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Baixas, Processo "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLBAIXAS "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNatProd,",")
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "TSE5"
      
      If nMes == 1
         DbSelectArea("TSE5")
         nVlr1MesBx := TSE5->VLBAIXAS
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("TSE5")
         nVlr2MesBx := TSE5->VLBAIXAS
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("TSE5")
         nVlr3MesBx := TSE5->VLBAIXAS
         DbCloseArea()
      EndIf
      
   Next nMes

   /************************
   *
   * COMISS�O
   *
   ************************/
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Comiss�es, Processo "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLCOMIS "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtComissao,",") //'100203'"
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "TSE5"
      
      If nMes == 1
         DbSelectArea("TSE5")
         nVl1MComis := TSE5->VLCOMIS
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("TSE5")
         nVl2MComis := TSE5->VLCOMIS
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("TSE5")
         nVl3MComis := TSE5->VLCOMIS
         DbCloseArea()
      EndIf
      
   Next nMes

   /************************
   *
   * GASTOS GERAIS
   *
   ************************/
   
   ProcRegua(33)
   
   For nMes := 1 To 33 //De 1 a 3 Gastos Gerais(cGGNatComl), 
                       //4 a 6 Gastos Gerais M�o de Obra.(cNatGGMO)
                       //7 a 9 Despesas Fixas
                      //10 a 12 Despesas Peri�dicas
                      //13 a 15 Rescis�es
                      //16 a 18 Premia��es
                      //19 a 21 Marketing
                      //22 a 24 Despesas co Viagens
                      //25 a 27 Fretes de Vendas
                      //28 a 30 Outras Despesas do Comercial
                      //31 a 33 Outras Naturezas
                      
      IncProc(" Calc.Gastos Gerais Coml. Proc. "+Alltrim(Str(nMes))+" / 33")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLGGCOML "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1 .Or. nMes == 4  .Or. nMes == 7 .Or. nMes == 10 .Or. nMes == 13 .Or. nMes == 16 .Or. nMes == 19 .Or. nMes == 22 .Or. nMes == 25 .Or. nMes == 28 .Or. nMes == 31 
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2 .Or. nMes == 5 .Or. nMes == 8 .Or. nMes == 11 .Or. nMes == 14 .Or. nMes == 17 .Or. nMes == 20 .Or. nMes == 23 .Or. nMes == 26 .Or. nMes == 29 .Or. nMes == 32
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3  .Or. nMes == 6 .Or. nMes == 9 .Or. nMes == 12 .Or. nMes == 15 .Or. nMes == 18 .Or. nMes == 21 .Or. nMes == 24 .Or. nMes == 27 .Or. nMes == 30 .Or. nMes == 33
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
      
      If nMes <= 3
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cGGNatComl,",")
      ElseIf nMes >= 4 .And. nMes <= 6
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNatGGMO,",")
      ElseIf nMes >= 7 .And. nMes <= 9
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cDFNatGG,",")
      ElseIf nMes >= 10 .And. nMes <= 12
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cDPNatGG,",")
      ElseIf nMes >= 13 .And. nMes <= 15
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNResGG,",")
      ElseIf nMes >= 16 .And. nMes <= 18
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtPremGG,",")
      ElseIf nMes >= 19 .And. nMes <= 21
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNatMark,",")
      ElseIf nMes >= 22 .And. nMes <= 24
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtDespV,",")
      ElseIf nMes >= 25 .And. nMes <= 27
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtFrtVend,",")
      ElseIf nMes >= 28 .And. nMes <= 30
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtOutDesp,",")
      ElseIf nMes >= 31 .And. nMes <= 33
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(nOutNatGG,",")
      EndIf
      
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      /*
      If nMes == 1 .Or. nMes == 4
         MemoWrite("C:\TEMP\GastosGer"+Alltrim(Str(nMes))+".SQL",cQuery)
      EndIf
      */
      
      TCQuery cQuery NEW ALIAS "TSE5"
      
      If nMes <= 3
         If nMes == 1
            DbSelectArea("TSE5")
            nVl1MGGComl := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("TSE5")
            nVl2MGGComl := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMEs == 3
            DbSelectArea("TSE5")
            nVl3MGGComl := TSE5->VLGGCOML
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("TSE5")
            nVl1MGGMO := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("TSE5")
            nVl2MGGMO := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 6
            DbSelectArea("TSE5")
            nVl3MGGMO := TSE5->VLGGCOML
            DbCloseArea()
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("TSE5")
            nVl1MDFGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 8
            DbSelectArea("TSE5")
            nVl2MDFGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 9
            DbSelectArea("TSE5")
            nVl3MDFGG := TSE5->VLGGCOML
            DbCloseArea()
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10
            DbSelectArea("TSE5")
            nVl1MDPGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 11
            DbSelectArea("TSE5")
            nVl2MDPGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 12
            DbSelectArea("TSE5")
            nVl3MDPGG := TSE5->VLGGCOML
            DbCloseArea()
         EndIf
      ElseIf nMes >= 13 .And. nMes <= 15
         If nMes == 13
            DbSelectArea("TSE5")
            nVl1MResGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 14
            DbSelectArea("TSE5")
            nVl2MResGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 15
            DbSelectArea("TSE5")
            nVl3MResGG := TSE5->VLGGCOML
            DbCloseArea()
         EndIf
      ElseIf nMes >= 16 .And. nMes <= 18
         If nMes == 16
            DbSelectArea("TSE5")
            nVl1MPrmGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 17
            DbSelectArea("TSE5")
            nVl2MPrmGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 18
            DbSelectArea("TSE5")
            nVl3MPrmGG := TSE5->VLGGCOML
            DbCloseArea()
         EndIf
      ElseIf nMes >= 19 .And. nMes <= 21
         If nMes == 19
            DbSelectArea("TSE5")
            nVl1MMktGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 20
            DbSelectArea("TSE5")
            nVl2MMktGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 21
            DbSelectArea("TSE5")
            nVl3MMktGG := TSE5->VLGGCOML
            DbCloseArea()
         EndIf
      ElseIf nMes >= 22 .And. nMes <= 24
         If nMes == 22
            DbSelectArea("TSE5")
            nVl1MDVgGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 23
            DbSelectArea("TSE5")
            nVl2MDVgGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 24
            DbSelectArea("TSE5")
            nVl3MDVgGG := TSE5->VLGGCOML
            DbCloseArea()
         EndIf
      ElseIf nMes >= 25 .And. nMes <= 27
         If nMes == 25
            DbSelectArea("TSE5")
            nVl1MDFrtGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 26
            DbSelectArea("TSE5")
            nVl2MDFrtGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 27
            DbSelectArea("TSE5")
            nVl3MDFrtGG := TSE5->VLGGCOML
            DbCloseArea()
         EndIf
      ElseIf nMes >= 28 .And. nMes <= 30
         If nMes == 28
            DbSelectArea("TSE5")
            nVl1MOutDGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 29
            DbSelectArea("TSE5")
            nVl2MOutDGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 30
            DbSelectArea("TSE5")
            nVl3MOutDGG := TSE5->VLGGCOML
            DbCloseArea()
         EndIf
      ElseIf nMes >= 31 .And. nMes <= 33
         If nMes == 31
            DbSelectArea("TSE5")
            nVl1MONtGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 32
            DbSelectArea("TSE5")
            nVl2MONtGG := TSE5->VLGGCOML
            DbCloseArea()
         ElseIf nMes == 33
            DbSelectArea("TSE5")
            nVl3MONtGG := TSE5->VLGGCOML
            DbCloseArea()
         EndIf
      EndIf
      
   Next nMes

  
   /************************
   *
   * GASTOS META ORCAMENTO COML (OR�ADO)
   *
   ************************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Gastos Metas/Orc Comercio(Or�ado) Proc. "+Alltrim(Str(nMes))+" / 3")
      
      If nMes == 1
         _xMES := Substr(fNomeMes(Alltrim(StrZero(Month(dDataDe),2))),1,3)
         cQuery := " SELECT SUM(SE7.E7_VAL"+_xMES+"1) AS VLORCMTORC "
      ElseIf nMes == 2
         _xMES := Substr(fNomeMes(Alltrim(StrZero(Month(dData02),2))),1,3)
         cQuery := " SELECT SUM(SE7.E7_VAL"+_xMES+"1) AS VLORCMTORC "
      ElseIf nMes == 3
         _xMES := Substr(fNomeMes(Alltrim(StrZero(Month(dData03),2))),1,3)
         cQuery := " SELECT SUM(SE7.E7_VAL"+_xMES+"1) AS VLORCMTORC "
      EndIf
      
      cQuery += " FROM " + RetSqlName("SE7")+" SE7 "
      
      If nMes == 1
         _xANO := ALLTRIM(STR(YEAR(dDataDe)))
         cQuery += " WHERE SE7.E7_ANO = '"+_xANO+"'" 
      ElseIf nMes == 2 
         _xANO := ALLTRIM(STR(YEAR(dData02)))
         cQuery += " WHERE SE7.E7_ANO = '"+_xANO+"'" 
      ElseIf nMes == 3 
         _xANO := ALLTRIM(STR(YEAR(dData03)))
         cQuery += " WHERE SE7.E7_ANO = '"+_xANO+"'" 
      EndIf
      
      cQuery += "  AND SE7.E7_FILIAL = '"+xFilial("SE7")+"'"
      cQuery += "  AND SUBSTRING(SE7.E7_NATUREZ,1,6) IN "+FormatIn(cNtGMtOrCml,",")
      cQuery += "  AND SE7.D_E_L_E_T_ = ' '"
      
      MemoWrite("C:\TEMP\MtOr�adoComercio"+Alltrim(Str(nMes))+".SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "OMOC"
      

      
      If nMes == 1
         DbSelectArea("OMOC")
         nCom1MMtOrc := OMOC->VLORCMTORC
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("OMOC")
         nCom2MMtOrc := OMOC->VLORCMTORC
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("OMOC")
         nCom3MMtOrc := OMOC->VLORCMTORC
         DbCloseArea()
      EndIf

   Next nMes

   /************************
   *
   * GASTOS META ORCAMENTO COML (REALIZADO)
   *
   ************************/
   ProcRegua(3)
   
   For nMes := 1 To 3 
                      
      IncProc("Gastos Metas/Orc Comercio(Realizado) Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLGMTORC "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1 
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3 
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf

      cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtGMtOrCml,",")
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      /*
      If nMes == 1 .Or. nMes == 4
         MemoWrite("C:\TEMP\GastosGer"+Alltrim(Str(nMes))+".SQL",cQuery)
      EndIf
      */
      
      TCQuery cQuery NEW ALIAS "TSE5"
      
      If nMes == 1
         DbSelectArea("TSE5")
         nGst1MMtOrC := TSE5->VLGMTORC
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("TSE5")
         nGst2MMtOrC := TSE5->VLGMTORC
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("TSE5")
         nGst3MMtOrC := TSE5->VLGMTORC
         DbCloseArea()
      EndIf
         
   Next nMes

   /****************************
   *
   * FATURAMENTO SUCATA
   *
   * Obs: Conforme orienta��o do Sr. Arimat�ia, selecionamentos os produtos Sucata pelo
   *      Grupo 80.
   *
   ****************************/
  
   /********************************
   * FATURAMENTO ESPECIFICO SUCATA
   *********************************/
   aTSUCList := {}
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. Fat.Especifico Sucata, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(D2_TOTAL+CASE WHEN SD2.D2_TIPO <> 'P' THEN D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA ELSE 0 END) AS VLRSUCAT, SUM(D2_X_CST2) AS CUSTOMPSUC FROM "
      cQuery += RetSqlName('SD2') + " SD2,"+RetSQLName("SB1")+" SB1,"+RetSQLName("SF2")+" SF2"
      cQuery += " WHERE SD2.D2_FILIAL  = '"+xFilial("SD2")+"' AND  SD2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SB1.B1_FILIAL  = '"+xFilial("SB1")+"' AND  SB1.D_E_L_E_T_ = ' ' "
      cQuery += " AND SF2.F2_FILIAL  = '"+xFilial("SF2")+"' AND  SF2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SD2.D2_COD = SB1.B1_COD"
      cQuery += " AND SUBSTRING(SB1.B1_GRUPO,1,2) = '80'" //S� Faturamento Sucatas
      cQuery += " AND SF2.F2_DOC = SD2.D2_DOC"
      cQuery += " AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQuery += " AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQuery += " AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      
     
      If nMes == 1
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
      ElseIf nMes == 2
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
      ElseIf nMes == 3
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
      EndIf
      

      cQuery += " GROUP BY SUBSTRING(SD2.D2_EMISSAO,1,6) "
      cQuery += " ORDER BY SUBSTRING(SD2.D2_EMISSAO,1,6) ASC"
      
      If nMes == 1
         MemoWrite("C:\TEMP\FatCusSucata.SQL",cQuery)
      EndIf
      
      TCQuery cQuery NEW ALIAS "TSUC"

      If nMes <= 3
         If nMes == 1
            DbSelectArea("TSUC")
            nVl1SUCMFat := TSUC->VLRSUCAT
            nVlM1CusSUC := TSUC->CUSTOMPSUC
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("TSUC")
            nVl2SUCMFat := TSUC->VLRSUCAT
            nVlM2CusSUC := TSUC->CUSTOMPSUC
            DbCloseArea()
         ElseIf nMes == 3
            DbSelectArea("TSUC")
            nVl3SUCMFat := TSUC->VLRSUCAT
            nVlM3CusSUC := TSUC->CUSTOMPSUC
            DbCloseArea()
         EndIf
      EndIf
      
      
   Next nMes   
   
   /**********************************************************
   *FATURAMENTO ESPECIFICO SUCATA POR CONDI��O DE FATURAMENTO
   *
   **********************************************************/

   aTSUCList := {}
   
   ProcRegua(3)
   
   For nMes := 1 To 3

      IncProc("Calc. Fat.Bruto Sucata p/Cond.Pg, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      //Faturamento de Sucatas - Antes, cQuery := " SELECT SF2.F2_COND,SUM(SD2.D2_TOTAL+SD2.D2_VALIPI+SD2.D2_ICMSRET+SD2.D2_VALFRE+SD2.D2_SEGURO+SD2.D2_DESPESA) AS VLRSUCAT FROM "
      cQuery := " SELECT SF2.F2_COND,SUM(D2_TOTAL+CASE WHEN SD2.D2_TIPO <> 'P' THEN D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA ELSE 0 END) AS VLRSUCAT FROM "
      cQuery += RetSqlName('SD2') + " SD2,"+RetSQLName("SB1")+" SB1,"+RetSQLName("SF2")+" SF2"
      cQuery += " WHERE SD2.D2_FILIAL  = '"+xFilial("SD2")+"' AND  SD2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SB1.B1_FILIAL  = '"+xFilial("SB1")+"' AND  SB1.D_E_L_E_T_ = ' ' "
      cQuery += " AND SF2.F2_FILIAL  = '"+xFilial("SF2")+"' AND  SF2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SD2.D2_COD = SB1.B1_COD"
      cQuery += " AND SUBSTRING(SB1.B1_GRUPO,1,2) = '80'" //Faturamento Especifico Sucatas
      cQuery += " AND SF2.F2_DOC = SD2.D2_DOC"
      cQuery += " AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQuery += " AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQuery += " AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      
      If nMes == 1
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
      ElseIf nMes == 2
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
      ElseIf nMes == 3
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
      EndIf
            
      cQuery += " GROUP BY SF2.F2_COND"
      cQuery += " ORDER BY SF2.F2_COND ASC"
            
      TCQuery cQuery NEW ALIAS "TSUC"
   
      If nMes <= 3
         DbSelectArea("TSUC")
         While !Eof()
            nPos := aScan(aTSUCList, {|x| x[1] == TSUC->F2_COND})
            If nPos == 0
               Aadd(aTSUCList, {TSUC->F2_COND,"X",IIF(nMes==1,TSUC->VLRSUCAT,0),IIF(nMes==2,TSUC->VLRSUCAT,0),IIF(nMes==3,TSUC->VLRSUCAT,0)})
            Else
               aTSUCList[nPos,((nMes)+2)] := TSUC->VLRSUCAT
            EndIf
            DbSelectArea("TSUC")
            DbSkip()
         EndDo
         DbSelectArea("TSUC")
         DbCloseArea()
      EndIf
      
   Next nMes
 
  /***********************
  * Qtd.: SUCATA Vendida
  ***********************/

  ProcRegua(3)
  
  For nMes := 1 To 3
  
     IncProc("Calc. Qtd. Sucatas Vendidas, Proc. "+Alltrim(Str(nMes))+" / 3")
  
     cQry2SD2 := "SELECT SUM(D2_QUANT) AS QTDSUCVEN"  //Pedidos Faturados(M�s) - Valor
     cQry2SD2 += " FROM "+RetSQLName("SD2")+" SD2,"+RetSQLName("SC5")+" SC5,"+RetSQLName("SE4")+" SE4,"+RetSQLName("SF2")+" SF2,"+RetSQLName("SB1")+" SB1"
     cQry2SD2 += " WHERE SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
     cQry2SD2 += "    AND SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
     cQry2SD2 += "    AND SE4.E4_FILIAL = '"+xFilial("SE4")+"'"
     cQry2SD2 += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"      
     cQry2SD2 += "    AND SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
     cQry2SD2 += "    AND SD2.D2_PEDIDO = SC5.C5_NUM"
     cQry2SD2 += "    AND SF2.F2_COND = SE4.E4_CODIGO"
     
     If nMes == 1
        cQry2SD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
        //cQry2SD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
     ElseIf nMes == 2
        cQry2SD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
        //cQry2SD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
     ElseIf nMes == 3
        cQry2SD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
        //cQry2SD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
     EndIf
      
     //Ita - 12/02/2008
     cQry2SD2 += " AND SUBSTRING(SB1.B1_GRUPO,1,2) = '80'" //Faturamento Sucatas
      
     //Filtra CFOPs de Vendas - Conforme Solicita��o de Patr�cia(Auditoria COMAFAL)
     cQry2SD2 += "    AND SF2.F2_DOC = SD2.D2_DOC"
     cQry2SD2 += "    AND SF2.F2_SERIE = SD2.D2_SERIE"
     cQry2SD2 += "    AND SD2.D2_COD = SB1.B1_COD"
     cQry2SD2 += "    AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
     cQry2SD2 += "    AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
     cQry2SD2 += "    AND SD2.D_E_L_E_T_ <> '*'"
     cQry2SD2 += "    AND SF2.D_E_L_E_T_ <> '*'"
     cQry2SD2 += "    AND SE4.D_E_L_E_T_ <> '*'"
     cQry2SD2 += "    AND SB1.D_E_L_E_T_ <> '*'"      
     cQry2SD2 += "    AND SC5.D_E_L_E_T_ <> '*'"      
     
     MemoWrite("C:\TEMP\QtdSucataVen.SQL",cQry2SD2)
     
     TCQuery cQry2SD2 NEW ALIAS "TSUC"
  
    If nMes <= 3
         If nMes == 1
            DbSelectArea("TSUC")
            nQt1MSUCVen := TSUC->QTDSUCVEN
            DbCloseArea()      
         ElseIf nMes == 2
            DbSelectArea("TSUC")
            nQt2MSUCVen := TSUC->QTDSUCVEN
            DbCloseArea()
         ElseIf nMes == 3
            DbSelectArea("TSUC")
            nQt3MSUCVen := TSUC->QTDSUCVEN
            DbCloseArea()
         EndIf
    EndIf
  
  Next nMes
    
EndIf

If lIndustria
   
   /************************
   *
   * Produ��o por M�quina
   *
   ************************/
   
   aMaqPrdM := {}
   
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Produ��o Maq.Ind. Proc. "+Alltrim(Str(nMes))+" / 3")
      
      //cQryPMQ := " SELECT SUBSTRING(SH6.H6_DTAPONT,1,6) AS EMISSAO,SH1.H1_DESCRI AS MAQUINA, SUM(SH6.H6_QTDPROD) AS PRODUZIDO "
      //cQryPMQ := " SELECT SH6.H6_RECURSO AS MAQUINA, SUM(SH6.H6_QTDPROD) AS MQPRODUZIDO "
      //cQryPMQ := " SELECT SH6.H6_RECURSO AS MAQUINA, SUM(SD3.D3_QUANT) AS MQPRODUZIDO "
      cQryPMQ := " SELECT SD3.D3_CC AS MAQUINA, SUM(SD3.D3_QUANT) AS MQPRODUZIDO "
      cQryPMQ += " FROM "
      //cQryPMQ += RetSqlName('SH6') + " SH6,"+RetSQLName("SD3")+" SD3"
      cQryPMQ += RetSQLName("SD3")+" SD3"
      cQryPMQ += " WHERE "
      //cQryPMQ += " SH6.H6_FILIAL = '"+xFilial("SH6")+"'  "
      cQryPMQ += " SD3.D3_FILIAL = '"+xFilial("SD3")+"'  "
      //cQryPMQ += " AND SD3.D3_CF = 'PR0'"
      cQryPMQ += " AND SD3.D3_TM = '300'"
      cQryPMQ += " AND SD3.D3_LOCAL = '01'"
      cQryPMQ += " AND SD3.D3_ESTORNO = ' '"
      
      If SM0->M0_CODFIL == "06"
         cQryPMQ += " AND SD3.D3_GRUPO IN ('40','52','60','70')"
      Else
         cQryPMQ += " AND SD3.D3_GRUPO IN ('60','70')"
      EndIf
      
      //cQryPMQ += " AND SD3.D3_OP = SH6.H6_OP"
      //cQryPMQ += " AND SD3.D3_COD = SH6.H6_PRODUTO"
      //cQryPMQ += " AND SD3.D3_LOTECTL = SH6.H6_LOTECTL"
        
      If nMes == 1
         cQryPMQ += " AND SUBSTRING(D3_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"' " 
      ElseIf nMes == 2
         cQryPMQ += " AND SUBSTRING(D3_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"' " 
      ElseIf nMes == 3
         cQryPMQ += " AND SUBSTRING(D3_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"' "
      EndIf
      
      //cQryPMQ += " AND SH6.D_E_L_E_T_ <> '*'  "
      cQryPMQ += " AND SD3.D_E_L_E_T_ <> '*'  "  
      //cQryPMQ += " AND SH6.H6_QTDPROD > 0 "
      //cQryPMQ += " GROUP BY SH6.H6_RECURSO "
      //cQryPMQ += " ORDER BY SH6.H6_RECURSO ASC "
      cQryPMQ += " GROUP BY SD3.D3_CC "
      cQryPMQ += " ORDER BY SD3.D3_CC ASC "
      
      MemoWrite("C:\TEMP\Maquina"+Alltrim(Str(nMes))+".SQL",cQryPMQ)
      
      TCQuery cQryPMQ NEW ALIAS "TPMQ"

      DbSelectArea("TPMQ")
      While !Eof()
         nPosMQ := aScan(aMaqPrdM, { |x| x[1] == TPMQ->MAQUINA } )

         If nPosMQ == 0
            Aadd(aMaqPrdM, {TPMQ->MAQUINA,IIF(nMes==1,TPMQ->MQPRODUZIDO,0),IIF(nMes==2,TPMQ->MQPRODUZIDO,0),IIF(nMes==3,TPMQ->MQPRODUZIDO,0)})
         Else
            aMaqPrdM[nPosMQ,nMes+1] := TPMQ->MQPRODUZIDO
         EndIf
         
         DbSelectArea("TPMQ")
         DbSkip()
      EndDo
      DbSelectArea("TPMQ")
      DbCloseArea()      
     
      /*
      cQryPMQ += " GROUP BY SUBSTRING(SH6.H6_DTAPONT,1,6),SH1.H1_DESCRI "
      cQryPMQ += " ORDER BY SUBSTRING(SH6.H6_DTAPONT,1,6),SH1.H1_DESCRI ASC "
      */
      
      
   
   Next nMes
   
   /**************************
   *
   * Calculando Mat�ria Prima
   *
   ***************************/
   ProcRegua(3)
   
   For nMes := 1 To 3

      IncProc("Calc. Mat.Prima Ind. Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryMP := " SELECT SUM(SB9.B9_QINI) AS QUANT, SUM(SB9.B9_QINI*SB1.B1_X_CST2) AS CUSTO "
      cQryMP += " FROM "
      cQryMP += RetSqlName('SB9') + " SB9, " 
      cQryMP += RetSqlName('SB1') + " SB1 "
      cQryMP += " WHERE SB9.D_E_L_E_T_ <> '*' AND "
      cQryMP += " SB9.B9_FILIAL = '"+xFilial("SB9")+"' AND "               
      
      If nMes == 1
         cQryMP += " SB9.B9_DATA = '"+DTOS(LastDay(dDataDe))+"' AND " 
      ElseIf nMes == 2
         cQryMP += " SB9.B9_DATA = '"+DTOS(LastDay(dData02))+"' AND " 
      ElseIf nMes == 3
         cQryMP += " SB9.B9_DATA = '"+DTOS(LastDay(dData03))+"' AND " 
      EndIf

      cQryMP += " SUBSTRING(SB9.B9_COD,1,1) <> 'R' AND "
      cQryMP += " SB9.B9_LOCAL IN "+FormatIn(cAlmoxMP,",")+" AND "
      cQryMP += " SB1.B1_FILIAL = '"+xFilial("SB1")+"' AND "               
      cQryMP += " SB1.D_E_L_E_T_ <> '*' AND "
      cQryMP += " SB1.B1_COD = SB9.B9_COD AND " 
      cQryMP += " (SB1.B1_TIPO = 'MP')" // OR (SB1.B1_TIPO <> 'MP' AND SUBSTRING(SB1.B1_DESC,1,4) = 'SLIT'))"
      
      //cQryMP += " GROUP BY SUBSTRING(SB9.B9_DATA,1,6) "
      
      MemoWrite("C:\TEMP\MateriaPrima"+Alltrim(Str(nMes))+".SQL",cQryMP)
      
      TCQuery cQryMP NEW ALIAS "RMP"
      
      If nMes == 1
         TcSetField("RMP","QUANT","N",14,03)
         TcSetField("RMP","CUSTO","N",12,02)
      EndIf
      
      If nMes == 1
         DbSelectArea("RMP")
         nQtdFM1MEst := RMP->QUANT
//         Alert("Peguei o Valor nQtdFM1MEst "+Alltrim(Str(nQtdFM1MEst)))
         nVlrFM1MEst := RMP->CUSTO
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("RMP")
         nQtdFM2MEst := RMP->QUANT
         nVlrFM2MEst := RMP->CUSTO
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("RMP")
         nQtdFM3MEst := RMP->QUANT
         nVlrFM3MEst := RMP->CUSTO
         DbCloseArea()
      EndIf

   Next nMes


   /**************************
   *
   * Calculando Mat�ria Prima(SLITER)
   *
   ***************************/
   ProcRegua(3)
   
   For nMes := 1 To 3

      IncProc("Calc. Mat.Prima(Sliter) Ind. Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryMP := " SELECT SUM(SB9.B9_QINI) AS QUANT, SUM(SB9.B9_QINI*SB1.B1_X_CST2) AS CUSTO "
      cQryMP += " FROM "
      cQryMP += RetSqlName('SB9') + " SB9, " 
      cQryMP += RetSqlName('SB1') + " SB1 "
      cQryMP += " WHERE SB9.D_E_L_E_T_ <> '*' AND "
      cQryMP += " SB9.B9_FILIAL = '"+xFilial("SB9")+"' AND "               
      
      If nMes == 1
         cQryMP += " SB9.B9_DATA = '"+DTOS(LastDay(dDataDe))+"' AND " 
      ElseIf nMes == 2
         cQryMP += " SB9.B9_DATA = '"+DTOS(LastDay(dData02))+"' AND " 
      ElseIf nMes == 3
         cQryMP += " SB9.B9_DATA = '"+DTOS(LastDay(dData03))+"' AND " 
      EndIf

      cQryMP += " SUBSTRING(SB9.B9_COD,1,1) <> 'R' AND "
      cQryMP += " SB9.B9_LOCAL IN "+FormatIn(cAlmoxMP,",")+" AND "
      cQryMP += " SB1.B1_FILIAL = '"+xFilial("SB1")+"' AND "               
      cQryMP += " SB1.D_E_L_E_T_ <> '*' AND "
      cQryMP += " SB1.B1_COD = SB9.B9_COD AND " 
      cQryMP += " (SB1.B1_TIPO <> 'MP' AND SUBSTRING(SB1.B1_DESC,1,4) = 'SLIT')"
      
      //cQryMP += " GROUP BY SUBSTRING(SB9.B9_DATA,1,6) "
      
      MemoWrite("C:\TEMP\MatPrimaSlit"+Alltrim(Str(nMes))+".SQL",cQryMP)
      
      TCQuery cQryMP NEW ALIAS "RMP"
      
      If nMes == 1
         TcSetField("RMP","QUANT","N",14,03)
         TcSetField("RMP","CUSTO","N",12,02)
      EndIf
      
      If nMes == 1
         DbSelectArea("RMP")
         nQt1MPSlit := RMP->QUANT
         nVl1MCSlit := RMP->CUSTO
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("RMP")
         nQt2MPSlit := RMP->QUANT
         nVl2MCSlit := RMP->CUSTO
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("RMP")
         nQt3MPSlit := RMP->QUANT
         nVl3MCSlit := RMP->CUSTO
         DbCloseArea()
      EndIf

   Next nMes
   
   /****************************
   *
   * Calculando Produto Acabado
   *
   ****************************/
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Prod.Acab. Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryPA := " SELECT SUM(SB9.B9_QINI) as QUANTPA, SUM(SB9.B9_QINI*SB1.B1_X_CST2) AS CUSTOPA "
      cQryPA += " FROM "
      cQryPA += RetSqlName('SB9') + " SB9,"+RetSQLName("SB1")+" SB1"
      cQryPA += " WHERE SB9.D_E_L_E_T_ <> '*'  "
      cQryPA += " AND SB1.D_E_L_E_T_ <> '*'  "
      cQryPA += " AND SB9.B9_FILIAL = '"+xFilial("SB9")+"'  "               
      cQryPA += " AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'  "
      cQryPA += " AND SB9.B9_COD = SB1.B1_COD"
      If nMes == 1
         cQryPA += " AND SB9.B9_DATA = '"+DTOS(LastDay(dDataDe))+"' " 
      ElseIf nMes == 2
         cQryPA += " AND SB9.B9_DATA = '"+DTOS(LastDay(dData02))+"' " 
      ElseIf nMes == 3
         cQryPA += " AND SB9.B9_DATA = '"+DTOS(LastDay(dData03))+"' " 
      EndIf
      cQryPA += " AND SB9.B9_LOCAL IN "+FormatIn(cAlmoxPA,",")+" "
      cQryPA += " AND SB1.B1_TIPO = 'PA'"
      //cQryPA += " AND SB1.B1_GRUPO = ''"
      
      /*
      cQryPA += " GROUP BY SUBSTRING(SB9.B9_DATA,1,6) "
      cQryPA += " ORDER BY SUBSTRING(SB9.B9_DATA,1,6) ASC "   
      */
      
      MemoWrite("C:\TEMP\ProdAcabado"+Alltrim(Str(nMes))+".SQL",cQryPA)
      
      TCQuery cQryPA NEW ALIAS "TCPA"
      
      If nMes == 1
         DbSelectArea("TCPA")
         nQt1PAM := TCPA->QUANTPA
         nCs1PAM := TCPA->CUSTOPA
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("TCPA")
         nQt2PAM := TCPA->QUANTPA
         nCs2PAM := TCPA->CUSTOPA
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("TCPA")
         nQt3PAM := TCPA->QUANTPA
         nCs3PAM := TCPA->CUSTOPA
         DbCloseArea()
      EndIf
      
   Next nMes


   /**************************
   * Calculando PRODUTIVIDADE
   **************************/
   ProcRegua(3)
   For nMes := 1 To 3
      
      IncProc("Calc. Produtividade, Processo "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT RD_DATARQ , SUM(RD_HORAS) AS TOTALPRO FROM " 
      cQuery += RetSqlName('SRD') 
      cQuery += " WHERE "
      
      If nMes == 1
         cQuery += " SUBSTRING(RD_DATARQ,1,6) = '" + SUBSTR(DTOS(dDataDe),1,6) +"' "
      ElseIf nMes == 2
         cQuery += " SUBSTRING(RD_DATARQ,1,6) = '" + SUBSTR(DTOS(dData02),1,6) +"' "
      ElseIf nMes == 3
         cQuery += " SUBSTRING(RD_DATARQ,1,6) = '" + SUBSTR(DTOS(dData03),1,6) +"' "
      EndIf
      
      cQuery += " AND RD_CC BETWEEN '12301' AND '12322' " 
      cQuery += " AND RD_PD IN ('002','031','033','034','036') " 
      cQuery += " AND RD_FILIAL = '"+xFilial("SRD")+"' "
      cQuery += " GROUP BY RD_DATARQ " 
      cQuery += " ORDER BY RD_DATARQ ASC" 
      
      MemoWrite("C:\TEMP\Produtividade"+Alltrim(Str(nMes))+".SQL",cQuery)
      
      
      TCQuery cQuery NEW ALIAS "PRVD"
      
      If nMes == 1
         DbSelectArea("PRVD")
         nProvd1M := PRVD->TOTALPRO
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("PRVD")
         nProvd2M := PRVD->TOTALPRO
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("PRVD")
         nProvd3M := PRVD->TOTALPRO
         DbCloseArea()
      EndIf
      
   Next nMes
   
   /***************************
   *
   * Custo de Produ��o
   *
   ***************************/
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Custo Produ��o, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLCUSPROD "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cCUSTPROD,",")
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "TCPR"
      
      If nMes == 1
         DbSelectArea("TCPR")
         nVlr1CusPro := TCPR->VLCUSPROD
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("TCPR")
         nVlr2CusPro := TCPR->VLCUSPROD
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("TCPR")
         nVlr3CusPro := TCPR->VLCUSPROD
         DbCloseArea()
      EndIf
      
   Next nMes

   /***************************
   *
   * Gastos Gerais Ind�stria
   *
   ***************************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Gastos Gerais Ind. Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS GGERAISIND "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cGASTGERIND,",")
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "GGIND"
      
      If nMes == 1
         DbSelectArea("GGIND")
         nVGG1MInd := GGIND->GGERAISIND
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("GGIND")
         nVGG2MInd := GGIND->GGERAISIND
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("GGIND")
         nVGG3MInd := GGIND->GGERAISIND
         DbCloseArea()
      EndIf
      
   Next nMes
    
   /***************************
   *
   * Imobilizado Total      
   *
   ***************************/
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Imobilizado Tot. Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLIMOBTOT "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cImobTotal,",")
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "VIMO"
      
      If nMes == 1
         DbSelectArea("VIMO")
         nVImob1MTt := VIMO->VLIMOBTOT
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("VIMO")
         nVImob2MTt := VIMO->VLIMOBTOT
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("VIMO")
         nVImob3MTt := VIMO->VLIMOBTOT
         DbCloseArea()
      EndIf
      
   Next nMes

   /***************************
   *
   * M�quinas e Motores     
   *
   ***************************/
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Maq. e Motores, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLMAQMOT "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNatMaqMot,",")
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      MemoWrite("C:\TEMP\MaquinaeMotores"+Alltrim(Str(nMes))+".SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "VMQMT"
      
      If nMes == 1
         DbSelectArea("VMQMT")
         nVMqMt1MTt := VMQMT->VLMAQMOT
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("VMQMT")
         nVMqMt2MTt := VMQMT->VLMAQMOT
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("VMQMT")
         nVMqMt3MTt := VMQMT->VLMAQMOT
         DbCloseArea()
      EndIf
      
   Next nMes

   /***************************
   *
   * Importado              
   *
   ***************************/
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Importados, Processo "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLIMPORT "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNatImpInd,",")
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "VLIMP"
      
      If nMes == 1
         DbSelectArea("VLIMP")
         nVlM1ImpInd := VLIMP->VLIMPORT
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("VLIMP")
         nVlM2ImpInd := VLIMP->VLIMPORT
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("VLIMP")
         nVlM3ImpInd := VLIMP->VLIMPORT
         DbCloseArea()
      EndIf
      
   Next nMes

   /***************************
   *
   * Informatica            
   *
   ***************************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Informatica, Processo "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLINFO "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNatInfInd,",")
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "VLRINF"
      
      If nMes == 1
         DbSelectArea("VLRINF")
         nVlInfM1Ind := VLRINF->VLINFO
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("VLRINF")
         nVlInfM2Ind := VLRINF->VLINFO
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("VLRINF")
         nVlInfM3Ind := VLRINF->VLINFO
         DbCloseArea()
      EndIf
      
   Next nMes

   /***************************
   *
   * Moveis e Utens�lios
   *
   ***************************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Mov. Utens�lios, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLMOVUTENS "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cMovUteInd,",")
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "VLMU"
      
      If nMes == 1
         DbSelectArea("VLMU")
         nVlMU1MInd := VLMU->VLMOVUTENS
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("VLMU")
         nVlMU2MInd := VLMU->VLMOVUTENS
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("VLMU")
         nVlMU3MInd := VLMU->VLMOVUTENS
         DbCloseArea()
      EndIf
      
   Next nMes

   /***************************
   *
   * INVESTIMENTOS (Total)
   *
   ***************************/
   
   ProcRegua(21)
   
   For nMes := 1 To 21 //  1 a  3 - Investimentos Total
                      //  4 a  6 - Empresas
                      //  7 a  9 - M�quinas
                      // 10 a 12 - Finame
                      // 13 a 15 - Leasing
                      // 16 a 18 - Benfeitorias
                      // 19 A 21 - ISO 9001 - 2000
   
      IncProc("Calc. Invest. Tot., Proc. "+Alltrim(Str(nMes))+" / 21")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLINVTOTAL "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1 .Or. nMes == 4 .Or. nMes == 7 .Or. nMes == 10 .Or. nMes == 13 .Or. nMes == 16 .Or. nMes == 19
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2 .Or. nMes == 5 .Or. nMes == 8 .Or. nMes == 11 .Or. nMes == 14 .Or. nMes == 17 .Or. nMes == 20
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3 .Or. nMes == 6 .Or. nMes == 9 .Or. nMes == 12 .Or. nMes == 15 .Or. nMes == 18 .Or. nMes == 21
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      If nMes <= 3
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cInvTotInd,",")
      ElseIf nMes >= 4 .And. nMes <= 6
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtEmpresas,",")
      ElseIf nMes >= 7 .And. nMes <= 9
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtMaquinas,",")
      ElseIf nMes >= 10 .And. nMes <= 12
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtFiname,",")
      ElseIf nMes >= 13 .And. nMes <= 15
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtLeasing,",")
      ElseIf nMes >= 16 .And. nMes <= 18
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtBenfeit,",")
      ElseIf nMes >= 19 .And. nMes <= 21
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtISO92,",")
      EndIf
      
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      MemoWrite("C:\TEMP\Investimento"+Alltrim(Str(nMes))+".SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "VITOT"
      
      If nMes <= 3
         If nMes == 1
            DbSelectArea("VITOT")
            nVl1MInvTt := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("VITOT")
            nVl2MInvTt := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMEs == 3
            DbSelectArea("VITOT")
            nVl3MInvTt := VITOT->VLINVTOTAL
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("VITOT")
            nVl1MEmpInv := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("VITOT")
            nVl2MEmpInv := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMEs == 6
            DbSelectArea("VITOT")
            nVl3MEmpInv := VITOT->VLINVTOTAL
            DbCloseArea()
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("VITOT")
            nVlMq1MInv := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMes == 8
            DbSelectArea("VITOT")
            nVlMq2MInv := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMEs == 9
            DbSelectArea("VITOT")
            nVlMq3MInv := VITOT->VLINVTOTAL
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10
            DbSelectArea("VITOT")
            nVlFin1MInv := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMes == 11
            DbSelectArea("VITOT")
            nVlFin2MInv := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMEs == 12
            DbSelectArea("VITOT")
            nVlFin3MInv := VITOT->VLINVTOTAL
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 13 .And. nMes <= 15
         If nMes == 13
            DbSelectArea("VITOT")
            nVlLea1MInv := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMes == 14
            DbSelectArea("VITOT")
            nVlLea2MInv := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMEs == 15
            DbSelectArea("VITOT")
            nVlLea3MInv := VITOT->VLINVTOTAL
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 16 .And. nMes <= 18
         If nMes == 16
            DbSelectArea("VITOT")
            nVlBenf1MI := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMes == 17
            DbSelectArea("VITOT")
            nVlBenf2MI := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMEs == 18
            DbSelectArea("VITOT")
            nVlBenf3MI := VITOT->VLINVTOTAL
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 19 .And. nMes <= 21
         If nMes == 19
            DbSelectArea("VITOT")
            nVlISO1MI := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMes == 20
            DbSelectArea("VITOT")
            nVlISO2MI := VITOT->VLINVTOTAL
            DbCloseArea()
         ElseIf nMEs == 21
            DbSelectArea("VITOT")
            nVlISO3MI := VITOT->VLINVTOTAL
            DbCloseArea()
         EndIf      
      EndIf
      
   Next nMes

   /***************************
   *
   * MATERIA PRIMA (Total) - Ind�stria
   *
   ***************************/
   
   ProcRegua(18)
   
   For nMes := 1 To 18 // 1 a  3 - Materia Prima (Total)
                       // 4 a  6 - M. P. Nacional
                      //  7 a  9 - M. P. Importada
                      // 10 a 12 - Tributos Importa��o
                      // 13 a 15 - Despesas Importa��o
                      // 16 a 18 - Frete Importa��o

   
      IncProc("Calc. Mat.Prima Ind., Proc. "+Alltrim(Str(nMes))+" / 18")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLMPTOTAL "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1 .Or. nMes == 4 .Or. nMes == 7 .Or. nMes == 10 .Or. nMes == 13 .Or. nMes == 16
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2 .Or. nMes == 5 .Or. nMes == 8 .Or. nMes == 11 .Or. nMes == 14 .Or. nMes == 17
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3 .Or. nMes == 6 .Or. nMes == 9 .Or. nMes == 12 .Or. nMes == 15 .Or. nMes == 18
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      If nMes <= 3
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtMPTotal,",")
      ElseIf nMes >= 4 .And. nMes <= 6
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtNacMP,",")
      ElseIf nMes >= 7 .And. nMes <= 9
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtImpMP,",")
      ElseIf nMes >= 10 .And. nMes <= 12
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtTrbImp,",")
      ElseIf nMes >= 13 .And. nMes <= 15
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtDespImp,",")
      ElseIf nMes >= 16 .And. nMes <= 18
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtFrtImp,",")
      EndIf
      
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "VMPTT"
      
      If nMes <= 3
         If nMes == 1
            DbSelectArea("VMPTT")
            nMP1MVlTt := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("VMPTT")
            nMP2MVlTt := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMEs == 3
            DbSelectArea("VMPTT")
            nMP3MVlTt := VMPTT->VLMPTOTAL
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("VMPTT")
            nNacMP1MVl := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("VMPTT")
            nNacMP2MVl := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMEs == 6
            DbSelectArea("VMPTT")
            nNacMP3MVl := VMPTT->VLMPTOTAL
            DbCloseArea()
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("VMPTT")
            nImpMP1MVl := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMes == 8
            DbSelectArea("VMPTT")
            nImpMP2MVl := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMEs == 9
            DbSelectArea("VMPTT")
            nImpMP3MVl := VMPTT->VLMPTOTAL
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10
            DbSelectArea("VMPTT")
            nTrbImp1MMP := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMes == 11
            DbSelectArea("VMPTT")
            nTrbImp2MMP := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMEs == 12
            DbSelectArea("VMPTT")
            nTrbImp3MMP := VMPTT->VLMPTOTAL
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 13 .And. nMes <= 15
         If nMes == 13
            DbSelectArea("VMPTT")
            nDspImp1MMP := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMes == 14
            DbSelectArea("VMPTT")
            nDspImp2MMP := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMEs == 15
            DbSelectArea("VMPTT")
            nDspImp3MMP := VMPTT->VLMPTOTAL
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 16 .And. nMes <= 18
         If nMes == 16
            DbSelectArea("VMPTT")
            nFrtImp1MMP := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMes == 17
            DbSelectArea("VMPTT")
            nFrtImp2MMP := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMEs == 18
            DbSelectArea("VMPTT")
            nFrtImp3MMP := VMPTT->VLMPTOTAL
            DbCloseArea()
         EndIf      
      EndIf
      
   Next nMes

   /***************************
   *
   * OUTRAS NATUREZAS - Ind�stria
   *
   ***************************/
   ProcRegua(51)
   
   For nMes := 1 To 51 // 1 a  3 - INSUMOS
                       // 4 a  6 - M�O DE OBRA
                       // 7 a  9 -  * Desp. Fixa
                       //10 a 12 -  * Desp. Period.
                       //13 a 15 -  * Rescis�es
                       //16 a 18 -  * Premia��es
                      // 19 a 21 - BENEFICIAMENTO
                      // 22 a 24 - MANUTEN��O
                      // 25 a 27 - SERVI�OS TERCEIRIZADOS
                      // 28 a 30 - EPI
                      // 31 a 33 - ENERGIA
                      // 33 a 36 - AGUA
                      // 37 a 39 - ALUGUEL
                      // 40 a 42 - CONSERV. PREDIO
                      // 43 a 45 - DESPESAS DE VIAGENS
                      // 46 a 48 - OUTRAS DESPESAS IND.
                      // 49 a 51 - OUTRAS NATUREZAS

      IncProc("Calc. Outras Naturezas, Proc. "+Alltrim(Str(nMes))+" / 51")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLOUTNTIMP "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1 .Or. nMes == 4 .Or. nMes == 7 .Or. nMes == 10 .Or. nMes == 13 .Or. nMes == 16 .Or. nMes == 19 .Or. nMes == 22 .Or. nMes == 25 .Or. nMes == 28 .Or. nMes == 31 .Or. nMes == 34 .Or. nMes == 37 .Or. nMes == 40 .Or. nMes == 43 .Or. nMes == 46 .Or. nMes == 49
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2 .Or. nMes == 5 .Or. nMes == 8 .Or. nMes == 11 .Or. nMes == 14 .Or. nMes == 17 .Or. nMes == 20 .Or. nMes == 23 .Or. nMes == 26 .Or. nMes == 29 .Or. nMes == 32 .Or. nMes == 35 .Or. nMes == 38 .Or. nMes == 41 .Or. nMes == 44 .Or. nMes == 47 .Or. nMes == 50
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3 .Or. nMes == 6 .Or. nMes == 9 .Or. nMes == 12 .Or. nMes == 15 .Or. nMes == 18 .Or. nMes == 21 .Or. nMes == 24 .Or. nMes == 27 .Or. nMes == 30 .Or. nMes == 33 .Or. nMes == 36 .Or. nMes == 39 .Or. nMes == 42 .Or. nMes == 45 .Or. nMes == 48 .Or. nMes == 51
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      If nMes <= 3
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtInsumos,",")
      ElseIf nMes >= 4 .And. nMes <= 6
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtMaoObInd,",")
      ElseIf nMes >= 7 .And. nMes <= 9
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtMODFInd,",")
      ElseIf nMes >= 10 .And. nMes <= 12
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtMODPInd,",")
      ElseIf nMes >= 13 .And. nMes <= 15
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtRecisInd,",")
      ElseIf nMes >= 16 .And. nMes <= 18
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtPremInd,",")
      ElseIf nMes >= 19 .And. nMes <= 21
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtBenefInd,",")      
      ElseIf nMes >= 22 .And. nMes <= 24
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtManuten,",")
      ElseIf nMes >= 25 .And. nMes <= 27
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtSerTerInd,",")
      ElseIf nMes >= 28 .And. nMes <= 30
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtEPIInd,",")
      ElseIf nMes >= 31 .And. nMes <= 33
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtEnergInd,",")
      ElseIf nMes >= 34 .And. nMes <= 36
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtAguaInd,",")
      ElseIf nMes >= 37 .And. nMes <= 39
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtAlugInd,",")
      ElseIf nMes >= 40 .And. nMes <= 42
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtCoPredInd,",")
      ElseIf nMes >= 43 .And. nMes <= 45
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtDespViaIn,",")
      ElseIf nMes >= 46 .And. nMes <= 48
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtOutDesInd,",")
      ElseIf nMes >= 49 .And. nMes <= 51
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtOutNatInd,",")
      EndIf
      
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      If nMes >= 22 .And. nMes <= 24
         MemoWrite("C:\TEMP\Manutencao"+Alltrim(Str(nMes))+".SQL",cQuery)
      EndIf
      
      If nMes >= 46 .And. nMes <= 48
         MemoWrite("C:\TEMP\OutrasDespInd"+Alltrim(Str(nMes))+".SQL",cQuery)
      EndIf
      
      TCQuery cQuery NEW ALIAS "ONIM"
      

      If nMes <= 3
         If nMes == 1
            DbSelectArea("ONIM")
            nInsVl1MInd := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("ONIM")
            nInsVl2MInd := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 3
            DbSelectArea("ONIM")
            nInsVl3MInd := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("ONIM")
            nMOInd1MVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("ONIM")
            nMOInd2MVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 6
            DbSelectArea("ONIM")
            nMOInd3MVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("ONIM")
            nFixMO1MInd := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 8
            DbSelectArea("ONIM")
            nFixMO2MInd := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 9
            DbSelectArea("ONIM")
            nFixMO3MInd := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10
            DbSelectArea("ONIM")
            nPerMO1MInd := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 11
            DbSelectArea("ONIM")
            nPerMO2MInd := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 12
            DbSelectArea("ONIM")
            nPerMO3MInd := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 13 .And. nMes <= 15
         If nMes == 13
            DbSelectArea("ONIM")
            nRes1MIndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 14
            DbSelectArea("ONIM")
            nRes2MIndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 15
            DbSelectArea("ONIM")
            nRes3MIndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 16 .And. nMes <= 18
         If nMes == 16
            DbSelectArea("ONIM")
            nPrm1IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 17
            DbSelectArea("ONIM")
            nPrm2IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 18
            DbSelectArea("ONIM")
            nPrm3IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 19 .And. nMes <= 21
         If nMes == 19
            DbSelectArea("ONIM")
            nBnf1IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 20
            DbSelectArea("ONIM")
            nBnf2IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 21
            DbSelectArea("ONIM")
            nBnf3IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      //EndIf
      ElseIf nMes >= 22 .And. nMes <= 24
         If nMes == 22
            DbSelectArea("ONIM")
            nMan1IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 23
            DbSelectArea("ONIM")
            nMan2IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 24
            DbSelectArea("ONIM")
            nMan3IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf            
      ElseIf nMes >= 25 .And. nMes <= 27
         If nMes == 25
            DbSelectArea("ONIM")
            nSTc1IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 26
            DbSelectArea("ONIM")
            nSTc2IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 27
            DbSelectArea("ONIM")
            nSTc3IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 28 .And. nMes <= 30
         If nMes == 28
            DbSelectArea("ONIM")
            nEPI1IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 29
            DbSelectArea("ONIM")
            nEPI2IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 30
            DbSelectArea("ONIM")
            nEPI3IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 31 .And. nMes <= 33
         If nMes == 31
            DbSelectArea("ONIM")
            nEner1IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 32
            DbSelectArea("ONIM")
            nEner2IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 33
            DbSelectArea("ONIM")
            nEner3IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 34 .And. nMes <= 36
         If nMes == 34
            DbSelectArea("ONIM")
            nAgua1IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 35
            DbSelectArea("ONIM")
            nAgua2IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 36
            DbSelectArea("ONIM")
            nAgua3IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 37 .And. nMes <= 39
         If nMes == 37
            DbSelectArea("ONIM")
            nAlg1IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 38
            DbSelectArea("ONIM")
            nAlg2IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 39
            DbSelectArea("ONIM")
            nAlg3IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 40 .And. nMes <= 42
         If nMes == 40
            DbSelectArea("ONIM")
            nCPr1IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 41
            DbSelectArea("ONIM")
            nCPr2IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 42
            DbSelectArea("ONIM")
            nCPr3IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 43 .And. nMes <= 45
         If nMes == 43
            DbSelectArea("ONIM")
            nDVg1IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 44
            DbSelectArea("ONIM")
            nDVg2IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 45
            DbSelectArea("ONIM")
            nDVg3IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 46 .And. nMes <= 48
         If nMes == 46
            DbSelectArea("ONIM")
            nODp1IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 47
            DbSelectArea("ONIM")
            nODp2IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 48
            DbSelectArea("ONIM")
            nODp3IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 49 .And. nMes <= 51
         If nMes == 49
            DbSelectArea("ONIM")
            nONt1IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMes == 50
            DbSelectArea("ONIM")
            nONt2IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         ElseIf nMEs == 51
            DbSelectArea("ONIM")
            nONt3IndVl := ONIM->VLOUTNTIMP
            DbCloseArea()
         EndIf      
      EndIf
   Next nMes
   

   /****************************************
   *
   * Gastos Meta Or�amento Ind. (OR�ADO)
   *
   *****************************************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Gastos Metas/Orc.Industria(Or�ado), Proc. "+Alltrim(Str(nMes))+" / 3")
      
      If nMes == 1
         _xMES := Substr(fNomeMes(Alltrim(StrZero(Month(dDataDe),2))),1,3)
         cQuery := " SELECT SUM(SE7.E7_VAL"+_xMES+"1) AS VORCMTORIND "
      ElseIf nMes == 2
         _xMES := Substr(fNomeMes(Alltrim(StrZero(Month(dData02),2))),1,3)
         cQuery := " SELECT SUM(SE7.E7_VAL"+_xMES+"1) AS VORCMTORIND "
      ElseIf nMes == 3
         _xMES := Substr(fNomeMes(Alltrim(StrZero(Month(dData03),2))),1,3)
         cQuery := " SELECT SUM(SE7.E7_VAL"+_xMES+"1) AS VORCMTORIND "
      EndIf
      
      cQuery += " FROM " + RetSqlName("SE7")+" SE7 "
      
      If nMes == 1
         _xANO := ALLTRIM(STR(YEAR(dDataDe)))
         cQuery += " WHERE SE7.E7_ANO = '"+_xANO+"'" 
      ElseIf nMes == 2 
         _xANO := ALLTRIM(STR(YEAR(dData02)))
         cQuery += " WHERE SE7.E7_ANO = '"+_xANO+"'" 
      ElseIf nMes == 3 
         _xANO := ALLTRIM(STR(YEAR(dData03)))
         cQuery += " WHERE SE7.E7_ANO = '"+_xANO+"'" 
      EndIf
      
      cQuery += "  AND SE7.E7_FILIAL = '"+xFilial("SE7")+"'"
      cQuery += "  AND SUBSTRING(SE7.E7_NATUREZ,1,6) IN "+FormatIn(cNtGstMtOrInd,",")
      cQuery += "  AND SE7.D_E_L_E_T_ = ' '"
      
      MemoWrite("C:\TEMP\MtOr�adoIndustria"+Alltrim(Str(nMes))+".SQL",cQuery)
             
      TCQuery cQuery NEW ALIAS "OMOI"
      
      If nMes == 1
         DbSelectArea("OMOI")
         nInd1MMtOrc := OMOI->VORCMTORIND
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("OMOI")
         nInd2MMtOrc := OMOI->VORCMTORIND
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("OMOI")
         nInd3MMtOrc := OMOI->VORCMTORIND
         DbCloseArea()
      EndIf

   Next nMes


   /****************************************
   *
   * Gastos Meta Or�amento Ind. (REALIZADO)
   *
   *****************************************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Gastos Metas/Orc.Industria(Realizado), Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VGSTMTORIND "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1 
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2 
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3 
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtGstMtOrInd,",")
      
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "GMOI"
      
      If nMes == 1
         DbSelectArea("GMOI")
         nGst1MtOrInd := GMOI->VGSTMTORIND
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("GMOI")
         nGst2MtOrInd := GMOI->VGSTMTORIND
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("GMOI")
         nGst3MtOrInd := GMOI->VGSTMTORIND
         DbCloseArea()
      EndIf

   Next nMes
   
EndIf

If lAdmFin
   
   /***********************
   *
   * Pedidos na Trava (Total)
   *
   ***********************/
   
   ProcRegua(6)
   
   For nMes := 1 To 6
   
      IncProc("Calc. Pedidos na Trava, Proc. "+Alltrim(Str(nMes))+" / 6")
      
      cQrySC7 := " SELECT SUM(C7_TOTAL) AS VALOR FROM "
      cQrySC7 += RetSqlName("SC7")
      cQrySC7 += " WHERE D_E_L_E_T_ <> '*' "
      cQrySC7 += " AND C7_FILIAL = '"+xFilial("SC7")+"' "
      cQrySC7 += " AND C7_X_STAT = '1' "
      If nMes <= 3
         If nMes == 1
            cQrySC7 += " AND C7_EMISSAO = '"+Substr(DTOS(dDataDe),1,6)+"'" 
         ElseIf nMes == 2
            cQrySC7 += " AND C7_EMISSAO = '"+Substr(DTOS(dData02),1,6)+"'" 
         ElseIf nMes == 3
            cQrySC7 += " AND C7_EMISSAO = '"+Substr(DTOS(dData03),1,6)+"'" 
         EndIf
      ElseIf nMes >= 4
         If nMes == 4
            cQrySC7 += " AND C7_EMISSAO < '"+Substr(DTOS(dDataDe),1,6)+"'" 
         ElseIf nMes == 5
            cQrySC7 += " AND C7_EMISSAO < '"+Substr(DTOS(dData02),1,6)+"'" 
         ElseIf nMes == 6
            cQrySC7 += " AND C7_EMISSAO < '"+Substr(DTOS(dData03),1,6)+"'" 
         EndIf      
      EndIf
      
      MemoWrite("C:\TEMP\PedidosTrava"+Alltrim(Str(nMes))+".SQL",cQrySC7)
      
      TCQuery cQrySC7 NEW ALIAS "TPTRV"
      
      If nMes <= 3
         If nMes == 1
            DbSelectArea("TPTRV")
            nPdTrv1MVlr := TPTRV->VALOR
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("TPTRV")
            nPdTrv2MVlr := TPTRV->VALOR
            DbCloseArea()
         ElseIf nMes == 3
            DbSelectArea("TPTRV")
            nPdTrv3MVlr := TPTRV->VALOR
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4
         If nMes == 4
            DbSelectArea("TPTRV")
            nPdTrv1AntM := TPTRV->VALOR
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("TPTRV")
            nPdTrv2AntM := TPTRV->VALOR
            DbCloseArea()
         ElseIf nMes == 6
            DbSelectArea("TPTRV")
            nPdTrv3AntM := TPTRV->VALOR
            DbCloseArea()
         EndIf
      EndIf

   Next nMes 

   /***********************
   *
   * T�tulos na Trava (Total)
   *
   ***********************/
   ProcRegua(6)
   
   For nMes := 1 To 6
      
      IncProc("Calc. T�tulos na Trava, Proc. "+Alltrim(Str(nMes))+" / 6")
      
      cQryE2 := " SELECT SUM(E2_VALOR) AS VALOR FROM "
      cQryE2 += RetSqlName("SE2")
      cQryE2 += " WHERE D_E_L_E_T_ <> '*' "                                            

      cQryE2 += "  AND E2_TIPO NOT IN ('NCF','PA ','NDF')"
      cQryE2 += "  AND SUBSTRING(E2_TIPO,3,1) <> '-'"

      If nMes <= 3
         If nMes == 1
            cQryE2 += " AND SUBSTRING(E2_VENCTO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         ElseIf nMes == 2
            cQryE2 += " AND SUBSTRING(E2_VENCTO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
         ElseIf nMes == 3
            cQryE2 += " AND SUBSTRING(E2_VENCTO,1,6) = '"+Substr(DTOS(dData03),1,6)+"' "
         EndIf
      ElseIf nMes >= 4
         If nMes == 4
            cQryE2 += " AND SUBSTRING(E2_VENCTO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         ElseIf nMes == 5
            cQryE2 += " AND SUBSTRING(E2_VENCTO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
         ElseIf nMes == 6
            cQryE2 += " AND SUBSTRING(E2_VENCTO,1,6) = '"+Substr(DTOS(dData03),1,6)+"' "
         EndIf
      EndIf
      cQryE2 += " AND E2_FILIAL = '"+xFilial("SE2")+"' "
      cQryE2 += " AND E2_X_STAT = '1' "
      
      TCQuery cQryE2 NEW ALIAS "TSE2"
      
      If nMes <= 3
         If nMes == 1
            DbSelectArea("TSE2")
            nTit1MTrvVlr := TSE2->VALOR
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("TSE2")
            nTit2MTrvVlr := TSE2->VALOR
            DbCloseArea()
         ElseIf nMes == 3
            DbSelectArea("TSE2")
            nTit3MTrvVlr := TSE2->VALOR
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4
         If nMes == 4
            DbSelectArea("TSE2")
            nAntTit1MTrv := TSE2->VALOR
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("TSE2")
            nAntTit2MTrv := TSE2->VALOR
            DbCloseArea()
         ElseIf nMes == 6
            DbSelectArea("TSE2")
            nAntTit3MTrv := TSE2->VALOR
            DbCloseArea()
         EndIf
      EndIf
      
   Next nMes   

   /***********************
   *
   * T�tulos a Pagar (M�s)   
   *
   ***********************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. T�tulos A Pagar, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryAP := "SELECT SUM(SE2.E2_VALOR) AS TITAPAG"
      cQryAP += " FROM "+RetSQLName("SE2")+" SE2"
      cQryAP += " WHERE SE2.E2_FILIAL = '"+xFilial("SE2")+"'"

      cQryAP += "  AND SE2.E2_TIPO NOT IN ('NCF','PA ','NDF')"
      cQryAP += "  AND SUBSTRING(SE2.E2_TIPO,3,1) <> '-'"

      //25/03/2008 - Altera��o do crit�rio, conforme solicita��o de Patr�cia, passamos a adotar vencimento real ao inv�s de
      //vencimento.
      If nMes == 1
         cQryAP += " AND SUBSTRING(SE2.E2_VENCREA,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQryAP += " AND SUBSTRING(SE2.E2_VENCREA,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQryAP += " AND SUBSTRING(SE2.E2_VENCREA,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQryAP += " AND SE2.D_E_L_E_T_ <> '*'"
      
      MemoWrite("C:\TEMP\TitulosAPagar"+Alltrim(Str(nMes))+".SQL",cQryAP)
      
      TCQuery cQryAP NEW ALIAS "TTPG"
      
      If nMes == 1
         DbSelectArea("TTPG")
         nAPg1MTit := TTPG->TITAPAG
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("TTPG")
         nAPg2MTit := TTPG->TITAPAG
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("TTPG")
         nAPg3MTit := TTPG->TITAPAG
         DbCloseArea()
      EndIf
      
   Next nMes
   

   /***********************
   *
   * Pagamentos em Aberto(M�s) - Utilizando conceito definido por Patr�cia em 27/03/2008   
   *
   ***********************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. Pagamentos em Aberto, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryAP := "SELECT SUM(SE2.E2_SALDO) AS PAGABERTO"
      cQryAP += " FROM "+RetSQLName("SE2")+" SE2"
      cQryAP += " WHERE SE2.E2_FILIAL = '"+xFilial("SE2")+"'"

      cQryAP += "  AND SE2.E2_TIPO NOT IN ('NCF','PA ','NDF')"
      cQryAP += "  AND SUBSTRING(SE2.E2_TIPO,3,1) <> '-'"
      
      cQryAP += "  AND SE2.E2_SALDO > 0 " //Titulos com algum saldo em aberto

      //25/03/2008 - Altera��o do crit�rio, conforme solicita��o de Patr�cia, passamos a adotar vencimento real ao inv�s de
      //vencimento.
      If nMes == 1
         cQryAP += " AND SUBSTRING(SE2.E2_VENCREA,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQryAP += " AND SUBSTRING(SE2.E2_VENCREA,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQryAP += " AND SUBSTRING(SE2.E2_VENCREA,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQryAP += " AND SE2.D_E_L_E_T_ <> '*'"
      
      MemoWrite("C:\TEMP\PagamentosemAberto"+Alltrim(Str(nMes))+".SQL",cQryAP)
      
      TCQuery cQryAP NEW ALIAS "TTPG"
      
      If nMes == 1
         DbSelectArea("TTPG")
         nSld1MPag := TTPG->PAGABERTO
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("TTPG")
         nSld2MPag := TTPG->PAGABERTO
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("TTPG")
         nSld3MPag := TTPG->PAGABERTO
         DbCloseArea()
      EndIf
      
   Next nMes

   /***********************
   *
   * Pagamentos em Aberto(M�s) - Utilizando conceito definido por Patr�cia em 27/03/2008   
   *
   ***********************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc.Pagtos Aberto Meses Ant., Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryAP := "SELECT SUM(SE2.E2_SALDO) AS PAGABMA"
      cQryAP += " FROM "+RetSQLName("SE2")+" SE2"
      cQryAP += " WHERE SE2.E2_FILIAL = '"+xFilial("SE2")+"'"

      cQryAP += "  AND SE2.E2_TIPO NOT IN ('NCF','PA ','NDF')"
      cQryAP += "  AND SUBSTRING(SE2.E2_TIPO,3,1) <> '-'"
      
      cQryAP += "  AND SE2.E2_SALDO > 0 " //Titulos com algum saldo em aberto

      //25/03/2008 - Altera��o do crit�rio, conforme solicita��o de Patr�cia, passamos a adotar vencimento real ao inv�s de
      //vencimento.
      If nMes == 1
         cQryAP += " AND SUBSTRING(SE2.E2_VENCREA,1,6) < '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQryAP += " AND SUBSTRING(SE2.E2_VENCREA,1,6) < '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQryAP += " AND SUBSTRING(SE2.E2_VENCREA,1,6) < '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQryAP += " AND SE2.D_E_L_E_T_ <> '*'"
      
      MemoWrite("C:\TEMP\PagtosAbertoMA"+Alltrim(Str(nMes))+".SQL",cQryAP)
      
      TCQuery cQryAP NEW ALIAS "TTPG"
      
      If nMes == 1
         DbSelectArea("TTPG")
         nSld1MAnt := TTPG->PAGABMA
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("TTPG")
         nSld2MAnt := TTPG->PAGABMA
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("TTPG")
         nSld3MAnt := TTPG->PAGABMA
         DbCloseArea()
      EndIf
      
   Next nMes


   /************************
   *
   * Pagamentos Realizados
   *
   *************************/
   
   /* Ita - 03/03/2008
   
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Pagtos. Realizados, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLPGREALIZ "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.D_E_L_E_T_ <> '*'"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"

      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCF','PA ','NDF')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"

      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "PGREAL"
      
      If nMes == 1
         DbSelectArea("PGREAL")
         nPg1MRealiz := PGREAL->VLPGREALIZ
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("PGREAL")
         nPg2MRealiz := PGREAL->VLPGREALIZ
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("PGREAL")
         nPg3MRealiz := PGREAL->VLPGREALIZ
         DbCloseArea()
      EndIf
      
   Next nMes

   */
   
   /************************
   *
   * COMPESAN��O (Baixas por Compensa��o) - A Pagar
   *
   *************************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Bx.p/Compensa��o, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS BXCOMPENSA "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.D_E_L_E_T_ <> '*'"
      cQuery += "  AND SE5.E5_TIPODOC IN ('CP')"

      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCF','PA ','NDF')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"

      cQuery += "  AND SE5.E5_MOTBX = 'CMP'"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "BXCMP"
      
      If nMes == 1
         DbSelectArea("BXCMP")
         nBx1MComp := BXCMP->BXCOMPENSA
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("BXCMP")
         nBx2MComp := BXCMP->BXCOMPENSA
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("BXCMP")
         nBx3MComp := BXCMP->BXCOMPENSA
         DbCloseArea()
      EndIf
      
   Next nMes

   /************************
   *
   * DEBITO CC 
   *
   *************************/

   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Bx.p/Debito CC, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLDEBCC "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.D_E_L_E_T_ <> '*'"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"

      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCF','PA ','NDF')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"

      cQuery += "  AND SUBSTRING(SE5.E5_BANCO,1,1) <> ' '"
      cQuery += "  AND SE5.E5_MOTBX = 'DEB'"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "DBCC"
      
      If nMes == 1
         DbSelectArea("DBCC")
         nDBT1MCC := DBCC->VLDEBCC
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("DBCC")
         nDBT2MCC := DBCC->VLDEBCC
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("DBCC")
         nDBT3MCC := DBCC->VLDEBCC
         DbCloseArea()
      EndIf
      
   Next nMes
      
   /************************
   *
   * NORMAL - PAGAR
   *
   *************************/

   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Bx. Normal, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLBXNOR "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.D_E_L_E_T_ <> '*'"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"

      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCF','PA ','NDF')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"

      cQuery += "  AND SUBSTRING(SE5.E5_BANCO,1,1) <> ' '"
      cQuery += "  AND SE5.E5_MOTBX = 'NOR'"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "BXNOR"
      
      If nMes == 1
         DbSelectArea("BXNOR")
         nBx1MNorm := BXNOR->VLBXNOR
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("BXNOR")
         nBx2MNorm := BXNOR->VLBXNOR
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("BXNOR")
         nBx3MNorm := BXNOR->VLBXNOR
         DbCloseArea()
      EndIf
      
   Next nMes

   /************************
   *
   * DEVOLU��O A PAGAR
   *
   *************************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Bx.p/Devolu��o, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS BXDEVOL "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.D_E_L_E_T_ <> '*'"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      
      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCF','PA ','NDF')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"
      
      //Five Solutions - 10/07/2008
      //Devolu��es n�o registram bancos - cQuery += "  AND SUBSTRING(SE5.E5_BANCO,1,1) <> ' '"
      cQuery += "  AND SE5.E5_MOTBX = 'DEV'"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "BXDEV"
      
      If nMes == 1
         DbSelectArea("BXDEV")
         nBx1MDevol := BXDEV->BXDEVOL
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("BXDEV")
         nBx2MDevol := BXDEV->BXDEVOL
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("BXDEV")
         nBx3MDevol := BXDEV->BXDEVOL
         DbCloseArea()
      EndIf
      
   Next nMes

   /************************
   *
   * PAGAMENTOS EM ABERTO
   *
   *************************/

   //"Calculando Saldo a Pagar ..."
   //RunSldPg()
   /* Comentado devido a utiliza��o de novo conceito adotado por Patr�cia em 27/03/2008
   ProcRegua(6)
   nCt := 1
   While nCt <= 6
      IncProc("Calc. Saldo A Pagar, Proc. "+Alltrim(Str(nCt))+" / 6")
      // Ita 03/03/2008 - Novo Teste com Procedure 
      If nCt == 1
         nSld1MPag := ItaSldPagar(dDataDe,1,.T.)
      ElseIf nCt == 2
         nSld2MPag := ItaSldPagar(LastDay(dData02),1,.T.) 
      ElseIf nCt == 3
         nSld3MPag := ItaSldPagar(LastDay(dData03),1,.T.) 
      ElseIf nCt == 4
         nSld1MAnt := ItaSldPagar(FirstDay(dDataDe)-1,1,.T.)
      ElseIf nCt == 5
         nSld2MAnt := ItaSldPagar(FirstDay(dData02)-1,1,.T.)
      ElseIf nCt == 6
         nSld3MAnt := ItaSldPagar(FirstDay(dData03)-1,1,.T.)
      EndIf
      
      nCt ++
   EndDo
   */

   /***********************
   *
   * T�tulos a Pagar (Total)   
   *
   ***********************/
   
   ProcRegua(12)
   nPAChamada := 1
   
   For nMes := 1 To 12  // 1 a 3 - A pagar de 1 a 30 dias
                        // 4 a 6 - A pagar de 31 a 60 dias
                        // 7 a 9 - A pagar de 61 a 90 dias
                        //10 a 12 - A pagar a mais de 90 dias
      
      IncProc("Calc. Tit.A Pagar(Fluxo), Proc. "+Alltrim(Str(nMes))+" / 12")
      /*
      cQryAP := "SELECT SUM(SE2.E2_VALOR) AS TITAPAG"
      cQryAP += " FROM "+RetSQLName("SE2")+" SE2"
      cQryAP += " WHERE SE2.E2_FILIAL = '"+xFilial("SE2")+"'"

      cQryAP += "  AND SE2.E2_TIPO NOT IN ('NCF','PA ','NDF')"
      cQryAP += "  AND SUBSTRING(SE2.E2_TIPO,3,1) <> '-'"
            
      If nMes <= 3
         If nMes == 1                                                     
            cQryAP += " AND SE2.E2_VENCTO BETWEEN '"+DTOS((LastDay(dDataDe)+1))+"' AND '"+DTOS((LastDay(dDataDe)+30))+"'"
         ElseIf nMes == 2
            cQryAP += " AND SE2.E2_VENCTO BETWEEN '"+DTOS((LastDay(dData02)+1))+"' AND '"+DTOS((LastDay(dData02)+30))+"'"
         ElseIf nMes == 3
            cQryAP += " AND SE2.E2_VENCTO BETWEEN '"+DTOS((LastDay(dData03)+1))+"' AND '"+DTOS((LastDay(dData03)+30))+"'"
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4                                                     
            cQryAP += " AND SE2.E2_VENCTO BETWEEN '"+DTOS((LastDay(dDataDe)+31))+"' AND '"+DTOS((LastDay(dDataDe)+60))+"'"
         ElseIf nMes == 5
            cQryAP += " AND SE2.E2_VENCTO BETWEEN '"+DTOS((LastDay(dData02)+31))+"' AND '"+DTOS((LastDay(dData02)+60))+"'"
         ElseIf nMes == 6
            cQryAP += " AND SE2.E2_VENCTO BETWEEN '"+DTOS((LastDay(dData03)+31))+"' AND '"+DTOS((LastDay(dData03)+60))+"'"
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7                                                     
            cQryAP += " AND SE2.E2_VENCTO BETWEEN '"+DTOS((LastDay(dDataDe)+61))+"' AND '"+DTOS((LastDay(dDataDe)+90))+"'"
         ElseIf nMes == 8
            cQryAP += " AND SE2.E2_VENCTO BETWEEN '"+DTOS((LastDay(dData02)+61))+"' AND '"+DTOS((LastDay(dData02)+90))+"'"
         ElseIf nMes == 9
            cQryAP += " AND SE2.E2_VENCTO BETWEEN '"+DTOS((LastDay(dData03)+61))+"' AND '"+DTOS((LastDay(dData03)+90))+"'"
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10                                                     
            cQryAP += " AND SE2.E2_VENCTO > '"+DTOS((LastDay(dDataDe)+90))+"'"
         ElseIf nMes == 11
            cQryAP += " AND SE2.E2_VENCTO > '"+DTOS((LastDay(dData02)+90))+"'"
         ElseIf nMes == 12
            cQryAP += " AND SE2.E2_VENCTO > '"+DTOS((LastDay(dData03)+90))+"'"
         EndIf
      EndIf
      
      cQryAP += " AND SE2.D_E_L_E_T_ <> '*'"
      
      TCQuery cQryAP NEW ALIAS "TTPG"
      */
      If nMes <= 3
         If nMes == 1
            
            //DbSelectArea("TTPG")
            //nAPg1M1a30 := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dDataDe)
            dPA2DtParam:= LastDay(dDataDe)+1
            dPA3DtParam:= LastDay(dPA2DtParam)
            nAPg1M1a30 := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         ElseIf nMes == 2
            //DbSelectArea("TTPG")
            //nAPg2M1a30 := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dData02)
            dPA2DtParam:= LastDay(dData02)+1
            dPA3DtParam:= LastDay(dPA2DtParam)
            nAPg2M1a30 := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         ElseIf nMes == 3
            //DbSelectArea("TTPG")
            //nAPg3M1a30 := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dData03)
            dPA2DtParam:= LastDay(dData03)+1
            dPA3DtParam:= LastDay(dPA2DtParam)
            nAPg3M1a30 := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            //DbSelectArea("TTPG")
            //nAPg1M31a60 := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dDataDe)
            dPA2DtParam:= FirstDay(LastDay(dDataDe)+32)
            dPA3DtParam:= LastDay(dPA2DtParam)
            nAPg1M31a60 := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         ElseIf nMes == 5
            //DbSelectArea("TTPG")
            //nAPg2M31a60 := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dData02)
            dPA2DtParam:= FirstDay(LastDay(dData02)+32)
            dPA3DtParam:= LastDay(dPA2DtParam)
            nAPg2M31a60 := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         ElseIf nMes == 6
            //DbSelectArea("TTPG")
            //nAPg3M31a60 := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dData03)
            dPA2DtParam:= FirstDay(LastDay(dData03)+32)
            dPA3DtParam:= LastDay(dPA2DtParam)
            nAPg3M31a60 := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         EndIf
      /* Five Solutions Consultoria
         20 de agosto de 2008
         
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            //DbSelectArea("TTPG")
            //nAPg1M61a90 := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dDataDe)
            dPA2DtParam:= FirstDay(LastDay(dDataDe)+63)
            dPA3DtParam:= LastDay(dPA2DtParam)
            nAPg1M61a90 := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         ElseIf nMes == 8
            //DbSelectArea("TTPG")
            //nAPg2M61a90 := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dData02)
            dPA2DtParam:= FirstDay(LastDay(dData02)+63)
            dPA3DtParam:= LastDay(dPA2DtParam)
            nAPg2M61a90 := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         ElseIf nMes == 9
            //DbSelectArea("TTPG")
            //nAPg3M61a90 := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dData03)
            dPA2DtParam:= FirstDay(LastDay(dData03)+63)
            dPA3DtParam:= LastDay(dPA2DtParam)
            nAPg3M61a90 := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10
            //DbSelectArea("TTPG")
            //nAPg1M90M := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dDataDe)
            dPA2DtParam:= FirstDay(LastDay(dDataDe)+94)
            dPA3DtParam:= CTOD("31/12/49")
            //Alert("Data Base: "+DTOC(dPADtParam)+" Vencto De "+DTOC(dPA2DtParam)+" Vencto Ate "+DTOC(dPA3DtParam))
            nAPg1M90M := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         ElseIf nMes == 11
            //DbSelectArea("TTPG")
            //nAPg2M90M := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dData02)
            dPA2DtParam:= FirstDay(LastDay(dData02)+94)
            dPA3DtParam:= CTOD("31/12/49")
            nAPg2M90M := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         ElseIf nMes == 12
            //DbSelectArea("TTPG")
            //nAPg3M90M := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dData03)
            dPA2DtParam:= FirstDay(LastDay(dData03)+94)
            dPA3DtParam:= CTOD("31/12/49")
            nAPg3M90M := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         EndIf
      EndIf
      */
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            //DbSelectArea("TTPG")
            //nAPg1M61a90 := TTPG->TITAPAG
            //DbCloseArea()
/*
nAcm1M60dd := 0
nAcm2M60dd := 0
nAcm3M60dd := 0
*/
            dPADtParam := LastDay(dDataDe)
            dPA2DtParam:= FirstDay(LastDay(dDataDe)+63)
            dPA3DtParam:= CTOD("31/12/49")
            nAcm1M60dd := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         ElseIf nMes == 8
            //DbSelectArea("TTPG")
            //nAPg2M61a90 := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dData02)
            dPA2DtParam:= FirstDay(LastDay(dData02)+63)
            dPA3DtParam:= CTOD("31/12/49")
            nAcm2M60dd := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         ElseIf nMes == 9
            //DbSelectArea("TTPG")
            //nAPg3M61a90 := TTPG->TITAPAG
            //DbCloseArea()
            dPADtParam := LastDay(dData03)
            dPA2DtParam:= FirstDay(LastDay(dData03)+63)
            dPA3DtParam:= CTOD("31/12/49")
            nAcm3M60dd := SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)
         EndIf
      EndIf
   Next nMes
   
   /***********************
   *
   * T�tulos a Receber (M�s) 
   *
   ***********************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. Tit. A Receber, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryAP := "SELECT SUM(SE1.E1_VALOR) AS TITAREC"
      cQryAP += " FROM "+RetSQLName("SE1")+" SE1"
      cQryAP += " WHERE SE1.E1_FILIAL = '"+xFilial("SE1")+"'"
      cQryAP += "  AND SE1.E1_TIPO NOT IN ('NCC')"
      //25/03/2008 - Altera��o do crit�rio, conforme solicita��o de Patr�cia, passamos a adotar vencimento real ao inv�s de
      //vencimento.

      If nMes == 1
         cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQryAP += " AND SE1.D_E_L_E_T_ <> '*'"
      
      TCQuery cQryAP NEW ALIAS "TREC"
      
      If nMes == 1
         DbSelectArea("TREC")
         nARe1MTit := TREC->TITAREC
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("TREC")
         nARe2MTit := TREC->TITAREC
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("TREC")
         nARe3MTit := TREC->TITAREC
         DbCloseArea()
      EndIf
      
   Next nMes


   /***********************
   *
   * T�tulos a Receber em Aberto(M�s) - Utilizando novo conceito adotado por Patr�cia em 27/03/2008
   *
   ***********************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. Tit. A Receber, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryAP := "SELECT SUM(SE1.E1_SALDO) AS ARECABERTO"
      cQryAP += " FROM "+RetSQLName("SE1")+" SE1"
      cQryAP += " WHERE SE1.E1_FILIAL = '"+xFilial("SE1")+"'"
      cQryAP += "  AND SE1.E1_TIPO NOT IN ('NCC')"
      
      cQryAP += "  AND SE1.E1_SALDO > 0 " //Titulos com algum saldo em aberto.
      
      //25/03/2008 - Altera��o do crit�rio, conforme solicita��o de Patr�cia, passamos a adotar vencimento real ao inv�s de
      //vencimento.

      If nMes == 1
         cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQryAP += " AND SE1.D_E_L_E_T_ <> '*'"
      
      TCQuery cQryAP NEW ALIAS "TREC"
      
      If nMes == 1
         DbSelectArea("TREC")
         nSld1MRec := TREC->ARECABERTO
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("TREC")
         nSld2MRec := TREC->ARECABERTO
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("TREC")
         nSld3MRec := TREC->ARECABERTO
         DbCloseArea()
      EndIf
      
   Next nMes

   /************************
   *
   * T�tulos Recebidos      
   *
   *************************/
   
   /* Ita 03/03/2008 - Novo Conceito p/ Calcular Titulos Recebidos
   
   ProcRegua(3)

   For nMes := 1 To 3
   
      IncProc("Calc. Tit. Recebido, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLREREALIZ "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'R'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.D_E_L_E_T_ <> '*'"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"

      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCC','RA ')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"

      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "REREAL"
      
      If nMes == 1
         DbSelectArea("REREAL")
         nRe1MRealiz := REREAL->VLREREALIZ
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("REREAL")
         nRe2MRealiz := REREAL->VLREREALIZ
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("REREAL")
         nRe3MRealiz := REREAL->VLREREALIZ
         DbCloseArea()
      EndIf
      
   Next nMes

   */
   
   /************************
   *
   * COMPESAN��O (Baixas por Compensa��o) - A Receber
   *
   *************************/

   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Bx.p/Compensa��o(REC), Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS BXCOMPENRE "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'R'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.D_E_L_E_T_ <> '*'"
      cQuery += "  AND SE5.E5_TIPODOC IN ('CP')"

      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCC','RA ')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"

      cQuery += "  AND SE5.E5_MOTBX = 'CMP'"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "BXCMPRE"
      
      If nMes == 1
         DbSelectArea("BXCMPRE")
         nBxRE1MComp := BXCMPRE->BXCOMPENRE
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("BXCMPRE")
         nBxRE2MComp := BXCMPRE->BXCOMPENRE
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("BXCMPRE")
         nBxRE3MComp := BXCMPRE->BXCOMPENRE
         DbCloseArea()
      EndIf
      
   Next nMes

   /************************
   *
   * DEVOLU��O A RECEBER
   *
   *************************/
   
   ProcRegua(3)

   For nMes := 1 To 3
   
      IncProc("Calc.Bx.p/Devolu��o(REC), Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS BXDEVOLRE "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'R'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.D_E_L_E_T_ <> '*'"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"

      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCC','RA ')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"

      cQuery += "  AND SUBSTRING(SE5.E5_BANCO,1,1) <> ' '"
      cQuery += "  AND SE5.E5_MOTBX = 'DEV'"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "BXREDEV"
      
      If nMes == 1
         DbSelectArea("BXREDEV")
         nBx1MDvRE := BXREDEV->BXDEVOLRE
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("BXREDEV")
         nBx2MDvRE := BXREDEV->BXDEVOLRE
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("BXREDEV")
         nBx3MDvRE := BXREDEV->BXDEVOLRE
         DbCloseArea()
      EndIf
      
   Next nMes

   /************************
   *
   * NORMAL - RECEBER
   *
   *************************/

   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Bx. Normal, Processo "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLBXNORRE "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'R'"
      
      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.D_E_L_E_T_ <> '*'"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SUBSTRING(SE5.E5_BANCO,1,1) <> ' '"
      cQuery += "  AND SE5.E5_MOTBX = 'NOR'"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      
      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCC','RA ')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"
      
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "BXRENOR"
      
      If nMes == 1
         DbSelectArea("BXRENOR")
         nBxRE1MNorm := BXRENOR->VLBXNORRE
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("BXRENOR")
         nBxRE2MNorm := BXRENOR->VLBXNORRE
         DbCloseArea()
      ElseIf nMEs == 3
         DbSelectArea("BXRENOR")
         nBxRE3MNorm := BXRENOR->VLBXNORRE
         DbCloseArea()
      EndIf
      
   Next nMes

   /************************
   *
   * TITULOS EM ABERTO (M�s)
   *
   *************************/

   //"Calculando Saldo a Receber ..."
   //RunSldRe()
   /* Comentado devido a utiliza��o de novo conceito adotado por Patr�cia em 27/03/2008
   ProcRegua(3)
   nCt := 1
   WhIle nCt <= 3
      
      IncProc("Calc. Saldo A Receber, Proc. "+Alltrim(Str(nCt))+" / 3")
      
      If nCt == 1
         nSld1MRec := ItaSldReceb(dDataDe,1,.T.,.T.)
      ElseIf nCt == 2
         nSld2MRec := ItaSldReceb(LastDay(dData02),1,.T.,.T.) 
      ElseIf nCt == 3
         nSld3MRec := ItaSldReceb(LastDay(dData03),1,.T.,.T.) 
      EndIf
      
      nCt ++
      
   EndDo
   */
   
   /******************
   *
   * TITULOS VENCIDOS (A RECEBER)
   *
   ******************/
   
   cAnoAtu    := Alltrim(Str((YEAR(dDataDe))))
   cAno1Menos := Alltrim(Str((YEAR(dDataDe)-1)))
   cAno2Menos := Alltrim(Str((YEAR(dDataDe)-2)))
   cAnoAntes  := Alltrim(Str((YEAR(dDataDe)-3)))
   
   nChamada   := 2
   ProcRegua(15)
   
   
   
   For nMes := 1 To 15
      
      IncProc("Calc. T�t. Vencidos, Proc. "+Alltrim(Str(nMes))+" / 15")
      
      /*
      cQryTV := " SELECT SUM(SE1.E1_SALDO) AS SLDVENCID"
      cQryTV += " FROM "+RetSQLName("SE1")+" SE1"
      cQryTV += " WHERE SE1.E1_FILIAL = '"+(xFilial("SE1"))+"'"
      If nMes == 1 .Or. nMes == 4 .Or. nMes == 7 .Or. nMes == 10 .Or. nMes == 13
         cQryTV += "  AND SE1.E1_VENCTO <= '"+Substr(DTOS(FirstDay(dDataDe)-1),1,6)+"'"
      ElseIf nMes == 2 .Or. nMes == 5 .Or. nMes == 8 .Or. nMes == 11 .Or. nMes == 14
         cQryTV += "  AND SE1.E1_VENCTO <= '"+Substr(DTOS(FirstDay(dData02)-1),1,6)+"'"
      ElseIf nMes == 3 .Or. nMes == 6 .Or. nMes == 9 .Or. nMes == 12 .Or. nMes == 15
         cQryTV += "  AND SE1.E1_VENCTO <= '"+Substr(DTOS(FirstDay(dData03)-1),1,6)+"'"
      EndIf
      
      If nMes >= 4 .And. nMes <= 6
         cQryTV += "  AND SUBSTRING(SE1.E1_VENCTO,1,4) = '"+cAnoAtu+"'"
      ElseIf nMes >= 7 .And. nMes <= 9
         cQryTV += "  AND SUBSTRING(SE1.E1_VENCTO,1,4) = '"+cAno1Menos+"'"
      ElseIf nMes >= 10 .And. nMes <= 12
         cQryTV += "  AND SUBSTRING(SE1.E1_VENCTO,1,4) = '"+cAno2Menos+"'"
      ElseIf nMes >= 13 .And. nMes <= 15
            cQryTV += "  AND SUBSTRING(SE1.E1_VENCTO,1,4) <= '"+cAnoAntes+"'"
      EndIf
      
      cQryTV += "  AND SE1.E1_SALDO > 0"
      cQryTV += "  AND SE1.E1_TIPO NOT IN ('NCC','RA ')"
      cQryTV += "  AND SUBSTRING(SE1.E1_TIPO,3,1) <> '-'"
      cQryTV += "  AND SE1.D_E_L_E_T_ <> '*'"
      
      MemoWrite("C:\TEMP\TitulosVencidos"+Alltrim(Str(nMes))+".SQL",cQryTV)
      
      TCQuery cQryTV NEW ALIAS "TTVENC"
      */
      
      If nMes <= 3
         /*
         If nMes == 1
            DbSelectArea("TTVENC")
            nTt1MVc := TTVENC->SLDVENCID
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("TTVENC")
            nTt2MVc := TTVENC->SLDVENCID
            DbCloseArea()
         ElseIf nMes == 3
            DbSelectArea("TTVENC")
            nTt3MVc := TTVENC->SLDVENCID
            DbCloseArea()
         EndIf
         */
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            //DbSelectArea("TTVENC")
            //nAAtTt1MVc := TTVENC->SLDVENCID
            //DbCloseArea()
            dDtParam   := LastDay(dDataDe) 
            d2DtParam  := CTOD("01/01/"+ALLTRIM(STR(YEAR(dDataDe))))
            d3DtParam  := LastDay(dDataDe)
            //Alert("nMes "+Alltrim(Str(nMes))+" dDtParam "+DTOC(dDtParam)+" d2DtParam "+DTOC(d2DtParam)+" d3DtParam "+DTOC(d3DtParam))
            nAAtTt1MVc := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
            //Alert("nAAtTt1MVc "+Alltrim(Str(nAAtTt1MVc)))
         ElseIf nMes == 5
            //DbSelectArea("TTVENC")
            //nAAtTt2MVc := TTVENC->SLDVENCID
            //DbCloseArea()
            dDtParam   := LastDay(dData02) 
            d2DtParam  := CTOD("01/01/"+ALLTRIM(STR(YEAR(dData02))))
            d3DtParam  := LastDay(dData02)
            //Alert("nMes "+Alltrim(Str(nMes))+" dDtParam "+DTOC(dDtParam)+" d2DtParam "+DTOC(d2DtParam)+" d3DtParam "+DTOC(d3DtParam))
            nAAtTt2MVc := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
            //Alert("nAAtTt2MVc "+Alltrim(Str(nAAtTt2MVc)))
         ElseIf nMes == 6
            //DbSelectArea("TTVENC")
            //nAAtTt3MVc := TTVENC->SLDVENCID
            //DbCloseArea()
            dDtParam   := LastDay(dData03) 
            d2DtParam  := CTOD("01/01/"+ALLTRIM(STR(YEAR(dData03))))
            d3DtParam  := LastDay(dData03)
            //Alert("nMes "+Alltrim(Str(nMes))+" dDtParam "+DTOC(dDtParam)+" d2DtParam "+DTOC(d2DtParam)+" d3DtParam "+DTOC(d3DtParam))
            nAAtTt3MVc := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
            //Alert("nAAtTt3MVc "+Alltrim(Str(nAAtTt3MVc)))
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            //DbSelectArea("TTVENC")
            //nA1ATt1MVc := TTVENC->SLDVENCID
            //DbCloseArea()
            dDtParam   := LastDay(dDataDe) 
            d2DtParam  := CTOD("01/01/"+ALLTRIM(STR(YEAR(dDataDe)-1)))
            d3DtParam  := CTOD("31/12/"+ALLTRIM(STR(YEAR(dDataDe)-1)))
            //Alert("nMes "+Alltrim(Str(nMes))+" dDtParam "+DTOC(dDtParam)+" d2DtParam "+DTOC(d2DtParam)+" d3DtParam "+DTOC(d3DtParam))
            nA1ATt1MVc := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
            //Alert("nA1ATt1MVc "+Alltrim(Str(nA1ATt1MVc)))
         ElseIf nMes == 8
            //DbSelectArea("TTVENC")
            //nA1ATt2MVc := TTVENC->SLDVENCID
            //DbCloseArea()
            dDtParam   := LastDay(dData02) 
            d2DtParam  := CTOD("01/01/"+ALLTRIM(STR(YEAR(dData02)-1)))
            d3DtParam  := CTOD("31/12/"+ALLTRIM(STR(YEAR(dData02)-1)))
            //Alert("nMes "+Alltrim(Str(nMes))+" dDtParam "+DTOC(dDtParam)+" d2DtParam "+DTOC(d2DtParam)+" d3DtParam "+DTOC(d3DtParam))
            nA1ATt2MVc := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
            //Alert("nA1ATt2MVc "+Alltrim(Str(nA1ATt2MVc)))
         ElseIf nMes == 9
            //DbSelectArea("TTVENC")
            //nA1ATt3MVc := TTVENC->SLDVENCID
            //DbCloseArea()
            dDtParam   := LastDay(dData03) 
            d2DtParam  := CTOD("01/01/"+ALLTRIM(STR(YEAR(dData03)-1)))
            d3DtParam  := CTOD("31/12/"+ALLTRIM(STR(YEAR(dData03)-1)))
            //Alert("nMes "+Alltrim(Str(nMes))+" dDtParam "+DTOC(dDtParam)+" d2DtParam "+DTOC(d2DtParam)+" d3DtParam "+DTOC(d3DtParam))
            nA1ATt3MVc := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
            //Alert("nA1ATt3MVc "+Alltrim(Str(nA1ATt3MVc)))
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10
            //DbSelectArea("TTVENC")
            //nA2ATt1MVc := TTVENC->SLDVENCID
            //DbCloseArea()
            dDtParam   := LastDay(dDataDe) 
            d2DtParam  := CTOD("01/01/"+ALLTRIM(STR(YEAR(dDataDe)-2)))
            d3DtParam  := CTOD("31/12/"+ALLTRIM(STR(YEAR(dDataDe)-2)))
            //Alert("nMes "+Alltrim(Str(nMes))+" dDtParam "+DTOC(dDtParam)+" d2DtParam "+DTOC(d2DtParam)+" d3DtParam "+DTOC(d3DtParam))
            nA2ATt1MVc := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         ElseIf nMes == 11
            //DbSelectArea("TTVENC")
            //nA2ATt2MVc := TTVENC->SLDVENCID
            //DbCloseArea()
            dDtParam   := LastDay(dData02) 
            d2DtParam  := CTOD("01/01/"+ALLTRIM(STR(YEAR(dData02)-2)))
            d3DtParam  := CTOD("31/12/"+ALLTRIM(STR(YEAR(dData02)-2)))
            //Alert("nMes "+Alltrim(Str(nMes))+" dDtParam "+DTOC(dDtParam)+" d2DtParam "+DTOC(d2DtParam)+" d3DtParam "+DTOC(d3DtParam))
            nA2ATt2MVc := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         ElseIf nMes == 12
            //DbSelectArea("TTVENC")
            //nA2ATt3MVc := TTVENC->SLDVENCID
            //DbCloseArea()
            dDtParam   := LastDay(dData03) 
            d2DtParam  := CTOD("01/01/"+ALLTRIM(STR(YEAR(dData03)-2)))
            d3DtParam  := CTOD("31/12/"+ALLTRIM(STR(YEAR(dData03)-2)))
            //Alert("nMes "+Alltrim(Str(nMes))+" dDtParam "+DTOC(dDtParam)+" d2DtParam "+DTOC(d2DtParam)+" d3DtParam "+DTOC(d3DtParam))
            nA2ATt3MVc := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         EndIf
      ElseIf nMes >= 13 .And. nMes <= 15
         If nMes == 13
            //DbSelectArea("TTVENC")
            //nAAntTt1MVc := TTVENC->SLDVENCID
            //DbCloseArea()
            dDtParam   := LastDay(dDataDe) 
            d2DtParam  := CTOD("01/01/90")
            d3DtParam  := CTOD("31/12/"+ALLTRIM(STR(YEAR(dDataDe)-3)))
            //Alert("nMes "+Alltrim(Str(nMes))+" dDtParam "+DTOC(dDtParam)+" d2DtParam "+DTOC(d2DtParam)+" d3DtParam "+DTOC(d3DtParam))
            nAAntTt1MVc := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         ElseIf nMes == 14
            //DbSelectArea("TTVENC")
            //nAAntTt2MVc := TTVENC->SLDVENCID
            //DbCloseArea()
            dDtParam   := LastDay(dData02) 
            d2DtParam  := CTOD("01/01/90")
            d3DtParam  := CTOD("31/12/"+ALLTRIM(STR(YEAR(dData02)-3)))
            //Alert("nMes "+Alltrim(Str(nMes))+" dDtParam "+DTOC(dDtParam)+" d2DtParam "+DTOC(d2DtParam)+" d3DtParam "+DTOC(d3DtParam))
            nAAntTt2MVc := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         ElseIf nMes == 15
            //DbSelectArea("TTVENC")
            //nAAntTt3MVc := TTVENC->SLDVENCID
            //DbCloseArea()
            dDtParam   := LastDay(dData03) 
            d2DtParam  := CTOD("01/01/90")
            d3DtParam  := CTOD("31/12/"+ALLTRIM(STR(YEAR(dData03)-3)))
            //Alert("nMes "+Alltrim(Str(nMes))+" dDtParam "+DTOC(dDtParam)+" d2DtParam "+DTOC(d2DtParam)+" d3DtParam "+DTOC(d3DtParam))
            nAAntTt3MVc := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         EndIf
      EndIf
       
   Next nMes 
   
   /***********************
   *
   * T�tulos a Receber (Fluxo)
   *
   ***********************/
   
   nChamada   := 1
    
   ProcRegua(12)
   
   For nMes := 1 To 12  // 1 a 3 - A Receber de 1 a 30 dias
                        // 4 a 6 - A Receber de 31 a 60 dias
                        // 7 a 9 - A Receber de 61 a 90 dias
                        //10 a 12 - A Receber a mais de 90 dias
      
      IncProc("Calc. Tit. A Receber(Fluxo), Proc. "+Alltrim(Str(nMes))+" / 12")
      
      /*
      cQryAP := "SELECT SUM(SE1.E1_VALOR) AS TFLXAREC"
      cQryAP += " FROM "+RetSQLName("SE1")+" SE1"
      cQryAP += " WHERE SE1.E1_FILIAL = '"+xFilial("SE1")+"'"

      cQryAP += "  AND SE1.E1_TIPO NOT IN ('NCC','RA ')"
      cQryAP += "  AND SUBSTRING(SE1.E1_TIPO,3,1) <> '-'"
            
      If nMes <= 3
         If nMes == 1                                                     
            cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS((LastDay(dDataDe)+1)),1,6)+"'"
         ElseIf nMes == 2
            cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS((LastDay(dData02)+1)),1,6)+"'"
         ElseIf nMes == 3
            //cQryAP += " AND SE1.E1_VENCTO BETWEEN '"+DTOS((LastDay(dData03)+1))+"' AND '"+DTOS((LastDay(dData03)+30))+"'"
            cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS((LastDay(dData03)+1)),1,6)+"'"
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4                                                     
            cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS((LastDay(dDataDe)+31)),1,6)+"'"
         ElseIf nMes == 5
            cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS((LastDay(dData02)+31)),1,6)+"'"
         ElseIf nMes == 6
            //cQryAP += " AND SE1.E1_VENCTO BETWEEN '"+DTOS((LastDay(dData03)+31))+"' AND '"+DTOS((LastDay(dData03)+60))+"'"
            cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS((LastDay(dData03)+31)),1,6)+"'"
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7                                                     
            cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS((LastDay(dDataDe)+61)),1,6)+"'"
         ElseIf nMes == 8
            cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS((LastDay(dData02)+61)),1,6)+"'"
         ElseIf nMes == 9
            //cQryAP += " AND SE1.E1_VENCTO BETWEEN '"+DTOS((LastDay(dData03)+61))+"' AND '"+DTOS((LastDay(dData03)+90))+"'"
            cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) = '"+Substr(DTOS((LastDay(dData03)+61)),1,6)+"'"
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10                                                     
            cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) > '"+Substr(DTOS((LastDay(dDataDe)+90)),1,6)+"'"
         ElseIf nMes == 11
            cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) > '"+Substr(DTOS((LastDay(dData02)+90)),1,6)+"'"
         ElseIf nMes == 12
            //cQryAP += " AND SE1.E1_VENCTO > '"+DTOS((LastDay(dData03)+90))+"'"
            cQryAP += " AND SUBSTRING(SE1.E1_VENCREA,1,6) > '"+Substr(DTOS((LastDay(dData03)+90)),1,6)+"'"
         EndIf
      EndIf
      
      cQryAP += " AND SE1.D_E_L_E_T_ <> '*'"
      
      MemoWrite("C:\TEMP\TitAReceberFluxo"+Alltrim(Str(nMes))+".SQL",cQryAP)
      
      TCQuery cQryAP NEW ALIAS "REFLX"
      */
      If nMes <= 3
         If nMes == 1
            //DbSelectArea("REFLX")
            //nARe1M1a30 := REFLX->TFLXAREC
            //DbCloseArea()
            dMesAF := LastDay(dDataDe)+1
            //Alert("Chamando Fun��o  SaldoSR(Data Base= "+DTOC(LastDay(dDataDe))+" Vencto de= "+DTOC(dMesAF)+" Vencto Ate= "+DTOC(LastDay(dMesAF)))
            dDtParam  := LastDay(dDataDe)
            d2DtParam := dMesAF
            d3DtParam := LastDay(dMesAF)
            //Alert("nMes "+Alltrim(Str(nMes))+" Paramentos enviados a SaldoSR, DataBase "+DTOC(dDtParam)+" Vencto De "+dtoc(d2DtParam)+" Vencto Ate "+DTOC(d3DtParam)) 
            nARe1M1a30 := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
            //Alert("A Receber 1 a 30 dias "+Alltrim(Str(nARe1M1a30))+" Data em "+DTOC(FirstDay(dDataDe)-1))
         ElseIf nMes == 2
            //DbSelectArea("REFLX")
            //nARe2M1a30 := REFLX->TFLXAREC
            //DbCloseArea()
            d2MesAF := LastDay(dData02)+1
            dDtParam  := LastDay(dData02)
            d2DtParam := d2MesAF
            d3DtParam := LastDay(d2MesAF)
            
            nARe2M1a30 := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         ElseIf nMes == 3
            //DbSelectArea("REFLX")
            //nARe3M1a30 := REFLX->TFLXAREC
            //DbCloseArea()
            d3MesAF := LastDay(dData03)+1
            dDtParam  := LastDay(dData03)
            d2DtParam := d3MesAF
            d3DtParam := LastDay(d3MesAF)           
            nARe3M1a30 := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            //DbSelectArea("REFLX")
            //nARe1M31a60 := REFLX->TFLXAREC
            //DbCloseArea()
            d4TMPDt := LastDay(dDataDe)+32
            //Alert("d4TMPDt "+DTOC(d4TMPDt))
            d4MesAF := LastDay(d4TMPDt)
            dDtParam  := LastDay(dDataDe)
            d2DtParam := FirstDay(d4MesAF)
            d3DtParam := LastDay(d4MesAF)          
            //Alert("nMes "+Alltrim(Str(nMes))+" Paramentos enviados a SaldoSR, DataBase "+DTOC(dDtParam)+" Vencto De "+dtoc(d2DtParam)+" Vencto Ate "+DTOC(d3DtParam)) 
            nARe1M31a60 := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         ElseIf nMes == 5
            //DbSelectArea("REFLX")
            //nARe2M31a60 := REFLX->TFLXAREC
            //DbCloseArea()
            d5TMPDt := LastDay(dData02)+32 
            d5MesAF := LastDay(d5TMPDt)
            dDtParam  := LastDay(dData02)
            d2DtParam := FirstDay(d5MesAF)
            d3DtParam := LastDay(d5MesAF)            
            nARe2M31a60 := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         ElseIf nMes == 6
            //DbSelectArea("REFLX")
            //nARe3M31a60 := REFLX->TFLXAREC
            //DbCloseArea()
            d6TMPDt := LastDay(dData03)+32  
            d6MesAF := LastDay(d6TMPDt)
            dDtParam  := LastDay(dData03)
            d2DtParam := FirstDay(d6MesAF)
            d3DtParam := LastDay(d6MesAF)             
            nARe3M31a60 := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         EndIf
/* Five Solutions Consultoria
   20 de agosto de 2008
            
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            //DbSelectArea("REFLX")
            //nARe1M61a90 := REFLX->TFLXAREC
            //DbCloseArea()
            d7TMPDt := LastDay(dDataDe)+63  
            d7MesAF := LastDay(d7TMPDt)
            dDtParam  := LastDay(dDataDe)
            d2DtParam := FirstDay(d7MesAF)
            d3DtParam := LastDay(d7MesAF)  
            //Alert("nMes "+Alltrim(Str(nMes))+" Paramentos enviados a SaldoSR, DataBase "+DTOC(dDtParam)+" Vencto De "+dtoc(d2DtParam)+" Vencto Ate "+DTOC(d3DtParam)) 
            nARe1M61a90 := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         ElseIf nMes == 8
            //DbSelectArea("REFLX")
            //nARe2M61a90 := REFLX->TFLXAREC
            //DbCloseArea()
            d8TMPDt := LastDay(dData02)+63  
            d8MesAF := LastDay(d8TMPDt)
            dDtParam  := LastDay(dData02)
            d2DtParam := FirstDay(d8MesAF)
            d3DtParam := LastDay(d8MesAF)             
            nARe2M61a90 := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         ElseIf nMes == 9
            //DbSelectArea("REFLX")
            //nARe3M61a90 := REFLX->TFLXAREC
            //DbCloseArea()
            d9TMPDt := LastDay(dData03)+63
            d9MesAF := LastDay(d9TMPDt)
            dDtParam  := LastDay(dData03)
            d2DtParam := FirstDay(d9MesAF)
            d3DtParam := LastDay(d9MesAF)             
            nARe3M61a90 := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10
            //DbSelectArea("REFLX")
            //nARe1M90M := REFLX->TFLXAREC
            //DbCloseArea()
            d10TMPDt  := LastDay(dDataDe)+94
            d10MesAF := LastDay(d10TMPDt)
            dDtParam  := LastDay(dDataDe)
            d2DtParam := FirstDay(d10MesAF)
            d3DtParam := CTOD("31/12/49")
            nARe1M90M := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         ElseIf nMes == 11
            //DbSelectArea("REFLX")
            //nARe2M90M := REFLX->TFLXAREC
            //DbCloseArea()
            d11TMPDt := LastDay(dData02)+94 
            d11MesAF := LastDay(d11TMPDt)
            dDtParam  := LastDay(dData02)
            d2DtParam := FirstDay(d11MesAF)
            d3DtParam := CTOD("31/12/49")
            nARe2M90M := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         ElseIf nMes == 12
            //DbSelectArea("REFLX")
            //nARe3M90M := REFLX->TFLXAREC
            //DbCloseArea()
            d12TMPDt := LastDay(dData03)+94  
            d12MesAF := LastDay(d12TMPDt)
            dDtParam  := LastDay(dData03)
            d2DtParam := FirstDay(d12MesAF)
            d3DtParam := CTOD("31/12/49")
            nARe3M90M := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         EndIf
      EndIf
      */
      ElseIf nMes >= 7 .And. nMes <= 9
/*
nARAc160dd := 0
nARAc260dd := 0
nARAc360dd := 0
*/
         If nMes == 7
            //DbSelectArea("REFLX")
            //nARe1M61a90 := REFLX->TFLXAREC
            //DbCloseArea()
            d7TMPDt := LastDay(dDataDe)+63  
            d7MesAF := LastDay(d7TMPDt)
            dDtParam  := LastDay(dDataDe)
            d2DtParam := FirstDay(d7MesAF)
            d3DtParam := CTOD("31/12/49")
            //Alert("nMes "+Alltrim(Str(nMes))+" Paramentos enviados a SaldoSR, DataBase "+DTOC(dDtParam)+" Vencto De "+dtoc(d2DtParam)+" Vencto Ate "+DTOC(d3DtParam)) 
            nARAc160dd := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         ElseIf nMes == 8
            //DbSelectArea("REFLX")
            //nARe2M61a90 := REFLX->TFLXAREC
            //DbCloseArea()
            d8TMPDt := LastDay(dData02)+63  
            d8MesAF := LastDay(d8TMPDt)
            dDtParam  := LastDay(dData02)
            d2DtParam := FirstDay(d8MesAF)
            d3DtParam := CTOD("31/12/49")             
            nARAc260dd := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         ElseIf nMes == 9
            //DbSelectArea("REFLX")
            //nARe3M61a90 := REFLX->TFLXAREC
            //DbCloseArea()
            d9TMPDt := LastDay(dData03)+63
            d9MesAF := LastDay(d9TMPDt)
            dDtParam  := LastDay(dData03)
            d2DtParam := FirstDay(d9MesAF)
            d3DtParam := CTOD("31/12/49")             
            nARAc360dd := SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)
         EndIf
      EndIf      
      
   Next nMes

   /************************
   *
   * GASTOS META ORCAMENTO ADM/FIN (OR�ADO)
   *
   ************************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Gastos Metas/Orc Adm/Fin.(Or�ado) Proc. "+Alltrim(Str(nMes))+" / 3")
      
      If nMes == 1
         _xMES := Substr(fNomeMes(Alltrim(StrZero(Month(dDataDe),2))),1,3)
         cQuery := " SELECT SUM(SE7.E7_VAL"+_xMES+"1) AS VxORCMTADM "
      ElseIf nMes == 2
         _xMES := Substr(fNomeMes(Alltrim(StrZero(Month(dData02),2))),1,3)
         cQuery := " SELECT SUM(SE7.E7_VAL"+_xMES+"1) AS VxORCMTADM "
      ElseIf nMes == 3
         _xMES := Substr(fNomeMes(Alltrim(StrZero(Month(dData03),2))),1,3)
         cQuery := " SELECT SUM(SE7.E7_VAL"+_xMES+"1) AS VxORCMTADM "
      EndIf
      
      cQuery += " FROM " + RetSqlName("SE7")+" SE7 "
      
      If nMes == 1
         _xANO := ALLTRIM(STR(YEAR(dDataDe)))
         cQuery += " WHERE SE7.E7_ANO = '"+_xANO+"'" 
      ElseIf nMes == 2 
         _xANO := ALLTRIM(STR(YEAR(dData02)))
         cQuery += " WHERE SE7.E7_ANO = '"+_xANO+"'" 
      ElseIf nMes == 3 
         _xANO := ALLTRIM(STR(YEAR(dData03)))
         cQuery += " WHERE SE7.E7_ANO = '"+_xANO+"'" 
      EndIf
      
      cQuery += "  AND SE7.E7_FILIAL = '"+xFilial("SE7")+"'"
      cQuery += "  AND SUBSTRING(SE7.E7_NATUREZ,1,6) IN "+FormatIn(AdmMetasOrc,",")
      cQuery += "  AND SE7.D_E_L_E_T_ = ' '"
      
      MemoWrite("C:\TEMP\MtOr�adoAdmFin"+Alltrim(Str(nMes))+".SQL",cQuery) 
      
      TCQuery cQuery NEW ALIAS "OMOA"
      
      If nMes == 1
         DbSelectArea("OMOA")
         nAdm1MMtOrc := OMOA->VxORCMTADM
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("OMOA")
         nAdm2MMtOrc := OMOA->VxORCMTADM
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("OMOA")
         nAdm3MMtOrc := OMOA->VxORCMTADM
         DbCloseArea()
      EndIf

   Next nMes

   /************************
   *
   * GASTOS GERAIS ADM.
   *
   ************************/
   
   ProcRegua(48)
   For nMes := 1 To 48 //1 A 3 - Gastos Gerais Adm.
                      //4 A 6 - Gastos com M�o de Obra
                      //7 A 9 - Gastos com M�o de Obra(Despesas Fixas)
                      //10 A 12 - Gastos com M�o de Obra(Despesas Peri�dicas)
                      //13 A 15 - Gastos com M�o de Obra(Rescis�es)
                      //16 A 18 - Gastos com M�o de Obra(Premia��es)
                      //19 A 21 - Despesas de Viagens
                      //22 A 24 - Outras Despesas Adm
                      //25 A 27 - Servi�os Terceirizados
                      //28 A 30 - Outras Naturezas Adm.
                      //31 A 33 - Gastos Metas Or�amento Adm.
                      //34 A 36 - Empr�stimos entre empresas
                      //37 a 39 - Fundo Fixo
                      
                      //40 a 42 - Honor�rios Advocat�cios
                      //43 a 45 - Despesas Judiciais
                      //46 a 48 - Rescis�es/Acordos
                      
      IncProc(" Calc.Gastos Gerais Adm. Proc. "+Alltrim(Str(nMes))+" / 48")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLGGADM "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1 .Or. nMes == 4 .Or. nMes == 7 .Or. nMes == 10 .Or. nMes == 13 .Or. nMes == 16 .Or. nMes == 19 .Or. nMes == 22 .Or. nMes == 25 .Or. nMes == 28 .Or. nMes == 31 .Or. nMes == 34 .Or. nMes == 37 .Or. nMes == 40 .Or. nMes == 43 .Or. nMes == 46
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2 .Or. nMes == 5 .Or. nMes == 8 .Or. nMes == 11 .Or. nMes == 14 .Or. nMes == 17 .Or. nMes == 20 .Or. nMes == 23 .Or. nMes == 26 .Or. nMes == 29 .Or. nMes == 32 .Or. nMes == 35 .Or. nMes == 38 .Or. nMes == 41 .Or. nMes == 44 .Or. nMes == 47
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3 .Or. nMes == 6 .Or. nMes == 9 .Or. nMes == 12 .Or. nMes == 15 .Or. nMes == 18 .Or. nMes == 21 .Or. nMes == 24 .Or. nMes == 27 .Or. nMes == 30 .Or. nMes == 33 .Or. nMes == 36 .Or. nMes == 39 .Or. nMes == 42 .Or. nMes == 45 .Or. nMes == 48
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
      
      If nMes <= 3
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cGstGerADM,",")
      ElseIf nMes >= 4 .And. nMes <= 6
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cMOGstADM,",")
      ElseIf nMes >= 7 .And. nMes <= 9
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cDFGstADM,",")
      ElseIf nMes >= 10 .And. nMes <= 12
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cDPGstADM,",")
      ElseIf nMes >= 13 .And. nMes <= 15
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cResGstADM,",")
      ElseIf nMes >= 16 .And. nMes <= 18
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cPrmADMGst,",")
      ElseIf nMes >= 19 .And. nMes <= 21
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cDspViaADM,",")
      ElseIf nMes >= 22 .And. nMes <= 24
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cOutDADMGst,",")
      ElseIf nMes >= 25 .And. nMes <= 27
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cSrvTADMGst,",")
      ElseIf nMes >= 28 .And. nMes <= 30
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(OutrasADM,",")
      ElseIf nMes >= 31 .And. nMes <= 33
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(AdmMetasOrc,",")
      ElseIf nMes >= 34 .And. nMes <= 36
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtEmpEEmp,",")
      ElseIf nMes >= 37 .And. nMes <= 39
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtFundFix,",")
      ElseIf nMes >= 40 .And. nMes <= 42
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cHonorAdvNt,",")
      ElseIf nMes >= 43 .And. nMes <= 45
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cDespJdcNt,",")
      ElseIf nMes >= 46 .And. nMes <= 48
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cRescAcordNt,",")
      EndIf
      
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      
      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCF','PA ','NDF')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"      
      
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      If nMes == 7
         MemoWrite("C:\TEMP\GastGerADM"+Alltrim(Str(nMes))+".SQL",cQuery)
      ElseIf nMes == 31
         MemoWrite("C:\TEMP\GastGerADM"+Alltrim(Str(nMes))+".SQL",cQuery)
      EndIf
      
      TCQuery cQuery NEW ALIAS "TGADM"
      
      If nMes <= 3
         If nMes == 1
            DbSelectArea("TGADM")
            nVl1MAdmGG := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("TGADM")
            nVl2MAdmGG := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMEs == 3
            DbSelectArea("TGADM")
            nVl3MAdmGG := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("TGADM")
            nVGG1MMOADM := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("TGADM")
            nVGG2MMOADM := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 6
            DbSelectArea("TGADM")
            nVGG3MMOADM := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("TGADM")
            nVGG1MDFADM := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 8
            DbSelectArea("TGADM")
            nVGG2MDFADM := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 9
            DbSelectArea("TGADM")
            nVGG3MDFADM := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10
            DbSelectArea("TGADM")
            nVGG1MDPAD := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 11
            DbSelectArea("TGADM")
            nVGG2MDPAD := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 12
            DbSelectArea("TGADM")
            nVGG3MDPAD := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 13 .And. nMes <= 15
         If nMes == 13
            DbSelectArea("TGADM")
            nVRes1MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 14
            DbSelectArea("TGADM")
            nVRes2MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 15
            DbSelectArea("TGADM")
            nVRes3MAdm := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 16 .And. nMes <= 18
         If nMes == 16
            DbSelectArea("TGADM")
            nVPrm1MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 17
            DbSelectArea("TGADM")
            nVPrm2MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 18
            DbSelectArea("TGADM")
            nVPrm3MAdm := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 19 .And. nMes <= 21
         If nMes == 19
            DbSelectArea("TGADM")
            nVDVia1MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 20
            DbSelectArea("TGADM")
            nVDVia2MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 21
            DbSelectArea("TGADM")
            nVDVia3MAdm := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 22 .And. nMes <= 24
         If nMes == 22
            DbSelectArea("TGADM")
            nVOut1MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 23
            DbSelectArea("TGADM")
            nVOut2MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 24
            DbSelectArea("TGADM")
            nVOut3MAdm := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 25 .And. nMes <= 27
         If nMes == 25
            DbSelectArea("TGADM")
            nVSvT1MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 26
            DbSelectArea("TGADM")
            nVSvT2MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 27
            DbSelectArea("TGADM")
            nVSvT3MAdm := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 28 .And. nMes <= 30
         If nMes == 28
            DbSelectArea("TGADM")
            nOut1MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 29
            DbSelectArea("TGADM")
            nOut2MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 30
            DbSelectArea("TGADM")
            nOut3MAdm := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 31 .And. nMes <= 33
         If nMes == 31
            DbSelectArea("TGADM")
            nGMt1MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 32
            DbSelectArea("TGADM")
            nGMt2MAdm := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 33
            DbSelectArea("TGADM")
            nGMt3MAdm := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 34 .And. nMes <= 36
         If nMes == 34
            DbSelectArea("TGADM")
            nEmp1MAdmEE := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 35
            DbSelectArea("TGADM")
            nEmp2MAdmEE := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 36
            DbSelectArea("TGADM")
            nEmp3MAdmEE := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 37 .And. nMes <= 39
         If nMes == 37
            DbSelectArea("TGADM")
            nFnd1MFixo := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 38
            DbSelectArea("TGADM")
            nFnd2MFixo := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 39
            DbSelectArea("TGADM")
            nFnd3MFixo := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 40 .And. nMes <= 42

         If nMes == 40
            DbSelectArea("TGADM")
            nHonor1MAdv := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 41
            DbSelectArea("TGADM")
            nHonor2MAdv := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 42
            DbSelectArea("TGADM")
            nHonor3MAdv := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 43 .And. nMes <= 45
         If nMes == 43
            DbSelectArea("TGADM")
            nDsp1MJudc := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 44
            DbSelectArea("TGADM")
            nDsp2MJudc := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 45
            DbSelectArea("TGADM")
            nDsp3MJudc := TGADM->VLGGADM
            DbCloseArea()
         EndIf      
      ElseIf nMes >= 46 .And. nMes <= 48
         If nMes == 46
            DbSelectArea("TGADM")
            nRes1MAcAd := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 47
            DbSelectArea("TGADM")
            nRes2MAcAd := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 48
            DbSelectArea("TGADM")
            nRes3MAcAd := TGADM->VLGGADM
            DbCloseArea()
         EndIf      
      EndIf
      
   Next nMes   

   /************************
   *
   * IMPOSTOS CONTRIB. (Total)
   *
   ************************/
   ProcRegua(36)
   For nMes := 1 To 36 //1 A 3 - Impostos Contribui��o Total.
                      //4 A 6 - INSS
                      //7 A 9 - FGTS
                      //10 A 12 - CPMF
                      //13 A 15 - IPTU
                      //16 A 18 - IPI
                      //19 A 21 - ICMS
                      //22 A 24 - ISS
                      //25 A 27 - IR
                      //28 A 30 - TRIBUTOS IMPORTA��O
                      //31 A 33 - PIS/COFINS
                      //34 A 36 - TRIBUTOS

                      
      IncProc(" Impostos e Contribui��o Proc. "+Alltrim(Str(nMes))+" / 36")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLGGADM "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1 .Or. nMes == 4 .Or. nMes == 7 .Or. nMes == 10 .Or. nMes == 13 .Or. nMes == 16 .Or. nMes == 19 .Or. nMes == 22 .Or. nMes == 25 .Or. nMes == 28 .Or. nMes == 31
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2 .Or. nMes == 5 .Or. nMes == 8 .Or. nMes == 11 .Or. nMes == 14 .Or. nMes == 17 .Or. nMes == 20 .Or. nMes == 23 .Or. nMes == 26 .Or. nMes == 29 .Or. nMes == 32
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3 .Or. nMes == 6 .Or. nMes == 9 .Or. nMes == 12 .Or. nMes == 15 .Or. nMes == 18 .Or. nMes == 21 .Or. nMes == 24 .Or. nMes == 27 .Or. nMes == 30 .Or. nMes == 33
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
      
      If nMes <= 3
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtImpCtrbTt,",")
      ElseIf nMes >= 4 .And. nMes <= 6
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtINSS,",")
      ElseIf nMes >= 7 .And. nMes <= 9
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtFGTS,",")
      ElseIf nMes >= 10 .And. nMes <= 12
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtCPMF,",")
      ElseIf nMes >= 13 .And. nMes <= 15
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtIPTU,",")
      ElseIf nMes >= 16 .And. nMes <= 18
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtIPI,",")
      ElseIf nMes >= 19 .And. nMes <= 21
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtICMS,",")
      ElseIf nMes >= 22 .And. nMes <= 24
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtISS,",")
      ElseIf nMes >= 25 .And. nMes <= 27
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtIR,",")
      ElseIf nMes >= 28 .And. nMes <= 30
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtImpImport,",")
      ElseIf nMes >= 31 .And. nMes <= 33
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtPISCOF,",")
      ElseIf nMes >= 34 .And. nMes <= 36
         cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cNtTribut,",")
      EndIf
      
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "TGADM"
      
      //ww
      If nMes <= 3
         If nMes == 1
            DbSelectArea("TGADM")
            nImp1MCTt := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("TGADM")
            nImp2MCTt := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMEs == 3
            DbSelectArea("TGADM")
            nImp3MCTt := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("TGADM")
            nV1MINSS := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("TGADM")
            nV2MINSS := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 6
            DbSelectArea("TGADM")
            nV3MINSS := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("TGADM")
            nV1MFGTS := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 8
            DbSelectArea("TGADM")
            nV2MFGTS := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 9
            DbSelectArea("TGADM")
            nV3MFGTS := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10
            DbSelectArea("TGADM")
            nV1MCPMF := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 11
            DbSelectArea("TGADM")
            nV2MCPMF := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 12
            DbSelectArea("TGADM")
            nV3MCPMF := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 13 .And. nMes <= 15
         If nMes == 13
            DbSelectArea("TGADM")
            nV1MIPTU := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 14
            DbSelectArea("TGADM")
            nV2MIPTU := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 15
            DbSelectArea("TGADM")
            nV3MIPTU := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 16 .And. nMes <= 18
         If nMes == 16
            DbSelectArea("TGADM")
            nV1MIPI := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 17
            DbSelectArea("TGADM")
            nV2MIPI := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 18
            DbSelectArea("TGADM")
            nV3MIPI := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 19 .And. nMes <= 21
         If nMes == 19
            DbSelectArea("TGADM")
            nV1MICMS := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 20
            DbSelectArea("TGADM")
            nV2MICMS := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 21
            DbSelectArea("TGADM")
            nV3MICMS := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 22 .And. nMes <= 24
         If nMes == 22
            DbSelectArea("TGADM")
            nV1MISS := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 23
            DbSelectArea("TGADM")
            nV2MISS := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 24
            DbSelectArea("TGADM")
            nV3MISS := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 25 .And. nMes <= 27
         If nMes == 25
            DbSelectArea("TGADM")
            nV1MIRImp := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 26
            DbSelectArea("TGADM")
            nV2MIRImp := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 27
            DbSelectArea("TGADM")
            nV3MIRImp := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 28 .And. nMes <= 30
         If nMes == 28
            DbSelectArea("TGADM")
            nV1MImpImport := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 29
            DbSelectArea("TGADM")
            nV2MImpImport := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 30
            DbSelectArea("TGADM")
            nV3MImpImport := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 31 .And. nMes <= 33
         If nMes == 31
            DbSelectArea("TGADM")
            nPIS1MCOF := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 32
            DbSelectArea("TGADM")
            nPIS2MCOF := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 33
            DbSelectArea("TGADM")
            nPIS3MCOF := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 34 .And. nMes <= 36
         If nMes == 34
            DbSelectArea("TGADM")
            n1MTribut := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 35
            DbSelectArea("TGADM")
            n2MTribut := TGADM->VLGGADM
            DbCloseArea()
         ElseIf nMes == 36
            DbSelectArea("TGADM")
            n3MTribut := TGADM->VLGGADM
            DbCloseArea()
         EndIf
      EndIf
      
   Next nMes   

EndIf

If lLogistica
   
   /********************************
   * FATURAMENTO FOB / CIF 
   *********************************/
   
   ProcRegua(9)
   
   For nMes := 1 To 9
      
      IncProc("Calc. Faturamento FOB/CIF, Processo "+Alltrim(Str(nMes))+" / 9")

      cQuery := " SELECT SUBSTRING(SD2.D2_EMISSAO,1,6) AS EMISSAO, SUM(D2_TOTAL+CASE WHEN SD2.D2_TIPO <> 'P' THEN D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA ELSE 0 END) AS TFTFOBCIF FROM "
      cQuery += RetSqlName('SD2') + " SD2,"+RetSQLName("SB1")+" SB1,"+RetSQLName("SF2")+" SF2,"+RetSQLName("SC5")+" SC5"
      cQuery += " WHERE SD2.D2_FILIAL  = '"+xFilial("SD2")+"' AND  SD2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SB1.B1_FILIAL  = '"+xFilial("SB1")+"' AND  SB1.D_E_L_E_T_ = ' ' "
      cQuery += " AND SF2.F2_FILIAL  = '"+xFilial("SF2")+"' AND  SF2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SC5.C5_FILIAL  = '"+xFilial("SC5")+"' AND  SC5.D_E_L_E_T_ = ' ' "
      cQuery += " AND SD2.D2_COD = SB1.B1_COD"
      //cQuery += " AND SUBSTRING(SB1.B1_GRUPO,1,2) <> '80'" //Faturamento sem Sucatas
      cQuery += " AND SF2.F2_DOC = SD2.D2_DOC"
      cQuery += " AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQuery += " AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQuery += " AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      
      cQuery += " AND SD2.D2_PEDIDO = SC5.C5_NUM"
      
      If nMes <= 3
         cQuery += " AND SC5.C5_TPFRETE = 'F'"
      ElseIf nMEs >= 4 .And. nMes <= 6
         cQuery += " AND SC5.C5_TPFRETE = 'C'"
      ElseIf nMes >= 7 .And. nMes <= 9
         cQuery += " AND SC5.C5_TPFRETE = ' '"
      EndIf
      
      If nMes == 1 .Or. nMes == 4 .Or. nMes == 7
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
      ElseIf nMes == 2 .Or. nMes == 5 .Or. nMes == 8
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
      ElseIf nMes == 3 .Or. nMes == 6 .Or. nMes == 9
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
      EndIf
      
      cQuery += " GROUP BY SUBSTRING(SD2.D2_EMISSAO,1,6) "
      cQuery += " ORDER BY SUBSTRING(SD2.D2_EMISSAO,1,6) ASC"
            
      MemoWrite("C:\TEMP\FatCIFFOB.SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "CFT1"
      
      If nMes <= 3
         If nMes == 1
            DbSelectArea("CFT1")
               nFOB1MFt := CFT1->TFTFOBCIF
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("CFT1")
               nFOB2MFt := CFT1->TFTFOBCIF
            DbCloseArea()
         ElseIf nMes == 3
            DbSelectArea("CFT1")
               nFOB3MFt := CFT1->TFTFOBCIF
            DbCloseArea()
         EndIf
      ElseIf nMEs >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("CFT1")
               nFt1MCIF := CFT1->TFTFOBCIF
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("CFT1")
               nFt2MCIF := CFT1->TFTFOBCIF
            DbCloseArea()
         ElseIf nMes == 6
            DbSelectArea("CFT1")
               nFt3MCIF := CFT1->TFTFOBCIF
            DbCloseArea()
         EndIf
      ElseIf nMEs >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("CFT1")
               nFt1MSDEF := CFT1->TFTFOBCIF
            DbCloseArea()
         ElseIf nMes == 8
            DbSelectArea("CFT1")
               nFt2MSDEF := CFT1->TFTFOBCIF
            DbCloseArea()
         ElseIf nMes == 9
            DbSelectArea("CFT1")
               nFt3MSDEF := CFT1->TFTFOBCIF
            DbCloseArea()
         EndIf
      EndIf
      
   Next nMes
   

   /********************************
   * Calculando Frete nas Notas Fiscais - F2_FRETE
   *********************************/
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Cal.Fretes nas NFs Vendas, Proc.: "+Alltrim(Str(nMes))+" / 3")
      
      cQrySF2 := "SELECT SUM(SF2.F2_FRETE) AS VLRFRETE "
      cQrySF2 += " FROM "+RetSQLName("SF2")+" SF2"
      cQrySF2 += " WHERE SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      
      If nMes == 1
         cQrySF2 += " AND SUBSTRING(SF2.F2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
      ElseIf nMes == 2
         cQrySF2 += " AND SUBSTRING(SF2.F2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
      ElseIf nMes == 3
         cQrySF2 += " AND SUBSTRING(SF2.F2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
      EndIf
      
      cQrySF2 += " AND SF2.D_E_L_E_T_ <> '*'"
      
      MemoWrite("C:\TEMP\VlrFrete.SQL",cQrySF2)
      
      TCQuery cQrySF2 NEW ALIAS "F2FRT"
      
      If nMes == 1
         DbSelectArea("F2FRT")
         nVl1MsFrete := F2FRT->VLRFRETE
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("F2FRT")
         nVl2MsFrete := F2FRT->VLRFRETE
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("F2FRT")
         nVl3MsFrete := F2FRT->VLRFRETE
         DbCloseArea()
      EndIf
      
      
   Next nMes

   /*************************
   *
   * FRETES (Total)
   *
   *************************/

   ProcRegua(24)
   
   For nMes := 1 To 24
   
      IncProc("FRETES (Total), Processo "+Alltrim(Str(nMes))+" / 24")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLBAIXAS "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      If nMes == 1 .Or. nMes == 4 .Or. nMes == 7 .Or. nMes == 10 .Or. nMes == 13 .Or. nMes == 16 .Or. nMes == 19 .Or. nMes == 22
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2 .Or. nMes == 5 .Or. nMes == 8 .Or. nMes == 11 .Or. nMes == 14 .Or. nMes == 17 .Or. nMes == 20 .Or. nMes == 23
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3 .Or. nMes == 6 .Or. nMes == 9 .Or. nMes == 12 .Or. nMes == 15 .Or. nMes == 18 .Or. nMes == 21 .Or. nMes == 24
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      If nMes <= 3
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtVendFrt,",")
      ElseIf nMes >= 4 .And. nMes <= 6
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtImporFrt,",")
      ElseIf nMes >= 7 .And. nMes <= 9
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtExporFrt,",")
      ElseIf nMes >= 10 .And. nMes <= 12
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtBenefFrt,",")
      ElseIf nMes >= 13 .And. nMes <= 15
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtTransfFrt,",")
      ElseIf nMes >= 16 .And. nMes <= 18
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtComercFrt,",")
      ElseIf nMes >= 19 .And. nMes <= 21
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtIndusFrt,",")
      ElseIf nMes >= 22 .And. nMes <= 24
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtAdminFrt,",")
      EndIf
      
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "TSE5"
      
      If nMEs <= 3
         If nMes == 1
            DbSelectArea("TSE5")
            n1FrtVeMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("TSE5")
            n2FrtVeMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMEs == 3
            DbSelectArea("TSE5")
            n3FrtVeMVl := TSE5->VLBAIXAS
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("TSE5")
            n1FrtImpMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("TSE5")
            n2FrtImpMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMEs == 6
            DbSelectArea("TSE5")
            n3FrtImpMVl := TSE5->VLBAIXAS
            DbCloseArea()
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("TSE5")
            n1FrtExpMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMes == 8
            DbSelectArea("TSE5")
            n2FrtExpMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMEs == 9
            DbSelectArea("TSE5")
            n3FrtExpMVl := TSE5->VLBAIXAS
            DbCloseArea()
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10
            DbSelectArea("TSE5")
            n1FrtBenMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMes == 11
            DbSelectArea("TSE5")
            n2FrtBenMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMEs == 12
            DbSelectArea("TSE5")
            n3FrtBenMVl := TSE5->VLBAIXAS
            DbCloseArea()
         EndIf
      ElseIf nMes >= 13 .And. nMes <= 15
         If nMes == 13
            DbSelectArea("TSE5")
            n1FrtTrfMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMes == 14
            DbSelectArea("TSE5")
            n2FrtTrfMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMEs == 15
            DbSelectArea("TSE5")
            n3FrtTrfMVl := TSE5->VLBAIXAS
            DbCloseArea()
         EndIf
      ElseIf nMes >= 16 .And. nMes <= 18
         If nMes == 16
            DbSelectArea("TSE5")
            n1FrtComMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMes == 17
            DbSelectArea("TSE5")
            n2FrtComMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMEs == 18
            DbSelectArea("TSE5")
            n3FrtComMVl := TSE5->VLBAIXAS
            DbCloseArea()
         EndIf
      ElseIf nMes >= 19 .And. nMes <= 21
         If nMes == 19
            DbSelectArea("TSE5")
            n1FrtIndMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMes == 20
            DbSelectArea("TSE5")
            n2FrtIndMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMEs == 21
            DbSelectArea("TSE5")
            n3FrtIndMVl := TSE5->VLBAIXAS
            DbCloseArea()
         EndIf
      ElseIf nMes >= 22 .And. nMes <= 24
         If nMes == 22
            DbSelectArea("TSE5")
            n1FrtAdmMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMes == 23
            DbSelectArea("TSE5")
            n2FrtAdmMVl := TSE5->VLBAIXAS
            DbCloseArea()
         ElseIf nMEs == 24
            DbSelectArea("TSE5")
            n3FrtAdmMVl := TSE5->VLBAIXAS
            DbCloseArea()
         EndIf
      EndIf
      
   Next nMes
   
   /***************************
   * DESPESAS - DEMONSTRATIVO DE RESULTADOS
   *
   * Salarios - Demonstrativo de Resultados
   * TODAS AS DESPESAS DA M.O. DA IND�STRIA INCLUINDO A SUA AREA ADMINISTRATIVA '100301','100302','100303','100304','100305','100306','100307',
   * '100309','100310','100311','100312','100313','100317','100401','100402','100403','100404','100405','100406','100407','100409','100410','100411','100412','100417'
   *
   * //TODAS AS DESPESAS DO COMERCIAL (EXCLUI A NATUREZA DE EMPR�STIMO E REEMBOLSO A CLIENTE) '100201','100202','100203','100204','100205','100206',
   * //'100207','100208','100209','100210','100211','100213','100214','100215','100216','100222','200201','200202','200203','200204','200205','200206','200207','600101','600102','600303','600401'
   * 
   *
   ***************************/
   //kxx
   ProcRegua(12) //4)
   
   For nMes := 1 To 12 //4
   
      IncProc("Calc.Despesas DRE, Proc. "+Alltrim(Str(nMes))+" / 12")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLMPTOTAL "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"

      If nMes == 1 .Or. nMes == 4 .Or. nMes == 7 .Or. nMes == 10
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2 .Or. nMes == 5 .Or. nMes == 8 .Or. nMes == 11
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3 .Or. nMes == 6 .Or. nMes == 9 .Or. nMes == 12
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      If nMes <= 3
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtMOINDADM,",")
      ElseIf nMes >= 4 .And. nMes <= 6
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cDspCDER,",")
      ElseIf nMes  >= 7 .And. nMes <= 9
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNtAdmDER,",")
      ElseIf nMes  >= 10 .And. nMes <= 12
         cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNDpFinDER,",")
      EndIf
      
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "VMPTT"
      
         If nMes <= 3
            If nMes == 1
               DbSelectArea("VMPTT")
               nMPMO1IndAdm := VMPTT->VLMPTOTAL
               DbCloseArea()
            ElseIf nMes == 2
               DbSelectArea("VMPTT")
               nMPMO2IndAdm := VMPTT->VLMPTOTAL
               DbCloseArea()
            ElseIf nMes == 3
               DbSelectArea("VMPTT")
               nMPMO3IndAdm := VMPTT->VLMPTOTAL
               DbCloseArea()
            EndIf
         ElseIf nMes >= 4 .And. nMes <= 6
            If nMes == 4
               DbSelectArea("VMPTT")
               nDesp1ComDRE := VMPTT->VLMPTOTAL
               DbCloseArea()
            ElseIf nMes == 5
               DbSelectArea("VMPTT")
               nDesp2ComDRE := VMPTT->VLMPTOTAL
               DbCloseArea()
            ElseIf nMes == 6
               DbSelectArea("VMPTT")
               nDesp3ComDRE := VMPTT->VLMPTOTAL
               DbCloseArea()
            EndIf
         ElseIf nMes >= 7 .And. nMes <= 9
            If nMes == 7
               DbSelectArea("VMPTT")
               nDsp1ADMDER := VMPTT->VLMPTOTAL
               DbCloseArea()
            ElseIf nMes == 8
               DbSelectArea("VMPTT")
               nDsp2ADMDER := VMPTT->VLMPTOTAL
               DbCloseArea()
            ElseIf nMes == 9
               DbSelectArea("VMPTT")
               nDsp3ADMDER := VMPTT->VLMPTOTAL
               DbCloseArea()
            EndIf
         ElseIf nMes >= 10 .And. nMes <= 12
            If nMes == 10
               DbSelectArea("VMPTT")
               nDspFi1nsDER := VMPTT->VLMPTOTAL
               DbCloseArea()
            ElseIf nMes == 11
               DbSelectArea("VMPTT")
               nDspFi2nsDER := VMPTT->VLMPTOTAL
               DbCloseArea()
            ElseIf nMes == 12
               DbSelectArea("VMPTT")
               nDspFi3nsDER := VMPTT->VLMPTOTAL
               DbCloseArea()
            EndIf
         EndIf
      
   Next nMes

   /***************************
   * RECEITAS - DEMONSTRATIVO DE RESULTADOS
   ***************************/
   //kxx
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc.Receitas DRE, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SUM(SE5.E5_VALOR) AS VLMPTOTAL "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'R'"

      If nMes == 1
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dDataDe),1,6)+ "'"
      ElseIf nMes == 2
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData02),1,6)+ "'"
      ElseIf nMes == 3
         cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,6) = '" +Substr(DTOS(dData03),1,6)+ "'"
      EndIf
   
      cQuery += "  AND SE5.E5_NATUREZ IN "+FormatIn(cNRctFinDRE,",")
      
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      cQuery += "  AND SE5.E5_SITUACA = ' '"
      cQuery += "  AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"
      cQuery += "  AND SE5.E5_FILIAL = '" + xFilial("SE5") + "'"
      cQuery += "  AND (SELECT COUNT(*)"
      cQuery += "         FROM "+RetSQLName("SE5")+" TMP"
      cQuery += "       WHERE TMP.E5_FILIAL = '"+xFilial("SE5")+"'"
      cQuery += "         AND TMP.E5_FILIAL = SE5.E5_FILIAL"
      cQuery += "         AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"
      cQuery += "         AND TMP.E5_NUMERO = SE5.E5_NUMERO"
      cQuery += "         AND TMP.E5_PARCELA = SE5.E5_PARCELA"
      cQuery += "         AND TMP.E5_TIPO = SE5.E5_TIPO"
      cQuery += "         AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQuery += "         AND TMP.E5_LOJA = SE5.E5_LOJA"
      cQuery += "         AND TMP.E5_VALOR = SE5.E5_VALOR"   
      cQuery += "         AND TMP.E5_SEQ = SE5.E5_SEQ"
      cQuery += "         AND TMP.E5_TIPODOC = 'ES'"
      cQuery += "         AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      TCQuery cQuery NEW ALIAS "VMPTT"
      
         If nMes == 1
            DbSelectArea("VMPTT")
            nRecDR1EFin := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("VMPTT")
            nRecDR2EFin := VMPTT->VLMPTOTAL
            DbCloseArea()
         ElseIf nMes == 3
            DbSelectArea("VMPTT")
            nRecDR3EFin := VMPTT->VLMPTOTAL
            DbCloseArea()
         EndIf
      
   Next nMes

EndIf

If lDemonsRes
   
   ProcRegua(6)
   
   For nMes := 1 To 6  //Faturamento por Condi��o(Pagamento) A Vista/A Prazo
      
      IncProc("Calc.Fat. A Vista/A Prazo Proc. "+Alltrim(Str(nMes))+" / 6")
      
      //Faturamento A Vista/A Prazo - Antes, cQuery := " SELECT SUM(SD2.D2_TOTAL+SD2.D2_VALIPI+SD2.D2_ICMSRET+SD2.D2_VALFRE+SD2.D2_SEGURO+SD2.D2_DESPESA) AS VALBRUTO FROM "
      cQuery := " SELECT SUM(D2_TOTAL+CASE WHEN SD2.D2_TIPO <> 'P' THEN D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA ELSE 0 END) AS VALBRUTO FROM "
      cQuery += RetSqlName('SD2') + " SD2,"+RetSQLName("SB1")+" SB1,"+RetSQLName("SF2")+" SF2"
      cQuery += " WHERE SD2.D2_FILIAL  = '"+xFilial("SD2")+"' AND  SD2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SB1.B1_FILIAL  = '"+xFilial("SB1")+"' AND  SB1.D_E_L_E_T_ = ' ' "
      cQuery += " AND SF2.F2_FILIAL  = '"+xFilial("SF2")+"' AND  SF2.D_E_L_E_T_ = ' ' "
      cQuery += " AND SD2.D2_COD = SB1.B1_COD"
      cQuery += " AND SUBSTRING(SB1.B1_GRUPO,1,2) <> '80'" //Faturamento sem Sucatas
      cQuery += " AND SF2.F2_DOC = SD2.D2_DOC"
      cQuery += " AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQuery += " AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQuery += " AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      
      If nMes == 1 .Or. nMes == 4
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
      ElseIf nMes == 2 .Or. nMes == 5
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
      ElseIf nMes == 3 .Or. nMes == 6
         cQuery += " AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
      EndIf
            
      If nMes <= 3
         cQuery += " AND SF2.F2_COND = 'V01'"
      ElseIf nMes >= 4
         cQuery += " AND SF2.F2_COND <> 'V01'"
      EndIf
      
      If nMes == 1
         MemoWrite("C:\TEMP\FatAVisAPrz"+Alltrim(Str(nMes))+".SQL",cQuery)
      EndIf
      
      TCQuery cQuery NEW ALIAS "CFT"
      
      //If nMes == 1
      TCSetField("CFT","VALBRUTO","N",14,02)
      //EndIf
      
      If nMes <= 3
         If nMes == 1
            DbSelectArea("CFT")
            nFat1AVista := CFT->VALBRUTO
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("CFT")
            nFat2AVista := CFT->VALBRUTO
            DbCloseArea()
         ElseIf nMes == 3
            DbSelectArea("CFT")
            nFat3AVista := CFT->VALBRUTO
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4
         If nMes == 4
            DbSelectArea("CFT")
            nFat1APrazo := CFT->VALBRUTO  
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("CFT")
            nFat2APrazo := CFT->VALBRUTO  
            DbCloseArea()
         ElseIf nMes == 6
            DbSelectArea("CFT")
            nFat3APrazo := CFT->VALBRUTO  
            DbCloseArea()
         EndIf
      EndIf
      
   Next nMes

   /********************************
   * Calculando Impostos nas Notas Fiscais de Vendas
   *********************************/
   
   ProcRegua(6)
   
   For nMes := 1 To 6
      
      IncProc("Calc.Impostos NFs de Vendas, Processo "+Alltrim(Str(nMes))+" / 6")

      If nMes <= 3
         cQuery := " SELECT SUM(F3_VALIPI) AS VALIPI, SUM(F3_VALICM) AS VALICM FROM "
      ElseIf nMes >= 4
         cQuery := " SELECT SUM(F3_VALCONT) AS VALORVEN FROM "
      EndIf
      cQuery += RetSqlName('SF3') + " SF3"
      cQuery += " WHERE SF3.F3_FILIAL  = '"+xFilial("SF3")+"' AND  SF3.D_E_L_E_T_ = ' ' "
      cQuery += " AND SF3.F3_CFO IN "+FormatIn(cCFOPVenda,",")
      
      If nMes == 1 .Or. nMes == 4
         cQuery += " AND  SUBSTRING(SF3.F3_ENTRADA,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
      ElseIf nMes == 2 .Or. nMes == 5
         cQuery += " AND  SUBSTRING(SF3.F3_ENTRADA,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
      ElseIf nMes == 3 .Or. nMes == 6
         cQuery += " AND  SUBSTRING(SF3.F3_ENTRADA,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
      EndIf
      
      cQuery += " AND SF3.F3_OBSERV <> 'NF CANCELADA' "
      cQuery += " AND SF3.D_E_L_E_T_ <> '*'"
      //ElseIf nMes == 2
      //   cQuery += " AND  SUBSTRING(SF3.F3_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
      //ElseIf nMes == 3
      //   cQuery += " AND  SUBSTRING(SF3.F3_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
      //EndIf
      
      MemoWrite("C:\TEMP\ImpICMIPI.SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "CFT1"
      
      If nMes <= 3
         If nMes == 1
            DbSelectArea("CFT1")
            nIPI1Valor := CFT1->VALIPI
            nICM1Vlor  := CFT1->VALICM
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("CFT1")
            nIPI2Valor := CFT1->VALIPI
            nICM2Vlor  := CFT1->VALICM
            DbCloseArea()
         ElseIf nMes == 3
            DbSelectArea("CFT1")
            nIPI3Valor := CFT1->VALIPI
            nICM3Vlor  := CFT1->VALICM
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4
         //CALCULAR PIS (1,65%) E COFINS (7,60%) SOBRE O FATURAMENTO TOTAL DE VENDAS (LIVRO FISCAL DE SA�DA)
         If nMes == 4
            DbSelectArea("CFT1")
            nPIS1Valor := (CFT1->VALORVEN * (1.65))/100
            nCOF1Valor := (CFT1->VALORVEN * (7.60))/100
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("CFT1")
            nPIS2Valor := (CFT1->VALORVEN * (1.65))/100
            nCOF2Valor := (CFT1->VALORVEN * (7.60))/100
            DbCloseArea()
         ElseIf nMes == 6
            DbSelectArea("CFT1")
            nPIS3Valor := (CFT1->VALORVEN * (1.65))/100
            nCOF3Valor := (CFT1->VALORVEN * (7.60))/100
            DbCloseArea()
         EndIf
      EndIf
      
   Next nMes
   
EndIf

Return

/*
�������������������������������������������������������������������������Ĵ��
���Sintaxe   � SaldoTit(ExpC1,ExpC2,ExpC3,ExpC4,ExpC5,ExpC6,ExpC7,ExpN1)  ���
�������������������������������������������������������������������������Ĵ��
���Parametros� ExpC1=Numero do Prefixo                                    ���
���          � ExpC2=Numero do Titulo                                     ���
���          � ExpC3=Parcela                                              ���
���          � ExpC4=Tipo                                                 ���
���          � ExpC5=Natureza                                             ���
���          � ExpC6=Carteira  (R/P)                                      ���
���          � ExpC7=Fornecedor (se ExpC6 = 'R')                          ���
���          � ExpN1=Moeda                                                ���
���          � ExpD1=Data para conversao                                  ���
���          � ExpD2=Data data baixa a ser considerada (retroativa)       ���
���          � ExpC8=Loja do titulo                                       ���
���          � ExpC9=Filial do titulo                                     ���
���          � Expn2=Taxa da Moeda                                        ���
���          � Expn3=Tipo de data para compor saldo (baixa/dispo/digit)   ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function TitSaldo(cPrefixo,cNumero,cParcela,cTipo,cNatureza,cCart,cCliFor,nMoeda,;
						dData,dDataBaixa,cLoja,cFilTit,nTxMoeda,nTipoData)
//Tipos de Data (cTipoData ou xTipoData)
// 0 = Data Da Baixa (E5_DATA) (Default)
// 1 = Data de Disponibilidade (E5_DTDISPO)
// 2 = Data de Contabilida��o (E5_DTDIGIT)

#IFDEF TOP
	Local nSaldo     := 0
	Local cxFilial   := nil
	Local cProcedure := 'FIN002'
	Local cTiPoData  := "0"
	Local cAliasTit
	
	Local lPCCBaixa := SuperGetMv("MV_BX10925",.T.,"2") == "1"  .and. (!Empty( SE5->( FieldPos( "E5_VRETPIS" ) ) ) .And. !Empty( SE5->( FieldPos( "E5_VRETCOF" ) ) ) .And. ; 
				 !Empty( SE5->( FieldPos( "E5_VRETCSL" ) ) ) .And. !Empty( SE5->( FieldPos( "E5_PRETPIS" ) ) ) .And. ;
				 !Empty( SE5->( FieldPos( "E5_PRETCOF" ) ) ) .And. !Empty( SE5->( FieldPos( "E5_PRETCSL" ) ) ) .And. ;
				 !Empty( SE2->( FieldPos( "E2_SEQBX"   ) ) ) .And. !Empty( SFQ->( FieldPos( "FQ_SEQDES"  ) ) ) )       
				 
	Local cPCCBaixa   := iif(lPCCBaixa,"1","0")  
	Local cAdiant		:= IIF( (cTipo $ MVRECANT+"/"+MVPAGANT+"/"+MV_CRNEG+"/"+MV_CPNEG), "1","0")

	DEFAULT nTxMoeda := 0
	
	dDataBaixa  := iif(dDataBaixa ==nil, dDataBase, dDataBaixa )
	dData       := iif(dData      ==nil, dDataBase, dData )
	nMoeda      := iif(nMoeda     ==nil, 1        , nMoeda )
	cLoja       := iif(cLoja      ==nil, iif(cCart=="R",SE1->E1_LOJA   ,SE2->E2_LOJA   ), cLoja )
	nTipoData	:= iif(nTipoData  ==nil, 1 , nTipoData )

	If nTipoData == 1
		cTipoData := "0"  // E5_DATA
	ElseIf nTipodata == 2
		cTipoData := "1"  // E5_DTDISPO
	Else
		cTipoData := "2"  // E5_DTDIGIT
	Endif


	If ExistProc( cProcedure ) .and. ( TcSrvType() <> "AS/400" )
	    //Alert("Executando Procedure 'FIN002' IF 1  nSld1MPag "+Alltrim(Str(nSld1MPag)))
		aResult := {}

		cCliFor:=Iif( cCliFor=NIL,Iif(cCart=="R",SE1->E1_CLIENTE,SE2->E2_FORNECE),cCliFor )
		cLoja  :=Iif( cLoja  =NIL,Iif(cCart=="R",SE1->E1_LOJA   ,SE2->E2_LOJA   ),cLoja   )

		If cCart = "R"
			cAliasTit := "SE1"
			dbSelectArea("SE1")
			nSaldo    := E1_VALOR+SE1->E1_SDACRES-SE1->E1_SDDECRE  
			nMoedaTit := SE1->E1_MOEDA
			cCliFor   := Iif(Empty(cCliFor),SE1->E1_CLIENTE,cCliFor)
			cLoja     := Iif(Empty(cLoja  ),SE1->E1_LOJA,cLoja)
		Else
			cAliasTit := "SE2"
			dbSelectArea("SE2")
			nSaldo    := E2_VALOR+SE2->E2_SDACRES-SE2->E2_SDDECRE  
			nMoedaTit := SE2->E2_MOEDA
		Endif

		nMoeda    := ((nMoeda+1.00)-1.00)
		nMoedaTit := ((nMoedaTit+1.00)-1.00)

		aResult := TCSPEXEC( xProcedures(cProcedure),;
			cPrefixo,                cNumero,;
			cParcela,                cTipo,;
			cCliFor,                 DTOS(dData),;
			DTOS(dDataBaixa),        cLoja,;
			DTOS(dDataBase),         cFilAnt,;
			nSaldo,                  nMoedaTit,;
			cPaisLoc,                cTipoData,;
			cPCCBaixa,               cCart, cAdiant )

		nSaldo := aResult[1]
		// Zera o Saldo devido problema de arredondamento nos juros, ou seja, o valor dos juros que eh gravado com
		// 2 casas decimais, gera diferena na recomposicao do saldo no titulo
		// Exemplo: Titulo com valor de 24.450, com E1_PORCJUR de 0.13 e tres dias de atraso, grava em E5_JUROS o valor 
		// de 95.36, sendo que o valor dos juros seria 95.355
		// Movimentacao no SE5:
		//	      Baixa	Juros	       Saldo
		//		 		            24.450,00
		//-------------------------------
		//		4.001,04	95,36 	20.544,32 3 dias apos vencto.
		//		2.100,95		      18.443,37 mesma data
		//		3.474,23		      14.969,14 mesma data
		//		6.000,00		       8.969,14 5 dias apos vencto
		//		5.060,00		       3.909,14 10 dias apos vencto
		//		3.919,29	10,16	        0,01 12 dias apos vencto
		If Empty((cAliasTit)->&(Right(cAliasTit,2)+"_SALDO")) .And. Abs(nSaldo) <= 0.01
			nSaldo := 0
		Else
			nSaldo := xMoeda(nSaldo,nMoedaTit,nMoeda,dData,,nTxMoeda)
		Endif
		Return (nSaldo)
	Elseif ExistProc( cProcedure ) .and. ( TcSrvType() == "AS/400" )
	    //Alert("Executando Procedure 'FIN002' IF 22222222  nSld1MPag "+Alltrim(Str(nSld1MPag)))
		aResult := {}
		cxFilial := BuildStrFil("SE5")

		cCliFor:=Iif( cCliFor=NIL,Iif(cCart=="R",SE1->E1_CLIENTE,SE2->E2_FORNECE),cCliFor )
		cLoja  :=Iif( cLoja  =NIL,Iif(cCart=="R",SE1->E1_LOJA   ,SE2->E2_LOJA   ),cLoja   )

		If cCart = "R"
			dbSelectArea("SE1")
			nSaldo    := E1_VALOR+SE1->E1_SDACRES-SE1->E1_SDDECRE  
			nMoedaTit := SE1->E1_MOEDA
			cCliFor   := Iif(Empty(cCliFor),SE1->E1_CLIENTE,cCliFor)
			cLoja     := Iif(Empty(cLoja  ),SE1->E1_LOJA,cLoja)
		Else
			dbSelectArea("SE2")
			nSaldo    := E2_VALOR+SE2->E2_SDACRES-SE2->E2_SDDECRE  
			nMoedaTit := SE2->E2_MOEDA
		Endif

		nMoeda    := ((nMoeda+1.00)-1.00)
		nMoedaTit := ((nMoedaTit+1.00)-1.00)

		aResult := TCSPEXEC( xProcedures(cProcedure), cxFilial,;
			cPrefixo,                cNumero,;
			cParcela,                cTipo,;
			cNatureza,               cCart,;
			cCliFor,                 nMoeda,;
			DTOS(dData),             DTOS(dDataBaixa),;
			cLoja,                   DTOS(dDataBase),;
			cFilAnt,                 '',;
			nSaldo,                  nMoedaTit, cTipoData)

		nSaldo := aResult[1]
		// Zera o Saldo devido problema de arredondamento nos juros, ou seja, o valor dos juros que eh gravado com
		// 2 casas decimais, gera diferena na recomposicao do saldo no titulo
		// Exemplo: Titulo com valor de 24.450, com E1_PORCJUR de 0.13 e tres dias de atraso, grava em E5_JUROS o valor 
		// de 95.36, sendo que o valor dos juros seria 95.355
		// Movimentacao no SE5:
		//	      Baixa	Juros	       Saldo
		//		 		            24.450,00
		//-------------------------------
		//		4.001,04	95,36 	20.544,32 3 dias apos vencto.
		//		2.100,95		      18.443,37 mesma data
		//		3.474,23		      14.969,14 mesma data
		//		6.000,00		       8.969,14 5 dias apos vencto
		//		5.060,00		       3.909,14 10 dias apos vencto
		//		3.919,29	10,16	        0,01 12 dias apos vencto
		If Empty((cAliasTit)->&(Right(cAliasTit,2)+"_SALDO")) .And. Abs(nSaldo) <= 0.009
			nSaldo := 0
		Else
			nSaldo := xMoeda(nSaldo,nMoedaTit,nMoeda,dData,,nTxMoeda)
		Endif
		Return (nSaldo)
	Else
	    //Alert("Executando Fun��o TitxSaldo "+Alltrim(Str(nSld1MPag)))
		Return TitxSaldo(@cPrefixo,@cNumero,@cParcela,@cTipo,@cNatureza,@cCart,@cCliFor,@nMoeda,@dData,@dDataBaixa,@cLoja,@cFilTit,nTxMoeda,nTipoData)
	Endif

	/*/
	�����������������������������������������������������������������������������
	�����������������������������������������������������������������������������
	�������������������������������������������������������������������������Ŀ��
	���Fun��o    � SaldoTit � Autor � Wagner Xavier         � Data � 20/08/93 ���
	�������������������������������������������������������������������������Ĵ��
	���Descri��o � Calcula o valor de um titulo em uma determinada data       ���
	�������������������������������������������������������������������������Ĵ��
	���Sintaxe   � SaldoTit(ExpC1,ExpC2,ExpC3,ExpC4,ExpC5,ExpC6,ExpC7,ExpN1)  ���
	�������������������������������������������������������������������������Ĵ��
	���Parametros� ExpC1=Numero do Prefixo                                    ���
	���          � ExpC2=Numero do Titulo                                     ���
	���          � ExpC3=Parcela                                              ���
	���          � ExpC4=Tipo                                                 ���
	���          � ExpC5=Natureza                                             ���
	���          � ExpC6=Carteira  (R/P)                                      ���
	���          � ExpC7=Fornecedor (se ExpC6 = 'R')                          ���
	���          � ExpN1=Moeda                                                ���
	���          � ExpD1=Data para conversao                                  ���
	���          � ExpD2=Data data baixa a ser considerada (retroativa)       ���
	���          � ExpC8=Loja do titulo                                       ���
	���          � ExpC9=Filial do titulo                                     ���
	���          � ExpX10=Tipo de data para compor saldo (baixa/dispo/digit)  ���
	��������������������������������������������������������������������������ٱ�
	�����������������������������������������������������������������������������
	�����������������������������������������������������������������������������
	*/

	Static Function TitxSaldo(cPrefixo,cNumero,cParcela,cTipo,cNatureza,cCart,cCliFor,nMoeda,;
							 dData,dDataBaixa,cLoja,cFilTit,nTxMoeda,nTipoData)
#ENDIF

//Tipos de Data (cTipoData)
// 0 = Data Da Baixa (E5_DATA)
// 1 = Data de Disponibilidade (E5_DTDISPO)
// 2 = Data de Contabilida��o (E5_DTDIGIT)

LOCAL cAlias:=Alias(),nSaldo:=0,dDataMoeda, nMoedaTit
LOCAL cCarteira
LOCAL dDtFina 
Local cAliasTit

//Controla o Pis Cofins e Csll na baixa
Local lPCCBaixa := SuperGetMv("MV_BX10925",.T.,"2") == "1"  .and. (!Empty( SE5->( FieldPos( "E5_VRETPIS" ) ) ) .And. !Empty( SE5->( FieldPos( "E5_VRETCOF" ) ) ) .And. ; 
				 !Empty( SE5->( FieldPos( "E5_VRETCSL" ) ) ) .And. !Empty( SE5->( FieldPos( "E5_PRETPIS" ) ) ) .And. ;
				 !Empty( SE5->( FieldPos( "E5_PRETCOF" ) ) ) .And. !Empty( SE5->( FieldPos( "E5_PRETCSL" ) ) ) .And. ;
				 !Empty( SE2->( FieldPos( "E2_SEQBX"   ) ) ) .And. !Empty( SFQ->( FieldPos( "FQ_SEQDES"  ) ) ) )


DEFAULT nTxMoeda := 0

cCliFor:=Iif( cCliFor=NIL,Iif(cCart=="R",SE1->E1_CLIENTE,SE2->E2_FORNECE),cCliFor )
cLoja  :=Iif( cLoja  =NIL,Iif(cCart=="R",SE1->E1_LOJA   ,SE2->E2_LOJA   ),cLoja   )
nMoeda :=IIF( nMoeda==NIL,1,nMoeda )
dDataMoeda :=IIF( dData==NIL,dDataBase,dData )
dDataBaixa :=IIF( dDataBaixa==NIL,dDataBase,dDataBaixa )
nTipoData  := IIF( nTipoData  ==nil, 0 , nTipoData )

If nTipoData == 1
	cTipoData := "0"  // E5_DATA
ElseIf nTipodata == 2
	cTipoData := "1"  // E5_DTDISPO
Else
	cTipoData := "2"  // E5_DTDIGIT
Endif


// cFiltit somente e' usado no caso de relatorios que podem ser tirados
// por empresa (opcional)
cFilTit := Iif(cFilTit==Nil,xFilial("SE5"),cFilTit)

dbSelectArea("SE5")
If xFilial() == "  "
	cFilTit := "  "
Endif

If cCart = "R"
	cAliasTit := "SE1"
	dbSelectArea("SE1")
	nSaldo := E1_VALOR+SE1->E1_SDACRES-SE1->E1_SDDECRE  
	nMoedaTit := SE1->E1_MOEDA
	cCliFor := Iif(Empty(cCliFor),SE1->E1_CLIENTE,cCliFor)
	cLoja   := Iif(Empty(cLoja  ),SE1->E1_LOJA,cLoja)
Else
	cAliasTit := "SE2"
	dbSelectArea("SE2")
	nSaldo := SE2->E2_VALOR+SE2->E2_SDACRES-SE2->E2_SDDECRE  
	nMoedaTit := SE2->E2_MOEDA
Endif

If Select("__BAIXA") == 0
	ChkFile("SE5",.F.,"__BAIXA")
Else
	dbSelectArea("__BAIXA")
Endif
dbSetOrder(7)
dbSeek(cFilTit+cPrefixo+cNumero+cParcela+cTipo)

While E5_FILIAL + E5_PREFIXO+E5_NUMERO+E5_PARCELA+E5_TIPO == ;
		cFilTit + cPrefixo + cNumero + cParcela + cTipo .and. !EOF()

	IF E5_SITUACA = "C"
		dbSkip()
		Loop
	Endif
	//Posicionar o SE5 pois dentro da TEMBXCANC condulta o SE5 e nao o __BAIXA
   SE5->(DbGoTo(__BAIXA->(Recno())))
	If TemBxCanc(E5_PREFIXO+E5_NUMERO+E5_PARCELA+E5_TIPO+E5_CLIFOR+E5_LOJA+E5_SEQ)
		dbskip()
		Loop
	EndIf

	cCarteira := cCart
	If cCart == "R"
		If (E5_TIPO$MVRECANT+"/"+MV_CRNEG).and. E5_TIPODOC $ "BA|VL|MT|JR|CM" ;
				.and. Empty(E5_DOCUMEN) .and. IIF(!(MovBcoBx(E5_MOTBX)), .T. , !Empty(E5_BANCO) )
			cCarteira := "P"        //Baixa de adiantamento (inverte)
			cLoja     := SE1->E1_LOJA
		Endif
	Endif

	If cCart == "P"
		If (E5_TIPO$MVPAGANT+"/"+MV_CPNEG).and. E5_TIPODOC $ "BA|VL|MT|JR|CM" .and. ;
				Empty(E5_DOCUMEN) .and. IIF(!(MovBcoBx(E5_MOTBX)), .T. , !Empty(E5_BANCO) )
			cCarteira := "R"        //Baixa de adiantamento (inverte)
		Endif
	Endif

	IF (cCarteira == "P" .and. E5_RECPAG == "R") .or. (cCarteira == "R" .and. E5_RECPAG == "P")
		dbSkip( )
		Loop
	Endif

	IF cCliFor + cLoja  != E5_CLIFOR + E5_LOJA
		dbSkip()
		Loop
	EndIF

	//Defino qual o tipo de data a ser utilizado para compor o saldo do titulo
	If cTipoData == "0"
		dDtFina := E5_DATA
	ElseIf cTipoData == "1"
		dDtFina := E5_DTDISPO
	Else	
		dDtFina := E5_DTDIGIT
	Endif			

	IF dDtFina <= dDataBaixa
		IF E5_TIPODOC $ "VL#BA#V2#CP#LJ"
			nSaldo -= IIF((nMoedaTit < 2.And.cPaisLoc=="BRA").Or. (cPaisLoc <>"BRA" .And. nMoedaTit > 1 .And. !Empty(E5_BANCO)),E5_VALOR,E5_VLMOED2)
			nSaldo += Round(NoRound(xMoeda(E5_VLMULTA+E5_VLJUROS-E5_VLDESCO,1,nMoedaTit,E5_DATA,3,,nTxMoeda),3),2)
			//Retencao de impostos na baixa
			If lPccBaixa .and. cCarteira == "P" .and. Empty(E5_PRETPIS) .and. !(E5_MOTBX == "PCC")
				nSaldo -= Round(NoRound(xMoeda(E5_VRETPIS+E5_VRETCOF+E5_VRETCSLL,1,nMoedaTit,E5_DATA,3,,nTxMoeda),3),2)				
			Endif			
			If nSaldo <= 0.009
				nSaldo := 0
			Endif
		Endif
	EndIF
	dbSkip()
Enddo
// Zera o Saldo devido problema de arredondamento nos juros, ou seja, o valor dos juros que eh gravado com
// 2 casas decimais, gera diferena na recomposicao do saldo no titulo
// Exemplo: Titulo com valor de 24.450, com E1_PORCJUR de 0.13 e tres dias de atraso, grava em E5_JUROS o valor 
// de 95.36, sendo que o valor dos juros seria 95.355
// Movimentacao no SE5:
//	      Baixa	Juros	       Saldo
//		 		            24.450,00
//-------------------------------
//		4.001,04	95,36 	20.544,32 3 dias apos vencto.
//		2.100,95		      18.443,37 mesma data
//		3.474,23		      14.969,14 mesma data
//		6.000,00		       8.969,14 5 dias apos vencto
//		5.060,00		       3.909,14 10 dias apos vencto
//		3.919,29	10,16	        0,01 12 dias apos vencto
If Empty((cAliasTit)->&(Right(cAliasTit,2)+"_SALDO")) .And. Abs(nSaldo) <= 0.009
	nSaldo := 0
Else
	nSaldo := xMoeda(nSaldo,nMoedaTit,nMoeda,dDataMoeda,,nTxMoeda)
Endif
dbSelectArea(cAlias)
Return ( nSaldo )







/************************************** FUN��ES P/ CALCULO DE SALDOS DE TITULOS *********************************/







Static Function SaldoSR(dDtParam,d2DtParam,d3DtParam,nChamada)

Private dRGDBase:= dDtParam
Private dDoVenc := d2DtParam
Private dAteVenc:= d3DtParam
Private nFunCham:= nChamada
Private nTotFil4 := 0

//Alert("Estou em SaldoSR paramentos recebidos, DataBase "+DTOC(dRGDBase)+" Vencto De "+dtoc(d2DtParam)+" Vencto Ate "+DTOC(d3DtParam))
Private nRetValor := nSldAR()
Return(nRetValor) 
//Endif

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 � FINR130R3� Autor � Paulo Boschetti	     � Data � 01.06.92 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Posi��o dos Titulos a Receber 						 	  		  ���
�������������������������������������������������������������������������Ĵ��
���Sintaxe e � FINR130R3(void)									  					  ���
�������������������������������������������������������������������������Ĵ��
��� Uso		 � Generico 												  				  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
����������������������������������������������������������������������������
/*/
Static Function nSldAR()
Local cDesc1 :=OemToAnsi(STR0001)  //"Imprime a posi��o dos titulos a receber relativo a data ba-"
Local cDesc2 :=OemToAnsi(STR0002)  //"se do sistema."
Local cDesc3 :=""
Local wnrel
Local cString:="SE1"
Local nRegEmp:=SM0->(RecNo())
Local dOldDtBase := dDataBase
Local dOldData	:= dDatabase

//Alert("Estou em nSldAR")

Private titulo  :=""
Private cabec1  :=""
Private cabec2  :=""

Private aLinha  :={}
Private aReturn :={ OemToAnsi(STR0003), 1,OemToAnsi(STR0004), 1, 2, 1, "",1 }  //"Zebrado"###"Administracao"
Private cPerg	 :="FIN130"
Private nJuros  :=0
Private nLastKey:=0
Private nomeprog:="FINR130"
Private tamanho :="G"

//��������������������������Ŀ
//� Defini��o dos cabe�alhos �
//����������������������������
titulo := OemToAnsi(STR0005)  //"Posicao dos Titulos a Receber"
cabec1 := OemToAnsi(STR0006)  //"Codigo Nome do Cliente      Prf-Numero         TP  Natureza    Data de  Vencto   Vencto  Banco  Valor Original |        Titulos Vencidos          | Titulos a Vencer | Num        Vlr.juros ou  Dias   Historico     "
cabec2 := OemToAnsi(STR0007)  //"                            Parcela                            Emissao  Titulo    Real                         |  Valor Nominal   Valor Corrigido |   Valor Nominal  | Banco       permanencia  Atraso               "

//Nao retire esta chamada. Verifique antes !!!
//Ela � necessaria para o correto funcionamento da pergunte 36 (Data Base)
//AjustaSx1()
//PutDtBase()

//pergunte("FIN130",.F.)

mv_par01	:= SPACE(6)	 // Do Cliente 													   �
mv_par02	:= "zzzzzz"	 // Ate o Cliente													   �
mv_par03    := SPACE(3)		 // Do Prefixo														   �
mv_par04    := "zzz"		 // Ate o prefixo 												   �
mv_par05    := "         "		 // Do Titulo													      �
mv_par06	:= "zzzzzzzzz"	 // Ate o Titulo													   �
mv_par07	:= SPACE(3)	 // Do Banco														   �
mv_par08	:= "zzz"	 // Ate o Banco													   �
mv_par09	:= dDoVenc	 // Do Vencimento 												   �
mv_par10	:= dAteVenc	 // Ate o Vencimento												   �
mv_par11	:= "          "	 // Da Natureza														�
mv_par12	:= "zzzzzzzzzz"	 // Ate a Natureza													�
mv_par13	:= CTOD("01/01/80")	 // Da Emissao															�
mv_par14	:= CTOD("31/12/49")	 // Ate a Emissao														�
mv_par15	:= 1	 // Qual Moeda															�
mv_par16	:= 2	 // Imprime provisorios												�
mv_par17	:= 2	 // Reajuste pelo vecto												�
mv_par18    := 2		 // Impr Tit em Descont												�
mv_par19	:= 2	 // Relatorio Anal/Sint												�
mv_par20	:= 1	 // Consid Data Base?  												�
mv_par21	:= 1	 // Consid Filiais  ?  												�
mv_par22	:= SM0->M0_CODFIL	 // da filial													      �
mv_par23	:= SM0->M0_CODFIL	 // a flial 												         �
mv_par24	:= "  "	 // Da loja  															�
mv_par25	:= "zz"	 // Ate a loja															�
mv_par26	:= 2	 // Consid Adiantam.?												�
mv_par27	:= CTOD("01/01/94")	 // Da data contab. ?												�
mv_par28	:= CTOD("31/12/49")	 // Ate data contab.?												�
mv_par29	:= 2	 // Imprime Nome    ?												�
mv_par30	:= 2	 // Outras Moedas   ?												�
mv_par31    := "                                        "  // Imprimir os Tipos												�
mv_par32    := "NCC;AB-                                 "   // Nao Imprimir Tipos												�
mv_par33    := 3   // Abatimentos  - Lista/Nao Lista/Despreza					�
mv_par34    := 2  // Consid. Fluxo Caixa												�
mv_par35    := 2   // Salta pagina Cliente											�
mv_par36    := dRGDBase   // Data Base													      �
mv_par37    := 1  // Compoe Saldo por: Data da Baixa, Credito ou DtDigit  �
MV_PAR38    := 2  // Tit. Emissao Futura												�
MV_PAR39    := 2  // Converte Valores 												�

//Alert("Definido os Parametros")

//������������������������������������������������������������������������Ŀ
//� Variaveis utilizadas para parametros												�
//� mv_par01		 // Do Cliente 													   �
//� mv_par02		 // Ate o Cliente													   �
//� mv_par03		 // Do Prefixo														   �
//� mv_par04		 // Ate o prefixo 												   �
//� mv_par05		 // Do Titulo													      �
//� mv_par06		 // Ate o Titulo													   �
//� mv_par07		 // Do Banco														   �
//� mv_par08		 // Ate o Banco													   �
//� mv_par09		 // Do Vencimento 												   �
//� mv_par10		 // Ate o Vencimento												   �
//� mv_par11		 // Da Natureza														�
//� mv_par12		 // Ate a Natureza													�
//� mv_par13		 // Da Emissao															�
//� mv_par14		 // Ate a Emissao														�
//� mv_par15		 // Qual Moeda															�
//� mv_par16		 // Imprime provisorios												�
//� mv_par17		 // Reajuste pelo vecto												�
//� mv_par18		 // Impr Tit em Descont												�
//� mv_par19		 // Relatorio Anal/Sint												�
//� mv_par20		 // Consid Data Base?  												�
//� mv_par21		 // Consid Filiais  ?  												�
//� mv_par22		 // da filial													      �
//� mv_par23		 // a flial 												         �
//� mv_par24		 // Da loja  															�
//� mv_par25		 // Ate a loja															�
//� mv_par26		 // Consid Adiantam.?												�
//� mv_par27		 // Da data contab. ?												�
//� mv_par28		 // Ate data contab.?												�
//� mv_par29		 // Imprime Nome    ?												�
//� mv_par30		 // Outras Moedas   ?												�
//� mv_par31       // Imprimir os Tipos												�
//� mv_par32       // Nao Imprimir Tipos												�
//� mv_par33       // Abatimentos  - Lista/Nao Lista/Despreza					�
//� mv_par34       // Consid. Fluxo Caixa												�
//� mv_par35       // Salta pagina Cliente											�
//� mv_par36       // Data Base													      �
//� mv_par37       // Compoe Saldo por: Data da Baixa, Credito ou DtDigit  �
//� MV_PAR38       // Tit. Emissao Futura												�
//� MV_PAR39       // Converte Valores 												�
//��������������������������������������������������������������������������
//���������������������������������������Ŀ
//� Envia controle para a fun��o SETPRINT �
//�����������������������������������������

/*
wnrel:="FINR130"            //Nome Default do relatorio em Disco
aOrd :={	OemToAnsi(STR0008),;	//"Por Cliente"
	OemToAnsi(STR0009),;	//"Por Prefixo/Numero"
	OemToAnsi(STR0010),; //"Por Banco"
	OemToAnsi(STR0011),;	//"Por Venc/Cli"
	OemToAnsi(STR0012),;	//"Por Natureza"
	OemToAnsi(STR0013),; //"Por Emissao"
	OemToAnsi(STR0014),;	//"Por Ven\Bco"
	OemToAnsi(STR0015),; //"Por Cod.Cli."
	OemToAnsi(STR0016),; //"Banco/Situacao"
	OemToAnsi(STR0047) } //"Por Numero/Tipo/Prefixo"

wnrel:=SetPrint(cString,wnrel,cPerg,@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,,Tamanho)

If nLastKey == 27
	Return
Endif

SetDefault(aReturn,cString)

If nLastKey == 27
	Return
Endif

RptStatus({|lEnd| FA130Imp(@lEnd,wnRel,cString)},titulo)  // Chamada do Relatorio
*/
//Alert("Irei chamar CalcSAR()")

nValRet := CalcSAR()
SM0->(dbGoTo(nRegEmp))
cFilAnt := SM0->M0_CODFIL

//Acerta a database de acordo com a database real do sistema
If mv_par20 == 1    // Considera Data Base
	dDataBase := dOldDtBase
Endif	
//nValRet := nTot2+nTot3
//nValRet := nTotFil4
//Alert("Retornando o Valor "+Alltrim(Str(nValRet)))
//Alert("Retorno da nSldAR - nTotFil4 "+Alltrim(Str(nTotFil4))+" nValRet "+Alltrim(Str(nValRet)))
Return(nValRet)

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 � FA130Imp � Autor � Paulo Boschetti		  � Data � 01.06.92 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Imprime relat�rio dos T�tulos a Receber						  ���
�������������������������������������������������������������������������Ĵ��
���Sintaxe e � FA130Imp(lEnd,WnRel,cString)										  ���
�������������������������������������������������������������������������Ĵ��
���Parametros� lEnd	  - A��o do Codeblock				    					  ���
���			 � wnRel   - T�tulo do relat�rio 									  ���
���			 � cString - Mensagem													  ���
�������������������������������������������������������������������������Ĵ��
��� Uso		 � Generico 																  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/
Static Function CalcSAR() //FA130Imp(lEnd,WnRel,cString)
Local CbCont
Local CbTxt
Local nOrdem
Local lContinua := .T.
Local cCond1,cCond2,cCarAnt
Local nTit0:=0
Local nTit1:=0
Local nTit2:=0
Local nTit3:=0
Local nTit4:=0
Local nTit5:=0
Local nTotJ:=0
Local nTot0:=0
Local nTot1:=0
Local nTot2:=0
Local nTot3:=0
Local nTot4:=0
Local nTotTit:=0
Local nTotJur:=0
Local nTotFil0:=0
Local nTotFil1:=0
Local nTotFil2:=0
Local nTotFil3:=0
//Local nTotFil4:=0
Local nTotFilTit:=0
Local nTotFilJ:=0
Local nAtraso:=0
Local nTotAbat:=0
Local nSaldo:=0
Local dDataReaj
Local dDataAnt := dDataBase
Local lQuebra
Local nMesTit0 := 0
Local nMesTit1 := 0
Local nMesTit2 := 0
Local nMesTit3 := 0
Local nMesTit4 := 0
Local nMesTTit := 0
Local nMesTitj := 0
Local cIndexSe1
Local cChaveSe1
Local nIndexSE1
Local dDtContab
Local cTipos  := ""
#IFDEF TOP
	Local aStru := SE1->(dbStruct()), ni
#ENDIF	
Local nTotsRec := SE1->(RecCount())
Local aTamCli  := TAMSX3("E1_CLIENTE")
Local lF130Qry := ExistBlock("F130QRY")
// variavel  abaixo criada p/pegar o nr de casas decimais da moeda
Local ndecs := Msdecimais(mv_par15)
Local nAbatim := 0
Local nDescont:= 0
Local nVlrOrig:= 0

//Alert("Estou na CalcSAR()")
Private nVoltaVlr := 0

Private nRegSM0 := SM0->(Recno())
Private nAtuSM0 := SM0->(Recno())
PRIVATE dBaixa := dDataBase
PRIVATE cFilDe,cFilAte

//��������������������������������������������������������������Ŀ
//� Ponto de entrada para Filtrar os tipos sem entrar na tela do �
//� FINRTIPOS(), localizacao Argentina.                          �
//����������������������������Jose Lucas, Localiza��es Argentina��
IF EXISTBLOCK("F130FILT")
	cTipos	:=	EXECBLOCK("F130FILT",.f.,.f.)
ENDIF

nOrdem:=aReturn[8]
cMoeda:=Str(mv_par15,1)

//�����������������������������������������������������������Ŀ
//� Vari�veis utilizadas para Impress�o do Cabe�alho e Rodap� �
//�������������������������������������������������������������
cbtxt 	:= OemtoAnsi(STR0046)
cbcont	:= 1
li 		:= 80
m_pag 	:= 1

//������������������������������������������������������������������Ŀ
//� POR MAIS ESTRANHO QUE PARE�A, ESTA FUNCAO DEVE SER CHAMADA AQUI! �
//�                                                                  �
//� A fun��o SomaAbat reabre o SE1 com outro nome pela ChkFile para  �
//� efeito de performance. Se o alias auxiliar para a SumAbat() n�o  �
//� estiver aberto antes da IndRegua, ocorre Erro de & na ChkFile,   �
//� pois o Filtro do SE1 uptrapassa 255 Caracteres.                  �
//��������������������������������������������������������������������
SomaAbat("","","","R")

//Alert("Passei pelo SomaAbat")

//�����������������������������������������������������������Ŀ
//� Atribui valores as variaveis ref a filiais                �
//�������������������������������������������������������������
If mv_par21 == 2
	cFilDe  := cFilAnt
	cFilAte := cFilAnt
ELSE
	cFilDe := mv_par22	// Todas as filiais
	cFilAte:= mv_par23
Endif

//Acerta a database de acordo com o parametro
If mv_par20 == 1    // Considera Data Base
	dDataBase := mv_par36
Endif	

//Alert("Preparando para entrar no While")
dbSelectArea("SM0")
cEmpAnt := SM0->M0_CODIGO
cFilAte := SM0->M0_CODFIL

dbSeek(cEmpAnt+cFilDe,.T.)

nRegSM0 := SM0->(Recno())
nAtuSM0 := SM0->(Recno())

While !Eof() .and. M0_CODIGO == cEmpAnt .and. M0_CODFIL <= cFilAte
	
	dbSelectArea("SE1")
	cFilAnt := SM0->M0_CODFIL
	Set Softseek On
//	Alert("Entrou no While do SM0")
	/*
	If mv_par19 == 1
		titulo := titulo + OemToAnsi(STR0026)  //" - Analitico"
	Else
		titulo := titulo + OemToAnsi(STR0027)  //" - Sintetico"
		cabec1 := OemToAnsi(STR0044)  //"                                                                                                               |        Titulos Vencidos          | Titulos a Vencer |            Vlr.juros ou             (Vencidos+Vencer)"
		cabec2 := OemToAnsi(STR0045)  //"                                                                                                               |  Valor Nominal   Valor Corrigido |   Valor Nominal  |             permanencia                              "
	EndIf
	*/
	#IFDEF TOP
		
		If nOrdem = 1
			cQuery := ""
			aEval(SE1->(DbStruct()),{|e| If(!Alltrim(e[1])$"E1_FILIAL#E1_NOMCLI#E1_CLIENTE#E1_LOJA#E1_PREFIXO#E1_NUM#E1_PARCELA#E1_TIPO", cQuery += ","+AllTrim(e[1]),Nil)})
			cQuery := "SELECT E1_FILIAL, E1_NOMCLI, E1_CLIENTE, E1_LOJA, E1_PREFIXO, E1_NUM,E1_PARCELA, E1_TIPO, "+ SubStr(cQuery,2)
		Else
			cQuery := "SELECT * "
		EndIf
		
		cQuery += "  FROM "+	RetSqlName("SE1") + " SE1"
		cQuery += " WHERE E1_FILIAL = '" + xFilial("SE1") + "'"
		cQuery += "   AND D_E_L_E_T_ <> '*' "
	#ENDIF
	
	IF nOrdem = 1
		cChaveSe1 := "E1_FILIAL+E1_NOMCLI+E1_CLIENTE+E1_LOJA+E1_PREFIXO+E1_NUM+E1_PARCELA+E1_TIPO"
		#IFDEF TOP
			cOrder := SqlOrder(cChaveSe1)
		#ELSE
			cIndexSe1 := CriaTrab(nil,.f.)
			IndRegua("SE1",cIndexSe1,cChaveSe1,,Fr130IndR(),OemToAnsi(STR0022))
			nIndexSE1 := RetIndex("SE1")
			dbSetIndex(cIndexSe1+OrdBagExt())
			dbSetOrder(nIndexSe1+1)
			dbSeek(xFilial("SE1"))
		#ENDIF
		cCond1 := "E1_CLIENTE <= mv_par02"
		cCond2 := "E1_CLIENTE + E1_LOJA"
		titulo := titulo + OemToAnsi(STR0017)  //" - Por Cliente"
		
	Elseif nOrdem = 2
		SE1->(dbSetOrder(1))
		#IFNDEF TOP
			dbSeek(cFilial+mv_par03+mv_par05)
		#ELSE
			cOrder := SqlOrder(IndexKey())
		#ENDIF
		cCond1 := "E1_NUM <= mv_par06"
		cCond2 := "E1_NUM"
		titulo := titulo + OemToAnsi(STR0018)  //" - Por Numero"
	Elseif nOrdem = 3
		SE1->(dbSetOrder(4))
		#IFNDEF TOP
			dbSeek(cFilial+mv_par07)
		#ELSE
			cOrder := SqlOrder(IndexKey())
		#ENDIF
		cCond1 := "E1_PORTADO <= mv_par08"
		cCond2 := "E1_PORTADO"
		titulo := titulo + OemToAnsi(STR0019)  //" - Por Banco"
	Elseif nOrdem = 4
		SE1->(dbSetOrder(7))
		#IFNDEF TOP
			dbSeek(cFilial+DTOS(mv_par09))
		#ELSE
			cOrder := SqlOrder(IndexKey())
		#ENDIF
		cCond1 := "E1_VENCREA <= mv_par10"
		cCond2 := "E1_VENCREA"
		titulo := titulo + OemToAnsi(STR0020)  //" - Por Data de Vencimento"
	Elseif nOrdem = 5
		SE1->(dbSetOrder(3))
		#IFNDEF TOP
			dbSeek(cFilial+mv_par11)
		#ELSE
			cOrder := SqlOrder(IndexKey())
		#ENDIF
		cCond1 := "E1_NATUREZ <= mv_par12"
		cCond2 := "E1_NATUREZ"
		titulo := titulo + OemToAnsi(STR0021)  //" - Por Natureza"
	Elseif nOrdem = 6
		SE1->(dbSetOrder(6))
		#IFNDEF TOP
			dbSeek( cFilial+DTOS(mv_par13))
		#ELSE
			cOrder := SqlOrder(IndexKey())
		#ENDIF
		cCond1 := "E1_EMISSAO <= mv_par14"
		cCond2 := "E1_EMISSAO"
		titulo := titulo + OemToAnsi(STR0042)  //" - Por Emissao"
	Elseif nOrdem == 7
		cChaveSe1 := "E1_FILIAL+DTOS(E1_VENCREA)+E1_PORTADO+E1_PREFIXO+E1_NUM+E1_PARCELA+E1_TIPO"
		#IFNDEF TOP
			cIndexSe1 := CriaTrab(nil,.f.)
			IndRegua("SE1",cIndexSe1,cChaveSe1,,Fr130IndR(),OemToAnsi(STR0022))
			nIndexSE1 := RetIndex("SE1")
			dbSetIndex(cIndexSe1+OrdBagExt())
			dbSetOrder(nIndexSe1+1)
			dbSeek(xFilial("SE1"))
		#ELSE
			cOrder := SqlOrder(cChaveSe1)
		#ENDIF
		cCond1 := "E1_VENCREA <= mv_par10"
		cCond2 := "DtoS(E1_VENCREA)+E1_PORTADO"
		titulo := titulo + OemToAnsi(STR0023)  //" - Por Vencto/Banco"
	Elseif nOrdem = 8
		SE1->(dbSetOrder(2))
		#IFNDEF TOP
			dbSeek(cFilial+mv_par01,.T.)
		#ELSE
			cOrder := SqlOrder(IndexKey())
		#ENDIF
		cCond1 := "E1_CLIENTE <= mv_par02"
		cCond2 := "E1_CLIENTE"
		titulo := titulo + OemToAnsi(STR0024)  //" - Por Cod.Cliente"
	Elseif nOrdem = 9
		cChave := "E1_FILIAL+E1_PORTADO+E1_SITUACA+E1_NOMCLI+E1_PREFIXO+E1_NUM+E1_PARCELA+E1_TIPO"
		#IFNDEF TOP
			dbSelectArea("SE1")
			cIndex := CriaTrab(nil,.f.)
			IndRegua("SE1",cIndex,cChave,,fr130IndR(),OemToAnsi(STR0022))
			nIndex := RetIndex("SE1")
			dbSetIndex(cIndex+OrdBagExt())
			dbSetOrder(nIndex+1)
			dbSeek(xFilial("SE1"))
		#ELSE
			cOrder := SqlOrder(cChave)
		#ENDIF
		cCond1 := "E1_PORTADO <= mv_par08"
		cCond2 := "E1_PORTADO+E1_SITUACA"
		titulo := titulo + OemToAnsi(STR0025)  //" - Por Banco e Situacao"
	ElseIf nOrdem == 10
		cChave := "E1_FILIAL+E1_NUM+E1_TIPO+E1_PREFIXO+E1_PARCELA"
		#IFNDEF TOP
			dbSelectArea("SE1")
			cIndex := CriaTrab(nil,.f.)
			IndRegua("SE1",cIndex,cChave,,,OemToAnsi(STR0022))
			nIndex := RetIndex("SE1")
			dbSetIndex(cIndex+OrdBagExt())
			dbSetOrder(nIndex+1)
			dbSeek(xFilial("SE1")+mv_par05)
		#ELSE
			cOrder := SqlOrder(cChave)
		#ENDIF
		cCond1 := "E1_NUM <= mv_par06"
		cCond2 := "E1_NUM"
		titulo := titulo + OemToAnsi(STR0048)  //" - Numero/Prefixo"	
	Endif
	
	If mv_par19 == 1
		titulo := titulo + OemToAnsi(STR0026)  //" - Analitico"
	Else
		titulo := titulo + OemToAnsi(STR0027)  //" - Sintetico"
		cabec1 := OemToAnsi(STR0044)  //"Nome do Cliente      |        Titulos Vencidos          | Titulos a Vencer |            Vlr.juros ou             (Vencidos+Vencer)"
		cabec2 := OemToAnsi(STR0045)  //"|  Valor Nominal   Valor Corrigido |   Valor Nominal  |             permanencia                              "
	EndIf
	
	cFilterUser:=aReturn[7]
	Set Softseek Off
	
	#IFDEF TOP
		cQuery += " AND E1_CLIENTE between '" + mv_par01        + "' AND '" + mv_par02 + "'"
		cQuery += " AND E1_PREFIXO between '" + mv_par03        + "' AND '" + mv_par04 + "'"
		cQuery += " AND E1_NUM     between '" + mv_par05        + "' AND '" + mv_par06 + "'"
		cQuery += " AND E1_PORTADO between '" + mv_par07        + "' AND '" + mv_par08 + "'"
		cQuery += " AND E1_VENCREA between '" + DTOS(mv_par09)  + "' AND '" + DTOS(mv_par10) + "'"
		cQuery += " AND (E1_MULTNAT = '1' OR (E1_NATUREZ BETWEEN '"+MV_PAR11+"' AND '"+MV_PAR12+"'))"
		cQuery += " AND E1_EMISSAO between '" + DTOS(mv_par13)  + "' AND '" + DTOS(mv_par14) + "'"
		cQuery += " AND E1_LOJA    between '" + mv_par24        + "' AND '" + mv_par25 + "'"

		If MV_PAR38 == 2 //Nao considerar titulos com emissao futura
			cQuery += " AND E1_EMISSAO <=      '" + DTOS(dDataBase) + "'"
		Endif

		cQuery += " AND ((E1_EMIS1  Between '"+ DTOS(mv_par27)+"' AND '"+DTOS(mv_par28)+"') OR E1_EMISSAO Between '"+DTOS(mv_par27)+"' AND '"+DTOS(mv_par28)+"')"
		If !Empty(mv_par31) // Deseja imprimir apenas os tipos do parametro 31
			cQuery += " AND E1_TIPO IN "+FormatIn(mv_par31,";") 
		ElseIf !Empty(Mv_par32) // Deseja excluir os tipos do parametro 32
			cQuery += " AND E1_TIPO NOT IN "+FormatIn(mv_par32,";")
		EndIf
		If mv_par18 == 2
			cQuery += " AND E1_SITUACA NOT IN ('2','7')"
		Endif
		If mv_par20 == 2
			cQuery += ' AND E1_SALDO <> 0'
		Endif
		If mv_par34 == 1
			cQuery += " AND E1_FLUXO <> 'N'"
		Endif               
        //������������������������������������������������������������������������Ŀ
        //� Ponto de entrada para inclusao de parametros no filtro a ser executado �
        //��������������������������������������������������������������������������
	    If lF130Qry 
			cQuery += ExecBlock("F130QRY",.f.,.f.)
		Endif

		cQuery += " ORDER BY "+ cOrder
		
		cQuery := ChangeQuery(cQuery)
		
		dbSelectArea("SE1")
		dbCloseArea()
		dbSelectArea("SA1")
		
		dbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery), 'SE1', .F., .T.)
		
		For ni := 1 to Len(aStru)
			If aStru[ni,2] != 'C'
				TCSetField('SE1', aStru[ni,1], aStru[ni,2],aStru[ni,3],aStru[ni,4])
			Endif
		Next
	#ENDIF
	//SetRegua(nTotsRec)
	
	If MV_MULNATR .And. nOrdem == 5
		//Finr135R3(cTipos, lEnd, @nTot0, @nTot1, @nTot2, @nTot3, @nTotTit, @nTotJ )
		#IFDEF TOP
			dbSelectArea("SE1")
			dbCloseArea()
			ChKFile("SE1")
			dbSelectArea("SE1")
			dbSetOrder(1)
		#ENDIF
		If Empty(xFilial("SE1"))
			Exit
		Endif
		dbSelectArea("SM0")
		dbSkip()
		Loop
	Endif
//	Alert("Selecionei Registros")
	While &cCond1 .and. !Eof() .and. lContinua .and. E1_FILIAL == xFilial("SE1")
		
		/*
		IF	lEnd
			@PROW()+1,001 PSAY OemToAnsi(STR0028)  //"CANCELADO PELO OPERADOR"
			Exit
		Endif
		*/
		//Alert("Entrei no While da Query do SE1")
		//IncRegua()
		
		Store 0 To nTit1,nTit2,nTit3,nTit4,nTit5
		
		//����������������������������������������Ŀ
		//� Carrega data do registro para permitir �
		//� posterior analise de quebra por mes.   �
		//������������������������������������������
		dDataAnt := If(nOrdem == 6 , SE1->E1_EMISSAO,  SE1->E1_VENCREA)
		
		cCarAnt := &cCond2
		
		While &cCond2==cCarAnt .and. !Eof() .and. lContinua .and. E1_FILIAL == xFilial("SE1")
			
			IF lEnd
				//@PROW()+1,001 PSAY OemToAnsi(STR0028)  //"CANCELADO PELO OPERADOR"
				lContinua := .F.
				Exit
			EndIF
			//Alert("Entrei no While da Impressao")
			//IncRegua()
			
			dbSelectArea("SE1")
			//��������������������������������������������������������������Ŀ
			//� Filtrar com base no Pto de entrada do Usuario...             �
			//����������������������������Jose Lucas, Localiza��es Argentina��
			If !Empty(cTipos)
				If !(SE1->E1_TIPO $ cTipos)
					dbSkip()
					Loop
				Endif
			Endif
			
			 // Tratamento da correcao monetaria para a Argentina
			If  cPaisLoc=="ARG" .And. mv_par15 <> 1  .And.  SE1->E1_CONVERT=='N'
					dbSkip()
					Loop
			Endif
			//��������������������������������������������������������������Ŀ
			//� Considera filtro do usuario                                  �
			//����������������������������������������������������������������
			If !Empty(cFilterUser).and.!(&cFilterUser)
				dbSkip()
				Loop
			Endif
			
			//��������������������������������������������������������������Ŀ
			//� Verifica se titulo, apesar do E1_SALDO = 0, deve aparecer ou �
			//� n�o no relat�rio quando se considera database (mv_par20 = 1) �
			//� ou caso n�o se considere a database, se o titulo foi totalmen�
			//� te baixado.																  �
			//����������������������������������������������������������������
			dbSelectArea("SE1")
			IF !Empty(SE1->E1_BAIXA) .and. Iif(mv_par20 == 2 ,SE1->E1_SALDO == 0 ,;
					IIF(mv_par37 == 1,(SE1->E1_SALDO == 0 .and. SE1->E1_BAIXA <= dDataBase),.F.))
				dbSkip()
				Loop
			EndIF
			
			//������������������������������������������������������Ŀ
			//� Verifica se trata-se de abatimento ou somente titulos�
			//� at� a data base. 									 �
			//��������������������������������������������������������
			IF (SE1->E1_TIPO $ MVABATIM .And. mv_par33 != 1) .Or.;
				(SE1->E1_EMISSAO > dDataBase .and. MV_PAR38 == 2)
				dbSkip()
				Loop
			Endif
			
			//������������������������������������������������������Ŀ
			//� Verifica se ser� impresso titulos provis�rios		 �
			//��������������������������������������������������������
			IF E1_TIPO $ MVPROVIS .and. mv_par16 == 2
				dbSkip()
				Loop
			Endif
			
			//������������������������������������������������������Ŀ
			//� Verifica se ser� impresso titulos de Adiantamento	 �
			//��������������������������������������������������������
			IF SE1->E1_TIPO $ MVRECANT+"/"+MV_CRNEG .and. mv_par26 == 2
				dbSkip()
				Loop
			Endif
			
			// dDtContab para casos em que o campo E1_EMIS1 esteja vazio
			dDtContab := Iif(Empty(SE1->E1_EMIS1),SE1->E1_EMISSAO,SE1->E1_EMIS1)
			
			//����������������������������������������Ŀ
			//� Verifica se esta dentro dos parametros �
			//������������������������������������������
			dbSelectArea("SE1")
			IF SE1->E1_CLIENTE < mv_par01 .OR. SE1->E1_CLIENTE > mv_par02 .OR. ;
				SE1->E1_PREFIXO < mv_par03 .OR. SE1->E1_PREFIXO > mv_par04 .OR. ;
				SE1->E1_NUM	 	 < mv_par05 .OR. SE1->E1_NUM 		> mv_par06 .OR. ;
				SE1->E1_PORTADO < mv_par07 .OR. SE1->E1_PORTADO > mv_par08 .OR. ;
				SE1->E1_VENCREA < mv_par09 .OR. SE1->E1_VENCREA > mv_par10 .OR. ;
				SE1->E1_NATUREZ < mv_par11 .OR. SE1->E1_NATUREZ > mv_par12 .OR. ;
				SE1->E1_EMISSAO < mv_par13 .OR. SE1->E1_EMISSAO > mv_par14 .OR. ;
				SE1->E1_LOJA    < mv_par24 .OR. SE1->E1_LOJA    > mv_par25 .OR. ;
				dDtContab       < mv_par27 .OR. dDtContab       > mv_par28 .OR. ;
				(SE1->E1_EMISSAO > dDataBase .and. MV_PAR38 == 2) .Or. !&(fr130IndR())
				dbSkip()
				Loop
			Endif
			
			If mv_par18 == 2 .and. E1_SITUACA $ "27"
				dbSkip()
				Loop
			Endif
			
			//����������������������������������������Ŀ
			//� Verifica se deve imprimir outras moedas�
			//������������������������������������������
			If mv_par30 == 2 // nao imprime
				if SE1->E1_MOEDA != mv_par15 //verifica moeda do campo=moeda parametro
					dbSkip()
					Loop
				endif
			Endif
			
			// Verifica se existe a taxa na data do vencimento do titulo, se nao existir, utiliza a taxa da database
			If SE1->E1_VENCREA < dDataBase
				If mv_par17 == 2 .And. RecMoeda(SE1->E1_VENCREA,cMoeda) > 0
					dDataReaj := SE1->E1_VENCREA
				Else
					dDataReaj := dDataBase
				EndIf	
			Else
				dDataReaj := dDataBase
			EndIf
						
			If mv_par20 == 1	// Considera Data Base
				//Ita - 01/04/2008 - nSaldo :=SaldoTit(SE1->E1_PREFIXO,SE1->E1_NUM,SE1->E1_PARCELA,SE1->E1_TIPO,SE1->E1_NATUREZ,"R",SE1->E1_CLIENTE,mv_par15,dDataReaj,,SE1->E1_LOJA,,Iif(MV_PAR39==2,Iif(!Empty(SE1->E1_TXMOEDA),SE1->E1_TXMOEDA,RecMoeda(SE1->E1_EMISSAO,SE1->E1_MOEDA)),0),mv_par37)
				nSaldo :=TitSaldo(SE1->E1_PREFIXO,SE1->E1_NUM,SE1->E1_PARCELA,SE1->E1_TIPO,SE1->E1_NATUREZ,"R",SE1->E1_CLIENTE,mv_par15,dDataReaj,,SE1->E1_LOJA,,Iif(MV_PAR39==2,Iif(!Empty(SE1->E1_TXMOEDA),SE1->E1_TXMOEDA,RecMoeda(SE1->E1_EMISSAO,SE1->E1_MOEDA)),0),mv_par37)
				
				// Subtrai decrescimo para recompor o saldo na data escolhida.
				If Str(SE1->E1_VALOR,17,2) == Str(nSaldo,17,2) .And. SE1->E1_DECRESC > 0 .And. SE1->E1_SDDECRE == 0
					nSAldo -= SE1->E1_DECRESC
				Endif
				// Soma Acrescimo para recompor o saldo na data escolhida.
				If Str(SE1->E1_VALOR,17,2) == Str(nSaldo,17,2) .And. SE1->E1_ACRESC > 0 .And. SE1->E1_SDACRES == 0
					nSAldo += SE1->E1_ACRESC
				Endif

				//Se abatimento verifico a data da baixa.
				//Por nao possuirem movimento de baixa no SE5, a saldotit retorna 
				//sempre saldo em aberto quando mv_par33 = 1 (Abatimentos = Lista)
				If SE1->E1_TIPO $ MVABATIM .and. ;
					((SE1->E1_BAIXA <= dDataBase .and. !Empty(SE1->E1_BAIXA)) .or. ;
					 (SE1->E1_MOVIMEN <= dDataBase .and. !Empty(SE1->E1_MOVIMEN))	)			
					nSaldo := 0
				Endif

			Else
				nSaldo := xMoeda((SE1->E1_SALDO+SE1->E1_SDACRES-SE1->E1_SDDECRE),SE1->E1_MOEDA,mv_par15,dDataReaj,ndecs+1,Iif(MV_PAR39==2,Iif(!Empty(SE1->E1_TXMOEDA),SE1->E1_TXMOEDA,RecMoeda(SE1->E1_EMISSAO,SE1->E1_MOEDA)),0))
			Endif

			//Caso exista desconto financeiro (cadastrado na inclusao do titulo), 
			//subtrai do valor principal.
			nDescont := FaDescFin("SE1",dBaixa,nSaldo,1)   
			If nDescont > 0
				nSaldo := nSaldo - nDescont
			Endif
			
			If ! SE1->E1_TIPO $ MVABATIM
				If ! (SE1->E1_TIPO $ MVRECANT+"/"+MV_CRNEG) .And. ;
						!( MV_PAR20 == 2 .And. nSaldo == 0 )  	// deve olhar abatimento pois e zerado o saldo na liquidacao final do titulo

					//Quando considerar Titulos com emissao futura, eh necessario
					//colocar-se a database para o futuro de forma que a Somaabat()
					//considere os titulos de abatimento
					If mv_par38 == 1
						dOldData := dDataBase
						dDataBase := CTOD("31/12/40")
					Endif

					nAbatim := SomaAbat(SE1->E1_PREFIXO,SE1->E1_NUM,SE1->E1_PARCELA,"R",mv_par15,dDataReaj,SE1->E1_CLIENTE,SE1->E1_LOJA)

					If mv_par38 == 1
						dDataBase := dOldData
					Endif

					If STR(nSaldo,17,2) == STR(nAbatim,17,2)
						nSaldo := 0
					ElseIf mv_par33 != 3
						nSaldo-= nAbatim
					Endif
				EndIf
			Endif	
			nSaldo:=Round(NoRound(nSaldo,3),2)
			
			//������������������������������������������������������Ŀ
			//� Desconsidera caso saldo seja menor ou igual a zero   �
			//��������������������������������������������������������
			If nSaldo <= 0
				dbSkip()
				Loop
			Endif
			
			dbSelectArea("SA1")
			MSSeek(cFilial+SE1->E1_CLIENTE+SE1->E1_LOJA)
			dbSelectArea("SA6")
			MSSeek(cFilial+SE1->E1_PORTADO)
			dbSelectArea("SE1")
			
			IF li > 58
				nAtuSM0 := SM0->(Recno())
				SM0->(dbGoto(nRegSM0))
				//cabec(titulo,cabec1,cabec2,nomeprog,tamanho,GetMv("MV_COMP"))
				SM0->(dbGoTo(nAtuSM0))
			EndIF
			
			If mv_par19 == 1
			    /*
				@li,	0 PSAY SE1->E1_CLIENTE + "-" + SE1->E1_LOJA + "-" +;
					IIF(mv_par29 == 1, SubStr(SA1->A1_NREDUZ,1,17), SubStr(SA1->A1_NOME,1,17))
				li := IIf (aTamCli[1] > 6,li+1,li)
				@li, 28 PSAY SE1->E1_PREFIXO+"-"+SE1->E1_NUM+"-"+SE1->E1_PARCELA
				If SE1->E1_TIPO$MVABATIM
					@li,46 PSAY "["
				Endif
				@li, 47 PSAY SE1->E1_TIPO
				If SE1->E1_TIPO$MVABATIM
					@li,50 PSAY "]"
				Endif
				@li, 51 PSAY SE1->E1_NATUREZ
				@li, 62 PSAY SE1->E1_EMISSAO
				@li, 73 PSAY SE1->E1_VENCTO
				@li, 84 PSAY SE1->E1_VENCREA
				@li, 95 PSAY SE1->E1_PORTADO+" "+SE1->E1_SITUACA
				*/
				nVlrOrig := Round(NoRound(xMoeda(SE1->E1_VALOR,SE1->E1_MOEDA,mv_par15,SE1->E1_EMISSAO,ndecs+1,Iif(MV_PAR39==2,Iif(!Empty(SE1->E1_TXMOEDA),SE1->E1_TXMOEDA,RecMoeda(SE1->E1_EMISSAO,SE1->E1_MOEDA)),0))* If(SE1->E1_TIPO$MVABATIM+"/"+MV_CRNEG+"/"+MVRECANT, -1,1),nDecs+1),nDecs)
				//@li,101 PSAY nVlrOrig Picture TM(nVlrOrig,15,nDecs)
			Endif
			
			If dDataBase > E1_VENCREA	//vencidos
				/*
				If mv_par19 == 1
					If ! SE1->E1_TIPO $ MVABATIM
						@li, 117 PSAY nSaldo * If(SE1->E1_TIPO$MV_CRNEG+"/"+MVRECANT, -1,1)  Picture TM(nSaldo,15,nDecs)
					EndIf
				EndIf
				*/
				nJuros:=0
				fa070Juros(mv_par15)
				dbSelectArea("SE1")
				
				/*
				If mv_par19 == 1
					@li,133 PSAY (nSaldo+nJuros)* If(SE1->E1_TIPO$MV_CRNEG+"/"+MVRECANT, -1,1) Picture TM(nSaldo+nJuros,15,nDecs)
				EndIf
				*/
				
				If SE1->E1_TIPO $ MVRECANT+"/"+MV_CRNEG
					nTit0 -= Round(NoRound(xMoeda(SE1->E1_VALOR,SE1->E1_MOEDA,mv_par15,SE1->E1_EMISSAO,ndecs+1,Iif(MV_PAR39==2,Iif(!Empty(SE1->E1_TXMOEDA),SE1->E1_TXMOEDA,RecMoeda(SE1->E1_EMISSAO,SE1->E1_MOEDA)),0)),ndecs+1),ndecs)
					nTit1 -= (nSaldo)
					nTit2 -= (nSaldo+nJuros)
					nMesTit0 -= Round(NoRound( xMoeda(SE1->E1_VALOR,SE1->E1_MOEDA,mv_par15,SE1->E1_EMISSAO,ndecs+1,Iif(MV_PAR39==2,Iif(!Empty(SE1->E1_TXMOEDA),SE1->E1_TXMOEDA,RecMoeda(SE1->E1_EMISSAO,SE1->E1_MOEDA)),0)),ndecs+1),ndecs)
					nMesTit1 -= (nSaldo)
					nMesTit2 -= (nSaldo+nJuros)
					nTotJur  -= nJuros
					nMesTitj -= nJuros
					nTotFilJ -= nJuros
				Else
					If !SE1->E1_TIPO $ MVABATIM
						nTit0 += Round(NoRound(xMoeda(SE1->E1_VALOR,SE1->E1_MOEDA,mv_par15,SE1->E1_EMISSAO,ndecs+1,Iif(MV_PAR39==2,Iif(!Empty(SE1->E1_TXMOEDA),SE1->E1_TXMOEDA,RecMoeda(SE1->E1_EMISSAO,SE1->E1_MOEDA)),0)),ndecs+1),ndecs)						
						nTit1 += (nSaldo)
						nTit2 += (nSaldo+nJuros)
						nMesTit0 += Round(NoRound(xMoeda(SE1->E1_VALOR,SE1->E1_MOEDA,mv_par15,SE1->E1_EMISSAO,ndecs+1,Iif(MV_PAR39==2,Iif(!Empty(SE1->E1_TXMOEDA),SE1->E1_TXMOEDA,RecMoeda(SE1->E1_EMISSAO,SE1->E1_MOEDA)),0)),ndecs+1),ndecs)
						nMesTit1 += (nSaldo)
						nMesTit2 += (nSaldo+nJuros)
						nTotJur  += nJuros
						nMesTitj += nJuros
						nTotFilJ += nJuros
					Endif	
				Endif
			Else						//a vencer
				/*
				If mv_par19 == 1
					@li,149 PSAY nSaldo * If(SE1->E1_TIPO$MV_CRNEG+"/"+MVRECANT, -1,1)  Picture TM(nSaldo,15,nDecs)
				EndIf
			    */
				If ! ( SE1->E1_TIPO $ MVRECANT+"/"+MV_CRNEG)
					If !SE1->E1_TIPO $ MVABATIM
						nTit0 += Round(NoRound(xMoeda(SE1->E1_VALOR,SE1->E1_MOEDA,mv_par15,SE1->E1_EMISSAO,ndecs+1,Iif(MV_PAR39==2,Iif(!Empty(SE1->E1_TXMOEDA),SE1->E1_TXMOEDA,RecMoeda(SE1->E1_EMISSAO,SE1->E1_MOEDA)),0)),ndecs+1),ndecs)
						nTit3 += (nSaldo-nTotAbat)
						nTit4 += (nSaldo-nTotAbat)
						nMesTit0 += Round(NoRound(xMoeda(SE1->E1_VALOR,SE1->E1_MOEDA,mv_par15,SE1->E1_EMISSAO,ndecs+1,Iif(MV_PAR39==2,Iif(!Empty(SE1->E1_TXMOEDA),SE1->E1_TXMOEDA,RecMoeda(SE1->E1_EMISSAO,SE1->E1_MOEDA)),0)),ndecs+1),ndecs)
						nMesTit3 += (nSaldo-nTotAbat)
						nMesTit4 += (nSaldo-nTotAbat)
					EndIf
				Else
					nTit0 -= Round(NoRound(xMoeda(SE1->E1_VALOR,SE1->E1_MOEDA,mv_par15,SE1->E1_EMISSAO,ndecs+1,Iif(MV_PAR39==2,Iif(!Empty(SE1->E1_TXMOEDA),SE1->E1_TXMOEDA,RecMoeda(SE1->E1_EMISSAO,SE1->E1_MOEDA)),0)),ndecs+1),ndecs)
					nTit3 -= (nSaldo-nTotAbat)
					nTit4 -= (nSaldo-nTotAbat)
					nMesTit0 -= Round(NoRound(xMoeda(SE1->E1_VALOR,SE1->E1_MOEDA,mv_par15,SE1->E1_EMISSAO,ndecs+1,Iif(MV_PAR39==2,Iif(!Empty(SE1->E1_TXMOEDA),SE1->E1_TXMOEDA,RecMoeda(SE1->E1_EMISSAO,SE1->E1_MOEDA)),0)),ndecs+1),ndecs)
					nMesTit3 -= (nSaldo-nTotAbat)
					nMesTit4 -= (nSaldo-nTotAbat)
				Endif
			Endif
			/*
			If mv_par19 == 1
				@ li, 165 PSAY SE1->E1_NUMBCO
			EndIf
			*/
			If nJuros > 0
				/*
				If mv_par19 == 1
					@ Li,182 PSAY nJuros Picture PesqPict("SE1","E1_JUROS",10,MV_PAR15)
				EndIf
				*/
				nJuros := 0
			Endif
			
			IF dDataBase > SE1->E1_VENCREA
				nAtraso:=dDataBase-SE1->E1_VENCTO
				IF Dow(SE1->E1_VENCTO) == 1 .Or. Dow(SE1->E1_VENCTO) == 7
					IF Dow(dBaixa) == 2 .and. nAtraso <= 2
						nAtraso := 0
					EndIF
				EndIF
				nAtraso:=IIF(nAtraso<0,0,nAtraso)
				IF nAtraso>0
					/*
					If mv_par19 == 1
						@li ,195 PSAY nAtraso Picture "9999"
					EndIf
					*/
				EndIF
			EndIF
			If mv_par19 == 1
				//@li,200 PSAY SubStr(SE1->E1_HIST,1,20)+ ;
					IIF(E1_TIPO $ MVPROVIS,"*"," ")+ ;
					Iif(nSaldo == Round(NoRound(xMoeda(SE1->E1_VALOR,SE1->E1_MOEDA,mv_par15,SE1->E1_EMISSAO,ndecs+1,Iif(MV_PAR39==2,Iif(!Empty(SE1->E1_TXMOEDA),SE1->E1_TXMOEDA,RecMoeda(SE1->E1_EMISSAO,SE1->E1_MOEDA)),0)),ndecs+1),ndecs)," ","P")
			EndIf
			
			//����������������������������������������Ŀ
			//� Carrega data do registro para permitir �
			//� posterior an�lise de quebra por mes.   �
			//������������������������������������������
			dDataAnt := Iif(nOrdem == 6, SE1->E1_EMISSAO, SE1->E1_VENCREA)
			dbSkip()
			nTotTit ++
			nMesTTit ++
			nTotFiltit++
			nTit5 ++
			If mv_par19 == 1
				li++
			EndIf
		Enddo
		
		IF nTit5 > 0 .and. nOrdem != 2 .and. nOrdem != 10
			SubTot130(nTit0,nTit1,nTit2,nTit3,nTit4,nOrdem,cCarAnt,nTotJur,nDecs)
			If mv_par19 == 1
				Li++
			EndIf
		Endif
		
		//����������������������������������������Ŀ
		//� Verifica quebra por m�s	  			   �
		//������������������������������������������
		lQuebra := .F.
		If nOrdem == 4  .and. (Month(SE1->E1_VENCREA) # Month(dDataAnt) .or. SE1->E1_VENCREA > mv_par10)
			lQuebra := .T.
		Elseif nOrdem == 6 .and. (Month(SE1->E1_EMISSAO) # Month(dDataAnt) .or. SE1->E1_EMISSAO > mv_par14)
			lQuebra := .T.
		Endif
		If lQuebra .and. nMesTTit # 0
			//ImpMes130(nMesTit0,nMesTit1,nMesTit2,nMesTit3,nMesTit4,nMesTTit,nMesTitJ,nDecs)
			nMesTit0 := nMesTit1 := nMesTit2 := nMesTit3 := nMesTit4 := nMesTTit := nMesTitj := 0
		Endif
		nTot0+=nTit0
		nTot1+=nTit1
		nTot2+=nTit2
		nTot3+=nTit3
		nTot4+=nTit4
		nTotJ+=nTotJur
		
		nTotFil0+=nTit0
		nTotFil1+=nTit1
		nTotFil2+=nTit2
		nTotFil3+=nTit3
		nTotFil4+=nTit4
		Store 0 To nTit0,nTit1,nTit2,nTit3,nTit4,nTit5,nTotJur,nTotAbat
	Enddo
	
	//Alert("Verifica��o das Variaveis  neste ponto - nTot0 "+Alltrim(Str(nTot0))+" nTot1 "+Alltrim(Str(nTot1))+" nTot2 "+Alltrim(Str(nTot2))+" nTot3 "+Alltrim(Str(nTot3))+" nTot4 "+Alltrim(Str(nTot4))+" nTotJ "+Alltrim(Str(nTotJ)))
	
//	Alert("Verifica��o das Variaveis  neste ponto - nTotFil0 "+Alltrim(Str(nTotFil0))+" nTotFil1 "+Alltrim(Str(nTotFil1))+" nTotFil2 "+Alltrim(Str(nTot2))+" nTotFil2 "+Alltrim(Str(nTot3))+" nTotFil3 "+Alltrim(Str(nTotFil3))+" nTotFil4 "+Alltrim(Str(nTotFil4)))
	nVoltaVlr := IIF(nFunCham == 1,nTotFil4,nTotFil1)
	
	dbSelectArea("SE1")		// voltar para alias existente, se nao, nao funciona
	
	//����������������������������������������Ŀ
	//� Imprimir TOTAL por filial somente quan-�
	//� do houver mais do que 1 filial.        �
	//������������������������������������������
	if mv_par21 == 1 .and. SM0->(Reccount()) > 1
		//ImpFil130(nTotFil0,nTotFil1,nTotFil2,nTotFil3,nTotFil4,nTotFiltit,nTotFilJ,nDecs)
	Endif
	Store 0 To nTotFil0,nTotFil1,nTotFil2,nTotFil3,nTotFil4,nTotFilTit,nTotFilJ
	If Empty(xFilial("SE1"))
		Exit
	Endif
	
	#IFDEF TOP
		dbSelectArea("SE1")
		dbCloseArea()
		ChKFile("SE1")
		dbSelectArea("SE1")
		dbSetOrder(1)
	#ENDIF
	
	dbSelectArea("SM0")
	dbSkip()

Enddo

SM0->(dbGoTo(nRegSM0))
cFilAnt := SM0->M0_CODFIL
IF li != 80
    /*
	IF li > 58
		cabec(titulo,cabec1,cabec2,nomeprog,tamanho,GetMv("MV_COMP"))
	EndIF
	*/
	//TotGer130(nTot0,nTot1,nTot2,nTot3,nTot4,nTotTit,nTotJ,nDecs)
	//Roda(cbcont,cbtxt,"G")
EndIF

//Set Device To Screen

#IFNDEF TOP
	dbSelectArea("SE1")
	dbClearFil()
	RetIndex( "SE1" )
	If !Empty(cIndexSE1)
		FErase (cIndexSE1+OrdBagExt())
	Endif
	dbSetOrder(1)
#ELSE
	dbSelectArea("SE1")
	dbCloseArea()
	ChKFile("SE1")
	dbSelectArea("SE1")
	dbSetOrder(1)
#ENDIF
/*
If aReturn[5] = 1
	Set Printer TO
	dbCommitAll()
	Ourspool(wnrel)
Endif
MS_FLUSH()
*/

Return(nVoltaVlr)

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 �SubTot130 � Autor � Paulo Boschetti 		  � Data � 01.06.92 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o �Imprimir SubTotal do Relatorio										  ���
�������������������������������������������������������������������������Ĵ��
���Sintaxe e � SubTot130()																  ���
�������������������������������������������������������������������������Ĵ��
���Parametros�																				  ���
�������������������������������������������������������������������������Ĵ��
��� Uso 	    � Generico																	  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/
Static Function SubTot130(nTit0,nTit1,nTit2,nTit3,nTit4,nOrdem,cCarAnt,nTotJur,nDecs)

Local cCarteira := " "
Local cTelefone := Alltrim(Transform(SA1->A1_DDD, PesqPict("SA1","A1_DDD"))+"-"+ Transform(SA1->A1_TEL, PesqPict("SA1","A1_TEL")) )

DEFAULT nDecs := Msdecimais(mv_par15)

If mv_par19 == 1
	li++
EndIf
IF li > 58
	nAtuSM0 := SM0->(Recno())
	SM0->(dbGoto(nRegSM0))
	//cabec(titulo,cabec1,cabec2,nomeprog,tamanho,GetMv("MV_COMP"))
	SM0->(dbGoTo(nAtuSM0))
EndIF
If nOrdem = 1
	//@li,000 PSAY IIF(mv_par29 == 1,Substr(SA1->A1_NREDUZ,1,40),Substr(SA1->A1_NOME,1,40))+" "+ cTelefone + " "+ STR0054 + Right(cCarAnt,2)+Iif(mv_par21==1,STR0055+cFilAnt + " - " + Alltrim(SM0->M0_FILIAL),"") //"Loja - "###" Filial - "
Elseif nOrdem == 4 .or. nOrdem == 6
	//@li,000 PSAY OemToAnsi(STR0037)  // "S U B - T O T A L ----> "
	//@li,028 PSAY cCarAnt
	//@li,PCOL()+2 PSAY Iif(mv_par21==1,cFilAnt+ " - " + Alltrim(SM0->M0_FILIAL),"  ")
Elseif nOrdem = 3
	//@li,000 PSAY OemToAnsi(STR0037)  // "S U B - T O T A L ----> "
	//@li,028 PSAY Iif(Empty(SA6->A6_NREDUZ),OemToAnsi(STR0029),SA6->A6_NREDUZ) + " " + Iif(mv_par21==1,cFilAnt+ " - " + Alltrim(SM0->M0_FILIAL),"")
ElseIf nOrdem == 5
	dbSelectArea("SED")
	dbSeek(cFilial+cCarAnt)
	//@li,000 PSAY OemToAnsi(STR0037)  // "S U B - T O T A L ----> "
	//@li,028 PSAY cCarAnt + " "+Substr(ED_DESCRIC,1,50) + " " + Iif(mv_par21==1,cFilAnt+ " - " + Alltrim(SM0->M0_FILIAL),"")
	dbSelectArea("SE1")
Elseif nOrdem == 7
	//@li,000 PSAY OemToAnsi(STR0037)  // "S U B - T O T A L ----> "
	//@li,028 PSAY SubStr(cCarAnt,7,2)+"/"+SubStr(cCarAnt,5,2)+"/"+SubStr(cCarAnt,3,2)+" - "+SubStr(cCarAnt,9,3) + " " +Iif(mv_par21==1,cFilAnt+ " - " + Alltrim(SM0->M0_FILIAL),"")
ElseIf nOrdem = 8
	//@li,000 PSAY SA1->A1_COD+" "+Substr(SA1->A1_NOME,1,40)+" "+ cTelefone + " " + Iif(mv_par21==1,cFilAnt+ " - " + Alltrim(SM0->M0_FILIAL),"")
ElseIf nOrdem = 9
	cCarteira := Situcob(cCarAnt)
	//@li,000 PSAY SA6->A6_COD+" "+SA6->A6_NREDUZ + SubStr(cCarteira,1,2) + " "+SubStr(cCarteira,3,20) + " " + Iif(mv_par21==1,cFilAnt+ " - " + Alltrim(SM0->M0_FILIAL),"")
Endif
/*
If mv_par19 == 1
	@li,101 PSAY nTit0		  Picture TM(nTit0,15,nDecs)
Endif
@li,117 PSAY nTit1		  Picture TM(nTit1,15,nDecs)
@li,133 PSAY nTit2		  Picture TM(nTit2,15,nDecs)
@li,149 PSAY nTit3		  Picture TM(nTit3,15,nDecs)
If nTotJur > 0
	@li,180 PSAY nTotJur  Picture TM(nTotJur,12,nDecs)
Endif
@li,204 PSAY nTit2+nTit3 Picture TM(nTit2+nTit3,15,nDecs)
li++
If (nOrdem = 1 .Or. nOrdem == 8) .And. mv_par35 == 1 // Salta pag. por cliente
	cabec(titulo,cabec1,cabec2,nomeprog,tamanho,GetMv("MV_COMP"))
Endif
*/
Return .T.

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 � TotGer130� Autor � Paulo Boschetti       � Data � 01.06.92 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Imprimir total do relatorio										  ���
�������������������������������������������������������������������������Ĵ��
���Sintaxe e � TotGer130()																  ���
�������������������������������������������������������������������������Ĵ��
���Parametros�																				  ���
�������������������������������������������������������������������������Ĵ��
��� Uso      � Generico																	  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/
STATIC Function TotGer130(nTot0,nTot1,nTot2,nTot3,nTot4,nTotTit,nTotJ,nDecs)

DEFAULT nDecs := Msdecimais(mv_par15)

li++
IF li > 58
	nAtuSM0 := SM0->(Recno())
	SM0->(dbGoto(nRegSM0))
	//cabec(titulo,cabec1,cabec2,nomeprog,tamanho,GetMv("MV_COMP"))
	SM0->(dbGoTo(nAtuSM0))
EndIF

//@li,000 PSAY OemToAnsi(STR0038) //"T O T A L   G E R A L ----> " + " " + Iif(mv_par21==1,cFilAnt,"")
//@li,028 PSAY "("+ALLTRIM(STR(nTotTit))+" "+IIF(nTotTit > 1,OemToAnsi(STR0039),OemToAnsi(STR0040))+")"		//"TITULOS"###"TITULO"
/*
If mv_par19 == 1
	@li,101 PSAY nTot0		  Picture TM(nTot0,15,nDecs)
Endif
@li,117 PSAY nTot1		  Picture TM(nTot1,15,nDecs)
@li,133 PSAY nTot2		  Picture TM(nTot2,15,nDecs)
@li,149 PSAY nTot3		  Picture TM(nTot3,15,nDecs)
@li,177 PSAY nTotJ		  Picture TM(nTotJ,15,nDecs)
@li,204 PSAY nTot2+nTot3 Picture TM(nTot2+nTot3,15,nDecs)
li++
li++
*/
Return .T.

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 �ImpMes130 � Autor � Vinicius Barreira	  � Data � 12.12.94 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o �IMPRIMIR TOTAL DO RELATORIO - QUEBRA POR MES					  ���
�������������������������������������������������������������������������Ĵ��
���Sintaxe e � ImpMes130() 															  ���
�������������������������������������������������������������������������Ĵ��
���Parametros� 																			  ���
�������������������������������������������������������������������������Ĵ��
��� Uso		 � Generico 																  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/
STATIC Function ImpMes130(nMesTot0,nMesTot1,nMesTot2,nMesTot3,nMesTot4,nMesTTit,nMesTotJ,nDecs)

DEFAULT nDecs := Msdecimais(mv_par15)
li++
IF li > 58
	nAtuSM0 := SM0->(Recno())
	SM0->(dbGoto(nRegSM0))
	//cabec(titulo,cabec1,cabec2,nomeprog,tamanho,GetMv("MV_COMP"))
	SM0->(dbGoTo(nAtuSM0))
EndIF
//@li,000 PSAY OemToAnsi(STR0041)  //"T O T A L   D O  M E S ---> "
//@li,028 PSAY "("+ALLTRIM(STR(nMesTTit))+" "+IIF(nMesTTit > 1,OemToAnsi(STR0039),OemToAnsi(STR0040))+")"  //"TITULOS"###"TITULO"
/*
If mv_par19 == 1
	@li,101 PSAY nMesTot0   Picture TM(nMesTot0,14,nDecs)
Endif
@li,116 PSAY nMesTot1	Picture TM(nMesTot1,14,nDecs)
@li,133 PSAY nMesTot2	Picture TM(nMesTot2,14,nDecs)
@li,149 PSAY nMesTot3	Picture TM(nMesTot3,14,nDecs)
@li,179 PSAY nMesTotJ	Picture TM(nMesTotJ,12,nDecs)
@li,204 PSAY nMesTot2+nMesTot3 Picture TM(nMesTot2+nMesTot3,16,nDecs)
li+=2
*/
Return(.T.)

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 � ImpFil130� Autor � Paulo Boschetti  	  � Data � 01.06.92 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Imprimir total do relatorio										  ���
�������������������������������������������������������������������������Ĵ��
���Sintaxe e � ImpFil130()																  ���
�������������������������������������������������������������������������Ĵ��
���Parametros�																				  ���
�������������������������������������������������������������������������Ĵ��
��� Uso 	    � Generico																	  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/
STATIC Function ImpFil130(nTotFil0,nTotFil1,nTotFil2,nTotFil3,nTotFil4,nTotFilTit,nTotFilJ,nDecs)

DEFAULT nDecs := Msdecimais(mv_par15)

li++
IF li > 58
	nAtuSM0 := SM0->(Recno())
	SM0->(dbGoto(nRegSM0))
	//cabec(titulo,cabec1,cabec2,nomeprog,tamanho,GetMv("MV_COMP"))
	SM0->(dbGoTo(nAtuSM0))
EndIF
/*
@li,000 PSAY OemToAnsi(STR0043)+" "+Iif(mv_par21==1,cFilAnt+" - " + AllTrim(SM0->M0_FILIAL),"")  //"T O T A L   F I L I A L ----> "
If mv_par19 == 1
	@li,101 PSAY nTotFil0        Picture TM(nTotFil0,14,nDecs)
Endif
@li,116 PSAY nTotFil1        Picture TM(nTotFil1,14,nDecs)
@li,133 PSAY nTotFil2        Picture TM(nTotFil2,14,nDecs)
@li,149 PSAY nTotFil3        Picture TM(nTotFil3,14,nDecs)
@li,179 PSAY nTotFilJ		  Picture TM(nTotFilJ,12,nDecs)
@li,204 PSAY nTotFil2+nTotFil3 Picture TM(nTotFil2+nTotFil3,16,nDecs)
li+=2
*/
Return .T.

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 �fr130Indr � Autor � Wagner           	  � Data � 12.12.94 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o �Monta Indregua para impressao do relat�rio						  ���
�������������������������������������������������������������������������Ĵ��
��� Uso		 � Generico 																  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/
Static Function fr130IndR()
Local cString

cString := 'E1_FILIAL=="'+xFilial("SE1")+'".And.'
cString += 'E1_CLIENTE>="'+mv_par01+'".and.E1_CLIENTE<="'+mv_par02+'".And.'
cString += 'E1_PREFIXO>="'+mv_par03+'".and.E1_PREFIXO<="'+mv_par04+'".And.'
cString += 'E1_NUM>="'+mv_par05+'".and.E1_NUM<="'+mv_par06+'".And.'
cString += 'DTOS(E1_VENCREA)>="'+DTOS(mv_par09)+'".and.DTOS(E1_VENCREA)<="'+DTOS(mv_par10)+'".And.'
cString += '(E1_MULTNAT == "1" .OR. (E1_NATUREZ>="'+mv_par11+'".and.E1_NATUREZ<="'+mv_par12+'")).And.'
cString += 'DTOS(E1_EMISSAO)>="'+DTOS(mv_par13)+'".and.DTOS(E1_EMISSAO)<="'+DTOS(mv_par14)+'"'
If !Empty(mv_par31) // Deseja imprimir apenas os tipos do parametro 31
	cString += '.And.E1_TIPO$"'+mv_par31+'"'
ElseIf !Empty(Mv_par32) // Deseja excluir os tipos do parametro 32
	cString += '.And.!(E1_TIPO$'+'"'+mv_par32+'")'
EndIf
IF mv_par34 == 1  // Apenas titulos que estarao no fluxo de caixa
	cString += '.And.(E1_FLUXO!="N")'	
Endif
Return cString

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � PutDtBase� Autor � Mauricio Pequim Jr    � Data � 18/07/02 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Acerta parametro database do relatorio                     ���
�������������������������������������������������������������������������Ĵ��
���Uso       � Finr130.prx                                                ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/
Static Function PutDtBase()
Local _sAlias	:= Alias()

dbSelectArea("SX1")
dbSetOrder(1)
If MsSeek("FIN13036")
	//Acerto o parametro com a database
	RecLock("SX1",.F.)
	Replace x1_cnt01		With "'"+DTOC(dDataBase)+"'"
	MsUnlock()	
Endif

dbSelectArea(_sAlias)
Return

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �AjustaSX1 �Autor  �Mauricio Pequim Jr. � Data �29.12.2004   ���
�������������������������������������������������������������������������͹��
���Desc.     �Insere novas perguntas ao sx1                               ���
�������������������������������������������������������������������������͹��
���Uso       � FINR130                                                    ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function AjustaSX1()
/*
Local aHelpPor	:= {}
Local aHelpEng	:= {}
Local aHelpSpa	:= {}
					  
Aadd( aHelpPor, 'Selecione "Sim" para que sejam         '  )
Aadd( aHelpPor, 'considerados no relat�rio, t�tulos cuja'  )
Aadd( aHelpPor, 'emiss�o seja em data posterior a database' )
Aadd( aHelpPor, 'do relat�rio, ou "N�o" em caso contr�rio'  )

Aadd( aHelpSpa, 'Seleccione "S�" para que se consideren en'	)
Aadd( aHelpSpa, 'el informe los t�tulos cuya emisi�n sea en')
Aadd( aHelpSpa, 'fecha posterior a la fecha base de dicho '	)
Aadd( aHelpSpa, 'informe o "No" en caso contrario'	)

Aadd( aHelpEng, 'Choose "Yes" for bills which issue date '	)
Aadd( aHelpEng, 'is posterior to the report base date in ' 	)
Aadd( aHelpEng, 'the report, otherwise "No"' 			)

PutSx1( "FIN130", "38","Tit. Emissao Futura?","Tit.Emision Futura  ","Future Issue Bill   ","mv_chs","N",1,0,2,"C","","","","","mv_par38","Sim","Si","Yes","","Nao","No","No","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

aHelpPor	:= {}
aHelpEng	:= {}
aHelpSpa	:= {}

Aadd( aHelpPor, "Selecione qual taxa sera utilizada para "  )
Aadd( aHelpPor, "conversao dos valores"   )

Aadd( aHelpSpa, "Seleccione la tasa de conversi�n " )
Aadd( aHelpSpa, "que se utilizara " )

Aadd( aHelpEng, "Select conversion rate to be used"   )



PutSx1( "FIN130", "39","Converte valores pela?","�Convierte valores por?","Convert values by?","mv_cht","N",1,0,1,"C","","","","",;
	"mv_par39","Taxa do dia","Tasa del dia","Rate of the day","","Taxa do Mov.","Tasa del mov.","Movement rate","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

*/
Return


/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �Situcob   �Autor  �Mauricio Pequim Jr. � Data �13.04.2005   ���
�������������������������������������������������������������������������͹��
���Desc.     �Retorna situacao de cobranca do titulo                      ���
�������������������������������������������������������������������������͹��
���Uso       � FINR130                                                    ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function SituCob(cCarAnt)

Local aSituaca := {}
Local aArea		:= GetArea()
Local cCart		:= " "

//����������������������������������������������������������������������Ŀ
//� Monta a tabela de situa��es de T�tulos										 �
//������������������������������������������������������������������������
dbSelectArea("SX5")
dbSeek(cFilial+"07")
While SX5->X5_FILIAL+SX5->X5_tabela == cFilial+"07"
	cCapital := Capital(X5Descri())
	AADD( aSituaca,{SubStr(SX5->X5_CHAVE,1,2),OemToAnsi(SubStr(cCapital,1,20))})
	dbSkip()
EndDo

nOpcS := (Ascan(aSituaca,{|x| Alltrim(x[1])== Substr(cCarAnt,4,1) }))
If nOpcS > 0
	cCart := aSituaca[nOpcS,1]+aSituaca[nOpcs,2]		
ElseIf Empty(SE1->E1_SITUACA)
	cCart := "0 "+STR0029
Endif
RestArea(aArea)
Return cCart

/************************************ F U N � � O    D O   S A L D O   A   P A G A R ****************************/

Static Function SaldoSP(dPADtParam,dPA2DtParam,dPA3DtParam,nPAChamada)

Private dPARGDBase:= dPADtParam
Private dPADoVenc := dPA2DtParam
Private dPAAteVenc:= dPA3DtParam
Private nFunPACham:= nPAChamada

//Alert("Estou em SaldoSP paramentos recebidos, DataBase "+DTOC(dPARGDBase)+" Vencto De "+dtoc(dPADoVenc)+" Vencto Ate "+DTOC(dPAAteVenc))
Private nRetPAValor := nSldAP()
//Alert("Valor a sert retornado da SaldoSP, nRetPAValor "+Alltrim(Str(nRetPAValor)))
Return(nRetPAValor) 



/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 � FINR150R3� Autor � Wagner Xavier   	     � Data � 02.10.92 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Posi��o dos Titulos a Pagar										  ���
�������������������������������������������������������������������������Ĵ��
���Sintaxe e � FINR150R3(void)														  ���
�������������������������������������������������������������������������Ĵ��
���Par�metros� 																			  ���
�������������������������������������������������������������������������Ĵ��
��� Uso		 � Gen�rico 																  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function nSldAP()
//������������������Ŀ
//� Define Vari�veis �
//��������������������
Local cDesc1 :=OemToAnsi(STR0001)  //"Imprime a posi��o dos titulos a pagar relativo a data base"
Local cDesc2 :=OemToAnsi(STR0002)  //"do sistema."
LOCAL cDesc3 :=""
LOCAL wnrel
LOCAL cString:="SE2"
LOCAL nRegEmp := SM0->(RecNo())
Local dOldDtBase := dDataBase
Local dOldData := dDataBase

PRIVATE aReturn := { OemToAnsi(STR0003), 1,OemToAnsi(STR0004), 1, 2, 1, "",1 }  //"Zebrado"###"Administracao"
PRIVATE nomeprog:="FINR150"
PRIVATE aLinha  := { },nLastKey := 0
PRIVATE cPerg	 :="FIN150"
PRIVATE nJuros  :=0
PRIVATE tamanho:="G"

PRIVATE titulo  := ""
PRIVATE cabec1
PRIVATE cabec2

//��������������������������Ŀ
//� Definicao dos cabe�alhos �
//����������������������������
titulo := OemToAnsi(STR0005)  //"Posicao dos Titulos a Pagar"
cabec1 := OemToAnsi(STR0006)  //"Codigo Nome do Fornecedor   PRF-Numero         Tp  Natureza    Data de  Vencto   Vencto  Banco  Valor Original |        Titulos vencidos          | Titulos a vencer | Porta-|  Vlr.juros ou   Dias   Historico     "
cabec2 := OemToAnsi(STR0007)  //"                            Parcela                            Emissao  Titulo    Real                         |  Valor nominal   Valor corrigido |   Valor nominal  | dor   |   permanencia   Atraso               "

//Nao retire esta chamada. Verifique antes !!!
//Ela � necessaria para o correto funcionamento da pergunte 36 (Data Base)
//PutDtBase()

//������������������������������������Ŀ
//� Verifica as perguntas selecionadas �
//��������������������������������������
//AjustaSx1()
//pergunte("FIN150",.F.)
//��������������������������������������Ŀ
//� Variaveis utilizadas para parametros �
//� mv_par01	  // do Numero 			  �
//� mv_par02	  // at� o Numero 		  �
//� mv_par03	  // do Prefixo			  �
//� mv_par04	  // at� o Prefixo		  �
//� mv_par05	  // da Natureza  	     �
//� mv_par06	  // at� a Natureza		  �
//� mv_par07	  // do Vencimento		  �
//� mv_par08	  // at� o Vencimento	  �
//� mv_par09	  // do Banco			     �
//� mv_par10	  // at� o Banco		     �
//� mv_par11	  // do Fornecedor		  �
//� mv_par12	  // at� o Fornecedor	  �
//� mv_par13	  // Da Emiss�o			  �
//� mv_par14	  // Ate a Emiss�o		  �
//� mv_par15	  // qual Moeda			  �
//� mv_par16	  // Imprime Provis�rios  �
//� mv_par17	  // Reajuste pelo vencto �
//� mv_par18	  // Da data contabil	  �
//� mv_par19	  // Ate data contabil	  �
//� mv_par20	  // Imprime Rel anal/sint�
//� mv_par21	  // Considera  Data Base?�
//� mv_par22	  // Cons filiais abaixo ?�
//� mv_par23	  // Filial de            �
//� mv_par24	  // Filial ate           �
//� mv_par25	  // Loja de              �
//� mv_par26	  // Loja ate             �
//� mv_par27 	  // Considera Adiantam.? �
//� mv_par28	  // Imprime Nome 		  �
//� mv_par29	  // Outras Moedas 		  �
//� mv_par30     // Imprimir os Tipos    �
//� mv_par31     // Nao Imprimir Tipos	  �
//� mv_par32     // Consid. Fluxo Caixa  �
//� mv_par33     // DataBase             �
//� mv_par34     // Tipo de Data p/Saldo �
//� mv_par35     // Quanto a taxa		  �
//� mv_par36     // Tit.Emissao Futura	  �
//����������������������������������������
//���������������������������������������Ŀ
//� Envia controle para a funcao SETPRINT �
//�����������������������������������������

mv_par01 :=	"      "  // do Numero 			  �
mv_par02 :=	"zzzzzz"  // at� o Numero 		  �
mv_par03 :=	"   "  // do Prefixo			  �
mv_par04 :=	"zzz"  // at� o Prefixo		  �
mv_par05 :=	"          "  // da Natureza  	     �
mv_par06 :=	"zzzzzzzzzz"  // at� a Natureza		  �
mv_par07 :=	dPADoVenc   // do Vencimento		  �
mv_par08 :=	dPAAteVenc  // at� o Vencimento	  �
mv_par09 :=	"   "  // do Banco			     �
mv_par10 :=	"zzz"  // at� o Banco		     �
mv_par11 :=	"      "  // do Fornecedor		  �
mv_par12 :=	"zzzzzz"  // at� o Fornecedor	  �
mv_par13 :=	CTOD("01/01/90")  // Da Emiss�o			  �
mv_par14 :=	CTOD("31/12/49")  // Ate a Emiss�o		  �
mv_par15 :=	1  // qual Moeda			  �
mv_par16 :=	1  // Imprime Provis�rios  �
mv_par17 :=	1  // Reajuste pelo vencto �
mv_par18 :=	CTOD("01/01/90")  // Da data contabil	  �
mv_par19 :=	CTOD("31/12/49")  // Ate data contabil	  �
mv_par20 :=	2  // Imprime Rel anal/sint�
mv_par21 :=	1 // Considera  Data Base?�
mv_par22 :=	1  // Cons filiais abaixo ?�
mv_par23 :=	SM0->M0_CODFIL  // Filial de            �
mv_par24 :=	SM0->M0_CODFIL  // Filial ate           �
mv_par25 :=	"  "  // Loja de              �
mv_par26 :=	"zz"  // Loja ate             �
mv_par27 :=	2  // Considera Adiantam.? �
mv_par28 :=	2  // Imprime Nome 		  �
mv_par29 :=	2  // Outras Moedas 		  �
mv_par30 := "                                        "   // Imprimir os Tipos    �
mv_par31 := "NCC, AB-                                "   // Nao Imprimir Tipos	  �
mv_par32 := 2   // Consid. Fluxo Caixa  �
mv_par33 := dPARGDBase    // DataBase             �
mv_par34 := 2   // Tipo de Data p/Saldo �
mv_par35 := 1    // Quanto a taxa		  �
mv_par36 := 2    // Tit.Emissao Futura	  �

#IFDEF TOP
	IF TcSrvType() == "AS/400" .and. Select("__SE2") == 0
		ChkFile("SE2",.f.,"__SE2")
	Endif
#ENDIF

wnrel := "FINR150"            //Nome Default do relatorio em Disco
aOrd	:= {OemToAnsi(STR0008),OemToAnsi(STR0009),OemToAnsi(STR0010) ,;  //"Por Numero"###"Por Natureza"###"Por Vencimento"
	OemToAnsi(STR0011),OemToAnsi(STR0012),OemToAnsi(STR0013),OemToAnsi(STR0014) }  //"Por Banco"###"Fornecedor"###"Por Emissao"###"Por Cod.Fornec."
//wnrel := SetPrint(cString,wnrel,cPerg,@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,,Tamanho)
/*
If nLastKey == 27
	Return
Endif

SetDefault(aReturn,cString)

If nLastKey == 27
	Return
Endif

RptStatus({|lEnd| Fa150Imp(@lEnd,wnRel,cString)},Titulo)
*/
nValRet := CalcSAP()
//���������������������������������������Ŀ
//� Restaura empresa / filial original    �
//�����������������������������������������
SM0->(dbGoto(nRegEmp))
cFilAnt := SM0->M0_CODFIL

//Acerta a database de acordo com a database real do sistema
If mv_par21 == 1    // Considera Data Base
	dDataBase := dOldDtBase
Endif	

Return(nValRet)

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 � FA150Imp � Autor � Wagner Xavier 		  � Data � 02.10.92 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Posi��o dos Titulos a Pagar										  ���
�������������������������������������������������������������������������Ĵ��
���Sintaxe e � FA150Imp(lEnd,wnRel,cString)										  ���
�������������������������������������������������������������������������Ĵ��
���Parametros� lEnd	  - A��o do Codeblock										  ���
���			 � wnRel   - T�tulo do relat�rio 									  ���
���			 � cString - Mensagem										 			  ���
�������������������������������������������������������������������������Ĵ��
��� Uso		 � Gen�rico 																  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function CalcSAP() //FA150Imp(lEnd,wnRel,cString)

LOCAL CbCont
LOCAL CbTxt
LOCAL nOrdem :=0
LOCAL nQualIndice := 0
LOCAL lContinua := .T.
LOCAL nTit0:=0,nTit1:=0,nTit2:=0,nTit3:=0,nTit4:=0,nTit5:=0
LOCAL nTot0:=0,nTot1:=0,nTot2:=0,nTot3:=0,nTot4:=0,nTotTit:=0,nTotJ:=0,nTotJur:=0
LOCAL nFil0:=0,nFil1:=0,nFil2:=0,nFil3:=0,nFil4:=0,nFilTit:=0,nFilJ:=0
LOCAL cCond1,cCond2,cCarAnt,nSaldo:=0,nAtraso:=0
LOCAL dDataReaj
LOCAL dDataAnt := dDataBase , lQuebra
LOCAL nMestit0:= nMesTit1 := nMesTit2 := nMesTit3 := nMesTit4 := nMesTTit := nMesTitj := 0
LOCAL dDtContab
LOCAL	cIndexSe2
LOCAL	cChaveSe2
LOCAL	nIndexSE2
LOCAL cFilDe,cFilAte
Local nTotsRec := SE2->(RecCount())
Local aTamFor := TAMSX3("E2_FORNECE")
Local nDecs := Msdecimais(mv_par15)
Local lFr150Flt := EXISTBLOCK("FR150FLT")
Local cFr150Flt

#IFDEF TOP
	Local nI := 0
	Local aStru := SE2->(dbStruct())
#ENDIF

Private nRegSM0 := SM0->(Recno())
Private nAtuSM0 := SM0->(Recno())
PRIVATE dBaixa := dDataBase

//��������������������������������������������������������������Ŀ
//� Ponto de entrada para Filtrar 										  �
//��������������������������������������� Localiza��es Argentina��
If lFr150Flt
	cFr150Flt := EXECBLOCK("Fr150FLT",.f.,.f.)
Endif

//�����������������������������������������������������������Ŀ
//� Variaveis utilizadas para Impress�o do Cabe�alho e Rodap� �
//�������������������������������������������������������������
cbtxt  := OemToAnsi(STR0015)  //"* indica titulo provisorio, P indica Saldo Parcial"
cbcont := 0
li 	 := 80
m_pag  := 1

nOrdem := aReturn[8]
cMoeda := LTrim(Str(mv_par15))
Titulo += OemToAnsi(STR0035)+GetMv("MV_MOEDA"+cMoeda)  //" em "

dbSelectArea ( "SE2" )
Set Softseek On

If mv_par22 == 2
	cFilDe  := cFilAnt
	cFilAte := cFilAnt
ELSE
	cFilDe := mv_par23
	cFilAte:= mv_par24
Endif

//Acerta a database de acordo com o parametro
If mv_par21 == 1    // Considera Data Base
	dDataBase := mv_par33
Endif	

Private nVoltaPAVlr := 0

dbSelectArea("SM0")

cEmpAnt := SM0->M0_CODIGO
cFilAte := SM0->M0_CODFIL
cFilDe  := SM0->M0_CODFIL

dbSeek(cEmpAnt+cFilDe,.T.)

nRegSM0 := SM0->(Recno())
nAtuSM0 := SM0->(Recno())

While !Eof() .and. M0_CODIGO == cEmpAnt .and. M0_CODFIL <= cFilAte

	dbSelectArea("SE2")
	cFilAnt := SM0->M0_CODFIL

	#IFDEF TOP
		if TcSrvType() != "AS/400"
			cQuery := "SELECT * "
			cQuery += "  FROM "+	RetSqlName("SE2")
			cQuery += " WHERE E2_FILIAL = '" + xFilial("SE2") + "'"
			cQuery += "   AND D_E_L_E_T_ <> '*' "
		endif
	#ENDIF

	IF nOrdem == 1
		SE2->(dbSetOrder(1))
		#IFDEF TOP
			if TcSrvType() == "AS/400"
				dbSeek(xFilial("SE2")+mv_par03+mv_par01,.T.)
			else
				cOrder := SqlOrder(indexkey())
			endif
		#ELSE
			dbSeek(xFilial("SE2")+mv_par03+mv_par01,.T.)
		#ENDIF
		cCond1 := "E2_PREFIXO <= mv_par04"
		cCond2 := "E2_PREFIXO"
		titulo += OemToAnsi(STR0016)  //" - Por Numero"
		nQualIndice := 1
	Elseif nOrdem == 2
		SE2->(dbSetOrder(2))
		#IFDEF TOP
			if TcSrvType() == "AS/400"
				dbSeek(xFilial("SE2")+mv_par05,.T.)
			else
				cOrder := SqlOrder(indexkey())
			endif
		#ELSE
			dbSeek(xFilial("SE2")+mv_par05,.T.)
		#ENDIF
		cCond1 := "E2_NATUREZ <= mv_par06"
		cCond2 := "E2_NATUREZ"
		titulo += OemToAnsi(STR0017)  //" - Por Natureza"
		nQualIndice := 2
	Elseif nOrdem == 3
		SE2->(dbSetOrder(3))
		#IFDEF TOP
			if TcSrvType() == "AS/400"
				dbSeek(xFilial("SE2")+Dtos(mv_par07),.T.)
			else
				cOrder := SqlOrder(indexkey())
			endif
		#ELSE
			dbSeek(xFilial("SE2")+Dtos(mv_par07),.T.)
		#ENDIF
		cCond1 := "E2_VENCREA <= mv_par08"
		cCond2 := "E2_VENCREA"
		titulo += OemToAnsi(STR0018)  //" - Por Vencimento"
		nQualIndice := 3
	Elseif nOrdem == 4
		SE2->(dbSetOrder(4))
		#IFDEF TOP
			if TcSrvType() == "AS/400"
				dbSeek(xFilial("SE2")+mv_par09,.T.)
			else
				cOrder := SqlOrder(indexkey())
			endif
		#ELSE
			dbSeek(xFilial("SE2")+mv_par09,.T.)
		#ENDIF
		cCond1 := "E2_PORTADO <= mv_par10"
		cCond2 := "E2_PORTADO"
		titulo += OemToAnsi(STR0031)  //" - Por Banco"
		nQualIndice := 4
	Elseif nOrdem == 6
		SE2->(dbSetOrder(5))
		#IFDEF TOP
			if TcSrvType() == "AS/400"
				dbSeek(xFilial("SE2")+DTOS(mv_par13),.T.)
			else
				cOrder := SqlOrder(indexkey())
			endif
		#ELSE
			dbSeek(xFilial("SE2")+DTOS(mv_par13),.T.)
		#ENDIF
		cCond1 := "E2_EMISSAO <= mv_par14"
		cCond2 := "E2_EMISSAO"
		titulo += OemToAnsi(STR0019)  //" - Por Emissao"
		nQualIndice := 5
	Elseif nOrdem == 7
		SE2->(dbSetOrder(6))
		#IFDEF TOP
			if TcSrvType() == "AS/400"
				dbSeek(xFilial("SE2")+mv_par11,.T.)
			else
				cOrder := SqlOrder(indexkey())
			endif
		#ELSE
			dbSeek(xFilial("SE2")+mv_par11,.T.)
		#ENDIF			
		cCond1 := "E2_FORNECE <= mv_par12"
		cCond2 := "E2_FORNECE"
		titulo += OemToAnsi(STR0020)  //" - Por Cod.Fornecedor"
		nQualIndice := 6
	Else
		cChaveSe2 := "E2_FILIAL+E2_NOMFOR+E2_FORNECE+E2_LOJA+E2_PREFIXO+E2_NUM+E2_PARCELA+E2_TIPO"
		#IFDEF TOP
			if TcSrvType() == "AS/400"
				cIndexSe2 := CriaTrab(nil,.f.)
				IndRegua("SE2",cIndexSe2,cChaveSe2,,Fr150IndR(),OemToAnsi(STR0021))  // //"Selecionando Registros..."
				nIndexSE2 := RetIndex("SE2")
				dbSetOrder(nIndexSe2+1)
				dbSeek(xFilial("SE2"))
			else
				cOrder := SqlOrder(cChaveSe2)
			endif
		#ELSE
			cIndexSe2 := CriaTrab(nil,.f.)
			IndRegua("SE2",cIndexSe2,cChaveSe2,,Fr150IndR(),OemToAnsi(STR0021))  // //"Selecionando Registros..."
			nIndexSE2 := RetIndex("SE2")
			dbSetIndex(cIndexSe2+OrdBagExt())
			dbSetOrder(nIndexSe2+1)
			dbSeek(xFilial("SE2"))
		#ENDIF
		cCond1 := "E2_FORNECE <= mv_par12"
		cCond2 := "E2_FORNECE+E2_LOJA"
		titulo += OemToAnsi(STR0022)  //" - Por Fornecedor"
		nQualIndice := IndexOrd()
	EndIF

	If mv_par20 == 1
		titulo += OemToAnsi(STR0023)  //" - Analitico"
	Else
		titulo += OemToAnsi(STR0024)  // " - Sintetico"
		cabec1 := OemToAnsi(STR0033)  // "                                                                                          |        Titulos vencidos         |    Titulos a vencer     |           Valor dos juros ou                       (Vencidos+Vencer)"
		cabec2 := OemToAnsi(STR0034)  // "                                                                                          | Valor nominal   Valor corrigido |      Valor nominal      |            com. permanencia                                         "
	EndIf

	dbSelectArea("SE2")
	cFilterUser:=aReturn[7]

	Set Softseek Off

	#IFDEF TOP
		if TcSrvType() != "AS/400"
			cQuery += " AND E2_NUM     BETWEEN '"+ mv_par01+ "' AND '"+ mv_par02 + "'"
			cQuery += " AND E2_PREFIXO BETWEEN '"+ mv_par03+ "' AND '"+ mv_par04 + "'"
			cQuery += " AND (E2_MULTNAT = '1' OR (E2_NATUREZ BETWEEN '"+MV_PAR05+"' AND '"+MV_PAR06+"'))"
			cQuery += " AND E2_VENCREA BETWEEN '"+ DTOS(mv_par07)+ "' AND '"+ DTOS(mv_par08) + "'"
			cQuery += " AND E2_PORTADO BETWEEN '"+ mv_par09+ "' AND '"+ mv_par10 + "'"
			cQuery += " AND E2_FORNECE BETWEEN '"+ mv_par11+ "' AND '"+ mv_par12 + "'"
			cQuery += " AND E2_EMISSAO BETWEEN '"+ DTOS(mv_par13)+ "' AND '"+ DTOS(mv_par14) + "'"
			cQuery += " AND E2_LOJA    BETWEEN '"+ mv_par25 + "' AND '"+ mv_par26 + "'"

			//Considerar titulos cuja emissao seja maior que a database do sistema
			If mv_par36 == 2
				cQuery += " AND E2_EMISSAO <= '" + DTOS(dDataBase) +"'"
			Endif
	
			If !Empty(mv_par30) // Deseja imprimir apenas os tipos do parametro 30
				cQuery += " AND E2_TIPO IN "+FormatIn(mv_par30,";") 
			ElseIf !Empty(Mv_par31) // Deseja excluir os tipos do parametro 31
				cQuery += " AND E2_TIPO NOT IN "+FormatIn(mv_par31,";")
			EndIf
			If mv_par32 == 1
				cQuery += " AND E2_FLUXO <> 'N'"
			Endif
			cQuery += " ORDER BY "+ cOrder

			cQuery := ChangeQuery(cQuery)

			dbSelectArea("SE2")
			dbCloseArea()
			dbSelectArea("SA2")

			dbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery), 'SE2', .F., .T.)

			For ni := 1 to Len(aStru)
				If aStru[ni,2] != 'C'
					TCSetField('SE2', aStru[ni,1], aStru[ni,2],aStru[ni,3],aStru[ni,4])
				Endif
			Next
		endif
	#ENDIF

	//SetRegua(nTotsRec)
	
	If MV_MULNATP .And. nOrdem == 2
		//Finr155R3(cFr150Flt, lEnd, @nTot0, @nTot1, @nTot2, @nTot3, @nTotTit, @nTotJ )
		#IFDEF TOP
			if TcSrvType() != "AS/400"
				dbSelectArea("SE2")
				dbCloseArea()
				ChKFile("SE2")
				dbSetOrder(1)
			endif
		#ENDIF
		If Empty(xFilial("SE2"))
			Exit
		Endif
		dbSelectArea("SM0")
		dbSkip()
		Loop
	Endif

	While &cCond1 .and. !Eof() .and. lContinua .and. E2_FILIAL == xFilial("SE2")

		/*
		IF lEnd
			@PROW()+1,001 PSAY OemToAnsi(STR0025)  //"CANCELADO PELO OPERADOR"
			Exit
		End

		IncRegua()
        */
		dbSelectArea("SE2")

		Store 0 To nTit1,nTit2,nTit3,nTit4,nTit5

		//����������������������������������������Ŀ
		//� Carrega data do registro para permitir �
		//� posterior analise de quebra por mes.	 �
		//������������������������������������������
		dDataAnt := Iif(nOrdem == 3, SE2->E2_VENCREA, SE2->E2_EMISSAO)

		cCarAnt := &cCond2

		While &cCond2 == cCarAnt .and. !Eof() .and. lContinua .and. E2_FILIAL == xFilial("SE2")

			/*
			IF lEnd
				@PROW()+1,001 PSAY OemToAnsi(STR0025)  //"CANCELADO PELO OPERADOR"
				Exit
			End

			IncRegua()
            */
            
			dbSelectArea("SE2")

			//��������������������������������������������������������������Ŀ
			//� Considera filtro do usuario                                  �
			//����������������������������������������������������������������
			If !Empty(cFilterUser).and.!(&cFilterUser)
				dbSkip()
				Loop
			Endif
			//��������������������������������������������������������������Ŀ
			//� Considera filtro do usuario no ponto de entrada.             �
			//����������������������������������������������������������������
			If lFr150flt
				If &cFr150flt
					DbSkip()
					Loop
				Endif
			Endif
			//������������������������������������������������������Ŀ
			//� Verifica se trata-se de abatimento ou provisorio, ou �
			//� Somente titulos emitidos ate a data base				   �
			//��������������������������������������������������������
			IF SE2->E2_TIPO $ MVABATIM .Or. (SE2 -> E2_EMISSAO > dDataBase .and. mv_par36 == 2)
				dbSkip()
				Loop
			EndIF
			//������������������������������������������������������Ŀ
			//� Verifica se ser� impresso titulos provis�rios		   �
			//��������������������������������������������������������
			IF E2_TIPO $ MVPROVIS .and. mv_par16 == 2
				dbSkip()
				Loop
			EndIF

			//������������������������������������������������������Ŀ
			//� Verifica se ser� impresso titulos de Adiantamento	 �
			//��������������������������������������������������������
			IF SE2->E2_TIPO $ MVPAGANT+"/"+MV_CPNEG .and. mv_par27 == 2
				dbSkip()
				Loop
			EndIF

			// dDtContab para casos em que o campo E2_EMIS1 esteja vazio
			// compatibilizando com a vers�o 2.04 que n�o gerava para titulos
			// de ISS e FunRural

			dDtContab := Iif(Empty(SE2->E2_EMIS1),SE2->E2_EMISSAO,SE2->E2_EMIS1)

			//����������������������������������������Ŀ
			//� Verifica se esta dentro dos parametros �
			//������������������������������������������
			IF E2_NUM < mv_par01      .OR. E2_NUM > mv_par02 .OR. ;
					E2_PREFIXO < mv_par03  .OR. E2_PREFIXO > mv_par04 .OR. ;
					E2_NATUREZ < mv_par05  .OR. E2_NATUREZ > mv_par06 .OR. ;
					E2_VENCREA < mv_par07  .OR. E2_VENCREA > mv_par08 .OR. ;
					E2_PORTADO < mv_par09  .OR. E2_PORTADO > mv_par10 .OR. ;
					E2_FORNECE < mv_par11  .OR. E2_FORNECE > mv_par12 .OR. ;
					E2_EMISSAO < mv_par13  .OR. E2_EMISSAO > mv_par14 .OR. ;
					(E2_EMISSAO > dDataBase .and. mv_par36 == 2) .OR. dDtContab  < mv_par18 .OR. ;
					E2_LOJA    < mv_par25  .OR. E2_LOJA    > mv_par26 .OR. ;
					dDtContab  > mv_par19  .Or. !&(fr150IndR())

				dbSkip()
				Loop
			Endif

			//��������������������������������������������������������������Ŀ
			//� Verifica se titulo, apesar do E2_SALDO = 0, deve aparecer ou �
			//� n�o no relat�rio quando se considera database (mv_par21 = 1) �
			//� ou caso n�o se considere a database, se o titulo foi totalmen�
			//� te baixado.																  �
			//����������������������������������������������������������������
			dbSelectArea("SE2")
			IF !Empty(SE2->E2_BAIXA) .and. Iif(mv_par21 == 2 ,SE2->E2_SALDO == 0 ,;
					IIF(mv_par34 == 1,(SE2->E2_SALDO == 0 .and. SE2->E2_BAIXA <= dDataBase),.F.))
				dbSkip()
				Loop
			EndIF

			//����������������������������������������Ŀ
			//� Verifica se deve imprimir outras moedas�
			//������������������������������������������
			If mv_par29 == 2 // nao imprime
				if SE2->E2_MOEDA != mv_par15 //verifica moeda do campo=moeda parametro
					dbSkip()
					Loop
				endif	
			Endif
			
			 // Tratamento da correcao monetaria para a Argentina
			If  cPaisLoc=="ARG" .And. mv_par15 <> 1  .And.  SE2->E2_CONVERT=='N'
				dbSkip()
				Loop
			Endif
             

			dbSelectArea("SA2")
			MSSeek(xFilial("SA2")+SE2->E2_FORNECE+SE2->E2_LOJA)
			dbSelectArea("SE2")

			// Verifica se existe a taxa na data do vencimento do titulo, se nao existir, utiliza a taxa da database
			If SE2->E2_VENCREA < dDataBase
				If mv_par17 == 2 .And. RecMoeda(SE2->E2_VENCREA,cMoeda) > 0
					dDataReaj := SE2->E2_VENCREA
				Else
					dDataReaj := dDataBase
				EndIf	
			Else
				dDataReaj := dDataBase
			EndIf       

			If mv_par21 == 1
				//Ita - 03/03/2008 - nSaldo := SaldoTit(SE2->E2_PREFIXO,SE2->E2_NUM,SE2->E2_PARCELA,SE2->E2_TIPO,SE2->E2_NATUREZ,"P",SE2->E2_FORNECE,mv_par15,dDataReaj,,SE2->E2_LOJA,,If(mv_par35==1,SE2->E2_TXMOEDA,Nil),IIF(mv_par34 == 2,3,1)) // 1 = DT BAIXA    3 = DT DIGIT
				nSaldo := TitSaldo(SE2->E2_PREFIXO,SE2->E2_NUM,SE2->E2_PARCELA,SE2->E2_TIPO,SE2->E2_NATUREZ,"P",SE2->E2_FORNECE,mv_par15,dDataReaj,,SE2->E2_LOJA,,If(mv_par35==1,SE2->E2_TXMOEDA,Nil),IIF(mv_par34 == 2,3,1)) // 1 = DT BAIXA    3 = DT DIGIT
				// Subtrai decrescimo para recompor o saldo na data escolhida.
				If Str(SE2->E2_VALOR,17,2) == Str(nSaldo,17,2) .And. SE2->E2_DECRESC > 0 .And. SE2->E2_SDDECRE == 0
					nSAldo -= SE2->E2_DECRESC
				Endif
				// Soma Acrescimo para recompor o saldo na data escolhida.
				If Str(SE2->E2_VALOR,17,2) == Str(nSaldo,17,2) .And. SE2->E2_ACRESC > 0 .And. SE2->E2_SDACRES == 0
					nSAldo += SE2->E2_ACRESC
				Endif
			Else
				nSaldo := xMoeda((SE2->E2_SALDO+SE2->E2_SDACRES-SE2->E2_SDDECRE),SE2->E2_MOEDA,mv_par15,dDataReaj,ndecs+1,If(mv_par35==1,SE2->E2_TXMOEDA,Nil))
			Endif
			If ! (SE2->E2_TIPO $ MVPAGANT+"/"+MV_CPNEG) .And. ;
			   ! ( MV_PAR21 == 2 .And. nSaldo == 0 ) // nao deve olhar abatimento pois e zerado o saldo na liquidacao final do titulo

				//Quando considerar Titulos com emissao futura, eh necessario
				//colocar-se a database para o futuro de forma que a Somaabat()
				//considere os titulos de abatimento
				If mv_par36 == 1
					dOldData := dDataBase
					dDataBase := CTOD("31/12/40")
				Endif

				nSaldo-=SomaAbat(SE2->E2_PREFIXO,SE2->E2_NUM,SE2->E2_PARCELA,"P",mv_par15,dDataReaj,SE2->E2_FORNECE,SE2->E2_LOJA)

				If mv_par36 == 1
					dDataBase := dOldData
				Endif
			EndIf

			nSaldo:=Round(NoRound(nSaldo,3),2)
			//������������������������������������������������������Ŀ
			//� Desconsidera caso saldo seja menor ou igual a zero   �
			//��������������������������������������������������������
			If nSaldo <= 0
				dbSkip()
				Loop
			Endif

			IF li > 58
				nAtuSM0 := SM0->(Recno())
				SM0->(dbGoto(nRegSM0))
				//cabec(titulo,cabec1,cabec2,nomeprog,tamanho,GetMv("MV_COMP"))
				SM0->(dbGoTo(nAtuSM0))
			EndIF

			/*
			If mv_par20 == 1
				@li,	0 PSAY SE2->E2_FORNECE+"-"+SE2->E2_LOJA+"-"+IIF(mv_par28 == 1, SubStr(SA2->A2_NREDUZ,1,15), SubStr(SA2->A2_NOME,1,15))
				li := IIf (aTamFor[1] > 6,li+1,li)
				@li, 26 PSAY SE2->E2_PREFIXO+"-"+SE2->E2_NUM+"-"+SE2->E2_PARCELA
				@li, 47 PSAY SE2->E2_TIPO
				@li, 51 PSAY SE2->E2_NATUREZ
				@li, 62 PSAY SE2->E2_EMISSAO
				@Li, 73 PSAY SE2->E2_VENCTO
				@li, 84 PSAY SE2->E2_VENCREA
				@li, 96 PSAY xMoeda(SE2->E2_VALOR,SE2->E2_MOEDA,mv_par15,SE2->E2_EMISSAO,ndecs+1,If(mv_par35==1,SE2->E2_TXMOEDA,Nil)) * If(SE2->E2_TIPO$MV_CPNEG+"/"+MVPAGANT, -1,1) Picture PesqPict("SE2","E2_VALOR",14,MV_PAR15)
			EndIf
			*/
			#IFDEF TOP
				If TcSrvType() == "AS/400"
					dbSetOrder( nQualIndice )
				Endif
			#ELSE
				dbSetOrder( nQualIndice )
			#ENDIF

			If dDataBase > SE2->E2_VENCREA 		//vencidos
				/*
				If mv_par20 == 1
					@ li, 112 PSAY nSaldo  * If(SE2->E2_TIPO$MV_CPNEG+"/"+MVPAGANT, -1,1) Picture TM(nSaldo,14,nDecs)
				EndIf
				*/
				nJuros:=0
				dBaixa:=dDataBase
				fa080Juros(mv_par15)
				dbSelectArea("SE2")
				/*
				If mv_par20 == 1
					@li,129 PSAY (nSaldo+nJuros) * If(SE2->E2_TIPO$MV_CPNEG+"/"+MVPAGANT, -1,1) Picture TM(nSaldo+nJuros,14,nDecs)
				EndIf
				*/
				If SE2->E2_TIPO $ MVPAGANT+"/"+MV_CPNEG
					nTit0 -= xMoeda(SE2->E2_VALOR,SE2->E2_MOEDA,mv_par15,SE2->E2_EMISSAO,ndecs+1,If(mv_par35==1,SE2->E2_TXMOEDA,Nil))
					nTit1 -= nSaldo
					nTit2 -= nSaldo+nJuros
					nMesTit0 -= xMoeda(SE2->E2_VALOR,SE2->E2_MOEDA,mv_par15,SE2->E2_EMISSAO,ndecs+1,If(mv_par35==1,SE2->E2_TXMOEDA,Nil))
					nMesTit1 -= nSaldo
					nMesTit2 -= nSaldo+nJuros
				Else
					nTit0 += xMoeda(SE2->E2_VALOR,SE2->E2_MOEDA,mv_par15,SE2->E2_EMISSAO,ndecs+1,If(mv_par35==1,SE2->E2_TXMOEDA,Nil))
					nTit1 += nSaldo
					nTit2 += nSaldo+nJuros
					nMesTit0 += xMoeda(SE2->E2_VALOR,SE2->E2_MOEDA,mv_par15,SE2->E2_EMISSAO,ndecs+1,If(mv_par35==1,SE2->E2_TXMOEDA,Nil))
					nMesTit1 += nSaldo
					nMesTit2 += nSaldo+nJuros
				Endif
				nTotJur += (nJuros)
				nMesTitJ += (nJuros)
			Else				  //a vencer
				/*
				If mv_par20 == 1
					@li,147 PSAY nSaldo  * If(SE2->E2_TIPO$MV_CPNEG+"/"+MVPAGANT, -1,1) Picture TM(nSaldo,14,nDecs)
				EndIf
				*/
				If SE2->E2_TIPO $ MVPAGANT+"/"+MV_CPNEG
					nTit0 -= xMoeda(SE2->E2_VALOR,SE2->E2_MOEDA,mv_par15,SE2->E2_EMISSAO,ndecs+1,If(mv_par35==1,SE2->E2_TXMOEDA,Nil))
					nTit3 -= nSaldo
					nTit4 -= nSaldo
					nMesTit0 -= xMoeda(SE2->E2_VALOR,SE2->E2_MOEDA,mv_par15,SE2->E2_EMISSAO,ndecs+1,If(mv_par35==1,SE2->E2_TXMOEDA,Nil))
					nMesTit3 -= nSaldo
					nMesTit4 -= nSaldo
				Else
					nTit0 += xMoeda(SE2->E2_VALOR,SE2->E2_MOEDA,mv_par15,SE2->E2_EMISSAO,ndecs+1,If(mv_par35==1,SE2->E2_TXMOEDA,Nil))
					nTit3 += nSaldo
					nTit4 += nSaldo
					nMesTit0 += xMoeda(SE2->E2_VALOR,SE2->E2_MOEDA,mv_par15,SE2->E2_EMISSAO,ndecs+1,If(mv_par35==1,SE2->E2_TXMOEDA,Nil))
					nMesTit3 += nSaldo
					nMesTit4 += nSaldo
				Endif
			Endif
			/*
			If mv_par20 == 1
				@Li,165 PSAY SE2->E2_PORTADO
			EndIf
			*/
			If nJuros > 0
				/*
				If mv_par20 == 1
					@Li,173 PSAY nJuros Picture PesqPict("SE2","E2_JUROS",12,MV_PAR15)
				EndIf
				*/
				nJuros := 0
			Endif

			IF dDataBase > E2_VENCREA
				nAtraso:=dDataBase-E2_VENCTO
				IF Dow(E2_VENCTO) == 1 .Or. Dow(E2_VENCTO) == 7
					IF Dow(dBaixa) == 2 .and. nAtraso <= 2
						nAtraso:=0
					EndIF
				EndIF
				nAtraso:=IIF(nAtraso<0,0,nAtraso)
				IF nAtraso>0
					/*
					If mv_par20 == 1
						@li,189 PSAY nAtraso Picture "9999"
					EndIf
					*/
				EndIF
			EndIF
			If mv_par20 == 1
				//@li,194 PSAY SUBSTR(SE2->E2_HIST,1,24)+ ;
					IIF(E2_TIPO $ MVPROVIS,"*"," ")+ ;
					Iif(nSaldo - SE2->E2_ACRESC + SE2->E2_DECRESC == xMoeda(E2_VALOR,E2_MOEDA,mv_par15,dDataReaj,ndecs+1,If(mv_par35==1,SE2->E2_TXMOEDA,Nil))," ","P")
			EndIf

			//����������������������������������������Ŀ
			//� Carrega data do registro para permitir �
			//� posterior analise de quebra por mes.	 �
			//������������������������������������������
			dDataAnt := Iif(nOrdem == 3, SE2->E2_VENCREA, SE2->E2_EMISSAO)

			dbSkip()
			nTotTit ++
			nMesTTit ++
			nFiltit++
			nTit5 ++
			If mv_par20 == 1
				li ++
			EndIf

		EndDO

		IF nTit5 > 0 .and. nOrdem != 1
			SubTot150(nTit0,nTit1,nTit2,nTit3,nTit4,nOrdem,cCarAnt,nTotJur,nDecs)
			If mv_par20 == 1
				li++
			EndIf
		EndIF

		//����������������������������������������Ŀ
		//� Verifica quebra por mes					 �
		//������������������������������������������
		lQuebra := .F.
		If nOrdem == 3 .and. Month(SE2->E2_VENCREA) # Month(dDataAnt)
			lQuebra := .T.
		Elseif nOrdem == 6 .and. Month(SE2->E2_EMISSAO) # Month(dDataAnt)
			lQuebra := .T.
		Endif
		If lQuebra .and. nMesTTit # 0
			ImpMes150(nMesTit0,nMesTit1,nMesTit2,nMesTit3,nMesTit4,nMesTTit,nMesTitJ,nDecs)
			nMesTit0 := nMesTit1 := nMesTit2 := nMesTit3 := nMesTit4 := nMesTTit := nMesTitj := 0
		Endif

		dbSelectArea("SE2")

		nTot0 += nTit0
		nTot1 += nTit1
		nTot2 += nTit2
		nTot3 += nTit3
		nTot4 += nTit4
		nTotJ += nTotJur

		nFil0 += nTit0
		nFil1 += nTit1
		nFil2 += nTit2
		nFil3 += nTit3
		nFil4 += nTit4
		nFilJ += nTotJur
		Store 0 To nTit0,nTit1,nTit2,nTit3,nTit4,nTit5,nTotJur
	Enddo					
    nVoltaPAVlr := IIF(nFunPACham == 1,nFil4,nFil1)
	dbSelectArea("SE2")		// voltar para alias existente, se nao, nao funciona
	//����������������������������������������Ŀ
	//� Imprimir TOTAL por filial somente quan-�
	//� do houver mais do que 1 filial.        �
	//������������������������������������������
	if mv_par22 == 1 .and. SM0->(Reccount()) > 1
		//ImpFil150(nFil0,nFil1,nFil2,nFil3,nFil4,nFilTit,nFilj,nDecs)
	Endif
	Store 0 To nFil0,nFil1,nFil2,nFil3,nFil4,nFilTit,nFilJ
	If Empty(xFilial("SE2"))
		Exit
	Endif

	#IFDEF TOP
		if TcSrvType() != "AS/400"
			dbSelectArea("SE2")
			dbCloseArea()
			ChKFile("SE2")
			dbSetOrder(1)
		endif
	#ENDIF

	dbSelectArea("SM0")
	dbSkip()
EndDO

SM0->(dbGoTo(nRegSM0))
cFilAnt := SM0->M0_CODFIL

IF li != 80
	If mv_par20 == 1
		li +=2
	Endif
	//ImpTot150(nTot0,nTot1,nTot2,nTot3,nTot4,nTotTit,nTotJ,nDecs)
	cbcont := 1
	roda(cbcont,cbtxt,"G")
EndIF
//Set Device To Screen

#IFNDEF TOP
	dbSelectArea( "SE2" )
	dbClearFil()
	RetIndex( "SE2" )
	If !Empty(cIndexSE2)
		FErase (cIndexSE2+OrdBagExt())
	Endif
	dbSetOrder(1)
#ELSE
	if TcSrvType() != "AS/400"
		dbSelectArea("SE2")
		dbCloseArea()
		ChKFile("SE2")
		dbSetOrder(1)
	else
		dbSelectArea( "SE2" )
		dbClearFil()
		RetIndex( "SE2" )
		If !Empty(cIndexSE2)
			FErase (cIndexSE2+OrdBagExt())
		Endif
		dbSetOrder(1)
	endif
#ENDIF	
/*
If aReturn[5] = 1
	Set Printer To
	dbCommitAll()
	ourspool(wnrel)
Endif
MS_FLUSH()
*/
Return(nVoltaPAVlr)
/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 �SubTot150 � Autor � Wagner Xavier 		  � Data � 01.06.92 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o �IMPRIMIR SUBTOTAL DO RELATORIO 									  ���
�������������������������������������������������������������������������Ĵ��
���Sintaxe e � SubTot150() 															  ���
�������������������������������������������������������������������������Ĵ��
���Parametros� 																			  ���
�������������������������������������������������������������������������Ĵ��
��� Uso		 � Generico 																  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function SubTot150(nTit0,nTit1,nTit2,nTit3,nTit4,nOrdem,cCarAnt,nTotJur,nDecs)

DEFAULT nDecs := Msdecimais(mv_par15)

If mv_par20 == 1
	li++
EndIf

IF li > 58
	nAtuSM0 := SM0->(Recno())
	SM0->(dbGoto(nRegSM0))
	//cabec(titulo,cabec1,cabec2,nomeprog,tamanho,GetMv("MV_COMP"))
	SM0->(dbGoTo(nAtuSM0))
EndIF

if nOrdem == 1 .Or. nOrdem == 3 .Or. nOrdem == 4 .Or. nOrdem == 6
	//@li,000 PSAY OemToAnsi(STR0026)  //"S U B - T O T A L ----> "
	//@li,028 PSAY cCarAnt
ElseIf nOrdem == 2
	dbSelectArea("SED")
	dbSeek(xFilial("SED")+cCarAnt)
	//@li,000 PSAY cCarAnt +" "+SED->ED_DESCRIC
Elseif nOrdem == 5
	//@Li,000 PSAY IIF(mv_par28 == 1,SA2->A2_NREDUZ,SA2->A2_NOME)+" "+Substr(SA2->A2_TEL,1,15)
ElseIf nOrdem == 7
	//@li,000 PSAY SA2->A2_COD+" "+SA2->A2_LOJA+" "+SA2->A2_NOME+" "+Substr(SA2->A2_TEL,1,15)
Endif
if mv_par20 == 1
	//@li,096 PSAY nTit0		 Picture TM(nTit0,14,nDecs)
endif
/*
@li,112 PSAY nTit1		 Picture TM(nTit1,14,nDecs)
@li,129 PSAY nTit2		 Picture TM(nTit2,14,nDecs)
@li,147 PSAY nTit3		 Picture TM(nTit3,14,nDecs)
If nTotJur > 0
	@li,173 PSAY nTotJur 	 Picture TM(nTotJur,12,nDecs)
Endif
@li,204 PSAY nTit2+nTit3 Picture TM(nTit2+nTit3,16,nDecs)
li++
*/
Return(.T.)


/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 �ImpTot150 � Autor � Wagner Xavier 		  � Data � 01.06.92 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o �IMPRIMIR TOTAL DO RELATORIO 										  ���
�������������������������������������������������������������������������Ĵ��
���Sintaxe e � ImpTot150() 															  ���
�������������������������������������������������������������������������Ĵ��
���Parametros� 																			  ���
�������������������������������������������������������������������������Ĵ��
��� Uso		 � Generico 																  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
STATIC Function ImpTot150(nTot0,nTot1,nTot2,nTot3,nTot4,nTotTit,nTotJ,nDecs)

DEFAULT nDecs := Msdecimais(mv_par15)

li++
IF li > 58
	nAtuSM0 := SM0->(Recno())
	SM0->(dbGoto(nRegSM0))
	//cabec(titulo,cabec1,cabec2,nomeprog,tamanho,GetMv("MV_COMP"))
	SM0->(dbGoTo(nAtuSM0))
EndIF
//@li,000 PSAY OemToAnsi(STR0027)  //"T O T A L   G E R A L ----> "
//@li,028 PSAY "("+ALLTRIM(STR(nTotTit))+" "+IIF(nTotTit > 1,OemToAnsi(STR0028),OemToAnsi(STR0029))+")"  //"TITULOS"###"TITULO"
if mv_par20 == 1
//	@li,096 PSAY nTot0		 Picture TM(nTot0,14,nDecs)
endif
/*
@li,112 PSAY nTot1		 Picture TM(nTot1,14,nDecs)
@li,129 PSAY nTot2		 Picture TM(nTot2,14,nDecs)
@li,147 PSAY nTot3		 Picture TM(nTot3,14,nDecs)
@li,173 PSAY nTotJ		 Picture TM(nTotJ,12,nDecs)
@li,204 PSAY nTot2+nTot3 Picture TM(nTot2+nTot3,16,nDecs)
li+=2
*/
Return(.T.)

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 �ImpMes150 � Autor � Vinicius Barreira	  � Data � 12.12.94 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o �IMPRIMIR TOTAL DO RELATORIO - QUEBRA POR MES					  ���
�������������������������������������������������������������������������Ĵ��
���Sintaxe e � ImpMes150() 															  ���
�������������������������������������������������������������������������Ĵ��
���Parametros� 																			  ���
�������������������������������������������������������������������������Ĵ��
��� Uso		 � Generico 																  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
STATIC Function ImpMes150(nMesTot0,nMesTot1,nMesTot2,nMesTot3,nMesTot4,nMesTTit,nMesTotJ,nDecs)

DEFAULT nDecs := Msdecimais(mv_par15)

li++
IF li > 58
	nAtuSM0 := SM0->(Recno())
	SM0->(dbGoto(nRegSM0))
	//cabec(titulo,cabec1,cabec2,nomeprog,tamanho,GetMv("MV_COMP"))
	SM0->(dbGoTo(nAtuSM0))
EndIF
//@li,000 PSAY OemToAnsi(STR0030)  //"T O T A L   D O  M E S ---> "
//@li,028 PSAY "("+ALLTRIM(STR(nMesTTit))+" "+IIF(nMesTTit > 1,OemToAnsi(STR0028),OemToAnsi(STR0029))+")"  //"TITULOS"###"TITULO"
if mv_par20 == 1
  //	@li,096 PSAY nMesTot0   Picture TM(nMesTot0,14,nDecs)
endif
/*
@li,112 PSAY nMesTot1	Picture TM(nMesTot1,14,nDecs)
@li,129 PSAY nMesTot2	Picture TM(nMesTot2,14,nDecs)
@li,147 PSAY nMesTot3	Picture TM(nMesTot3,14,nDecs)
@li,173 PSAY nMesTotJ	Picture TM(nMesTotJ,12,nDecs)
@li,204 PSAY nMesTot2+nMesTot3 Picture TM(nMesTot2+nMesTot3,16,nDecs)
li+=2
*/
Return(.T.)

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 � ImpFil150� Autor � Paulo Boschetti 	     � Data � 01.06.92 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Imprimir total do relatorio										  ���
�������������������������������������������������������������������������Ĵ��
���Sintaxe e � ImpFil150()																  ���
�������������������������������������������������������������������������Ĵ��
���Parametros�																				  ���
�������������������������������������������������������������������������Ĵ��
��� Uso      � Generico				   									 			  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/
STATIC Function ImpFil150(nFil0,nFil1,nFil2,nFil3,nFil4,nFilTit,nFilJ,nDecs)

DEFAULT nDecs := Msdecimais(mv_par15)

li++
IF li > 58
	nAtuSM0 := SM0->(Recno())
	SM0->(dbGoto(nRegSM0))
//	cabec(titulo,cabec1,cabec2,nomeprog,tamanho,GetMv("MV_COMP"))
	SM0->(dbGoTo(nAtuSM0))
EndIF
/*
@li,000 PSAY OemToAnsi(STR0032)+" "+cFilAnt+" - " + AllTrim(SM0->M0_FILIAL)  //"T O T A L   F I L I A L ----> " 
if mv_par20 == 1
	@li,096 PSAY nFil0        Picture TM(nFil0,14,nDecs)
endif
@li,112 PSAY nFil1        Picture TM(nFil1,14,nDecs)
@li,129 PSAY nFil2        Picture TM(nFil2,14,nDecs)
@li,147 PSAY nFil3        Picture TM(nFil3,14,nDecs)
@li,173 PSAY nFilJ		  Picture TM(nFilJ,12,nDecs)
@li,204 PSAY nFil2+nFil3 Picture TM(nFil2+nFil3,16,nDecs)
li+=2
*/
Return .T.

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 �fr150Indr � Autor � Wagner           	  � Data � 12.12.94 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o �Monta Indregua para impressao do relat�rio						  ���
�������������������������������������������������������������������������Ĵ��
��� Uso		 � Generico 																  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/
Static Function fr150IndR()
Local cString
//������������������������������������������������������������Ŀ
//� ATENCAO !!!!                                               �
//� N�o adiconar mais nada a chave do filtro pois a mesma est� �
//� com 254 caracteres.                                        �
//��������������������������������������������������������������
cString := 'E2_FILIAL="'+xFilial()+'".And.'
cString += '(E2_MULTNAT="1" .OR. (E2_NATUREZ>="'+mv_par05+'".and.E2_NATUREZ<="'+mv_par06+'")).And.'
cString += 'E2_FORNECE>="'+mv_par11+'".and.E2_FORNECE<="'+mv_par12+'".And.'
cString += 'DTOS(E2_VENCREA)>="'+DTOS(mv_par07)+'".and.DTOS(E2_VENCREA)<="'+DTOS(mv_par08)+'".And.'
cString += 'DTOS(E2_EMISSAO)>="'+DTOS(mv_par13)+'".and.DTOS(E2_EMISSAO)<="'+DTOS(mv_par14)+'"'
If !Empty(mv_par30) // Deseja imprimir apenas os tipos do parametro 30
	cString += '.And.E2_TIPO$"'+mv_par30+'"'
ElseIf !Empty(Mv_par31) // Deseja excluir os tipos do parametro 31
	cString += '.And.!(E2_TIPO$'+'"'+mv_par31+'")'
EndIf
IF mv_par32 == 1  // Apenas titulos que estarao no fluxo de caixa
	cString += '.And.(E2_FLUXO!="N")'	
Endif
		
Return cString
