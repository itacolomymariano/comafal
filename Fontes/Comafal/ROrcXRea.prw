#INCLUDE "rwmake.ch"
#INCLUDE "TopConn.ch"

/*/
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????ͻ??
???Programa  ?ROrcXRea  ? Autor ? Five Solutions     ? Data ?  28/05/2008 ???
?????????????????????????????????????????????????????????????????????????͹??
???Descricao ? Mapa Comparativo de Naturezas - Or?ados X Real por Blocos  ???
???          ? de Naturezas.                                              ???
?????????????????????????????????????????????????????????????????????????͹??
???Uso       ? COMAFAL - PE/SP/RS                                         ???
?????????????????????????????????????????????????????????????????????????ͼ??
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
/*/

User Function ROrcXRea


//?????????????????????????????????????????????????????????????????????Ŀ
//? Declaracao de Variaveis                                             ?
//???????????????????????????????????????????????????????????????????????

Local cDesc1         := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2         := "de acordo com os parametros informados pelo usuario."
Local cDesc3         := "Mapa Comparativo de Naturezas - Or?ados X Real"
Local cPict          := ""
Local titulo       := "Mapa Comparativo de Naturezas - Or?ados X Real"
Local nLin         := 80

Local Cabec1       := "                                                                                                        G A S T O S   G E R A I S   D O   A D M I N I S T R A T I V O                                                       "
Local Cabec2       := "Natureza   Descri??o                                       Janeiro     Fevereiro         Mar?o         Abril          Maio         Junho         Julho        Agosto      Setembro       Outubro      Novembro      Dezembro"
Local imprime      := .T.
Local aOrd := {}
Private lEnd         := .F.
Private lAbortPrint  := .F.
Private CbTxt        := ""
Private limite           := 220
Private tamanho          := "G"
Private nomeprog         := "ROrcXRea" // Coloque aqui o nome do programa para impressao no cabecalho
Private nTipo            := 15
Private aReturn          := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey        := 0
Private cPerg       := "ORCREA"
Private cbtxt      := Space(10)
Private cbcont     := 00
Private CONTFL     := 01
Private m_pag      := 01
Private wnrel      := "ROrcXRea" // Coloque aqui o nome do arquivo usado para impressao em disco

Private cString := "SE7"


dDataRel := CTOD("")

/******************************************************************************
* Five Solutions Consultoria
* 15 de julho de 2008
* Melhorias: Implementado novas Linhas de Indicadores
* A Realizar: Titulos com Saldos em Aberta das Naturezas indicadas.
* Saldo     : Do Or?ado subtrair o Realizado e o A Realizar.
* Alterado o Indicador Percentual de Varia??o para Percentual de Utiliza??o
* Solicita??o : Patricia Melo e Roberto Lira
*******************************************************************************/

//Variaveis de Totalizadores Gastos Gerais/Metas
//Or?ado Gastos Gerais
nOrc1GADM := 0
nOrc2GADM := 0
nOrc3GADM := 0
nOrc4GADM := 0
nOrc5GADM := 0
nOrc6GADM := 0
nOrc7GADM := 0
nOrc8GADM := 0
nOrc9GADM := 0
nOrc10GADM := 0
nOrc11GADM := 0
nOrc12GADM := 0

//Variaveis de Totalizadores Gastos Gerais/Metas
//Realizado Gastos Gerais
nRea1GADM := 0
nRea2GADM := 0
nRea3GADM := 0
nRea4GADM := 0
nRea5GADM := 0
nRea6GADM := 0
nRea7GADM := 0
nRea8GADM := 0
nRea9GADM := 0
nRea10GADM := 0
nRea11GADM := 0
nRea12GADM := 0

//Realizado Gastos Gerais
naRea1GADM := 0
naRea2GADM := 0
naRea3GADM := 0
naRea4GADM := 0
naRea5GADM := 0
naRea6GADM := 0
naRea7GADM := 0
naRea8GADM := 0
naRea9GADM := 0
naRea10GADM := 0
naRea11GADM := 0
naRea12GADM := 0

//Saldo(Or?ado - (Rrealizado + A Realizar))
nSld1GG     := 0
nSld2GG     := 0
nSld3GG     := 0
nSld4GG     := 0
nSld5GG     := 0
nSld6GG     := 0
nSld7GG     := 0
nSld8GG     := 0
nSld9GG     := 0
nSld10GG    := 0
nSld11GG    := 0
nSld12GG    := 0

//Variaveis de Totalizadores Global Gastos Gerais/Metas
//Or?ado Global Gastos Gerais
nGlbOr1GG := 0
nGlbOr2GG := 0
nGlbOr3GG := 0
nGlbOr4GG := 0
nGlbOr5GG := 0
nGlbOr6GG := 0
nGlbOr7GG := 0
nGlbOr8GG := 0
nGlbOr9GG := 0
nGlbOr10GG := 0
nGlbOr11GG := 0
nGlbOr12GG := 0

//Variaveis de Totalizadores Global Gastos Gerais/Metas
//Realizado Global Gastos Gerais
nGbRea1GG := 0
nGbRea2GG := 0
nGbRea3GG := 0
nGbRea4GG := 0
nGbRea5GG := 0
nGbRea6GG := 0
nGbRea7GG := 0
nGbRea8GG := 0
nGbRea9GG := 0
nGbRea10GG := 0
nGbRea11GG := 0
nGbRea12GG := 0

//Variaveis de Totalizadores Global Gastos Gerais/Metas
//A Realizar Global Gastos Gerais
nGbaRea1GG := 0
nGbaRea2GG := 0
nGbaRea3GG := 0
nGbaRea4GG := 0
nGbaRea5GG := 0
nGbaRea6GG := 0
nGbaRea7GG := 0
nGbaRea8GG := 0
nGbaRea9GG := 0
nGbaRea10GG := 0
nGbaRea11GG := 0
nGbaRea12GG := 0

//Variaveis de Totalizadores Global Gastos Gerais/Metas
//Saldo(Or?ado - (Rrealizado + A Realizar))
nGbSld1GG     := 0
nGbSld2GG     := 0
nGbSld3GG     := 0
nGbSld4GG     := 0
nGbSld5GG     := 0
nGbSld6GG     := 0
nGbSld7GG     := 0
nGbSld8GG     := 0
nGbSld9GG     := 0
nGbSld10GG    := 0
nGbSld11GG    := 0
nGbSld12GG    := 0
/*

//Cria??o do Arquivo Tempor?rio
If SELECT("OXR") > 0
   DbSelectArea("OXR")
   DbCloseArea()
EndIf

cArqInd := "ORCXREA.CDX"
While File(cArqInd)
   DELETE FILE &cArqInd
EndDo

aCampos := {}
//Campos para guardar Metas Or?adas para cada natureza
Aadd(aCampos, {"NATUREZ","C",6,0})
Aadd(aCampos, {"DESC","C",30,0})
Aadd(aCampos, {"MES1","N",14,02})
Aadd(aCampos, {"MES2","N",14,02})
Aadd(aCampos, {"MES3","N",14,02})
Aadd(aCampos, {"MES4","N",14,02})
Aadd(aCampos, {"MES5","N",14,02})
Aadd(aCampos, {"MES6","N",14,02})
Aadd(aCampos, {"MES7","N",14,02})
Aadd(aCampos, {"MES8","N",14,02})
Aadd(aCampos, {"MES9","N",14,02})
Aadd(aCampos, {"MES10","N",14,02})
Aadd(aCampos, {"MES11","N",14,02})
Aadd(aCampos, {"MES12","N",14,02})
//Campos para guardar os Gastos Realizados para cada natureza
Aadd(aCampos, {"REA1","N",14,02})
Aadd(aCampos, {"REA2","N",14,02})
Aadd(aCampos, {"REA3","N",14,02})
Aadd(aCampos, {"REA4","N",14,02})
Aadd(aCampos, {"REA5","N",14,02})
Aadd(aCampos, {"REA6","N",14,02})
Aadd(aCampos, {"REA7","N",14,02})
Aadd(aCampos, {"REA8","N",14,02})
Aadd(aCampos, {"REA9","N",14,02})
Aadd(aCampos, {"REA10","N",14,02})
Aadd(aCampos, {"REA11","N",14,02})
Aadd(aCampos, {"REA12","N",14,02})
//Campo para identificar que a natureza do Adm tem Meta de Or?amento
Aadd(aCampos, {"MTORCADM","C",01,00})


DbCreate("ORCXREA.DBF",aCampos)

DbUseArea(.T.,"DBFCDX","ORCXREA.DBF","OXR")

INDEX ON OXR->NATUREZ TO ORCXREA

*/


//TOTAL DE GASTOS DA ?REA COMERCIAL
Private cGGNatComl := "100120,100201,100202,100203,100204,100206,100207,100208,100209,100210,100211,100212,100213,100214,100215,100216,100222,"
        cGGNatComl += "200201,200202,200203,200204,200205,200206,200207,600101,600102,600303,"
        cGGNatComl += "100217" //Five Solutions Consultoria - 26/08/2008 - Solicita??o: Patr?cia M?llo/Roberto Lira

//Naturezas de GASTOS META ORCAMENTO COML vers?o de Naturezas - Patr?cia 13/02/2008
Private cNtGMtOrCml:= "100204,100206,100207,100208,100209,100214,200201,200202,200203,200205,200206,600303,"
        cNtGMtOrCml+= "100217"//Five Solutions Consultoria - 26/08/2008 Solicita??o: Patr?cia M?llo/Roberto Lira.

//GASTOS GERIAS COMERCIO EXTERIOR
Private cGGCENat := "100205,600401"
//METAS DE OR?AMENTO PARA O COM?RCIO EXTERIOR
Private cMtCENat := "100205,600401"


//Nova vers?o de naturezas - Patr?cia 13/02/2008
//TOTAL DE GASTOS REALIZADOS PELA INDUSTRIA
Private cGASTGERIND := "100301,100302,100303,100304,100305,100306,100307,100308,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100408,100409,100410,100411,100412,100417,200109,"
        cGASTGERIND += "200110,200111,200132,200135,200301,200302,200303,200304,200305,200306,200307,200308,200309,200310,200311,200312,200313,200314,200315,200316,200317,200318,200319,200320,200321,200322,200360,200361,"
        cGASTGERIND += "200362,200363,200364,200365,200366,200367,200368,200369,200370,200372,200373,200374,200375,200376,200377,200401,200402,200403,200404,200405,200406,200407,200408,200410,200411,200412,200416,200417,"
        cGASTGERIND += "200418,200419,200420,200421,600201,600301,600501,600601,700301,700303,700304,700306,700307,700308,700309,"
        cGASTGERIND += "100315" //Five Solutions Consultoria - 26/08/2008 Soliti??o: Patricia M?llo/Roberto Lira

//Nova vers?o Naturezas - Patr?cia 13/02/2008
//TOTAL DE GASTOS REALIZADOS NAS NATUREZAS QUE COMP?E AS METAS DE OR?AMENTO DA INDUSTRIA
Private cNtGstMtOrInd := "100303,100305,100306,100307,100311,100312,100403,100405,100406,100407,100411,100412,200302,200304,200305,200308,200309,200310,200311,200312,200316,200317,200318,200319,200320,200321,"
        cNtGstMtOrInd += "200402,200404,200405,200408,200410,200411,200412,200416,200417,200418,200419,200420,200421,600301,"
        cNtGstMtOrInd += "100315"//Five Solutions Consultoria - 26/08/2008 - Solicita??o: Patr?cia M?llo/Roberto Lira



//Novas Naturezas do GASTOS GERAIS DO ADM em 28/05/2008 - Disponibilizada por Patricia.
Private cGstGerADM  := "100101,100102,100103,100104,100105,100106,100107,100108,100109,100110,100111,100112,100113,100114,100115,100116,100117,100118,100119,100121,"
        cGstGerADM  += "100122,200101,200102,200103,200104,200105,200106,200107,200108,200112,200113,200114,200115,200116,200117,200118,200119,200120,200121,200122,"
        cGstGerADM  += "200123,200124,200126,200127,200128,200129,200130,200131,200133,200134,200136,200137,200138,200139,200140,200141,200142,200143,200208,200211,"
        cGstGerADM  += "200501,200502,200503,200504,200505,200506,200507,200508,200509,200510,500101,500102,500103,500104,600302,700206,700310,COFINS,CSLL,INSS,IRF,ISS,PIS"

//NATUREZAS QUE COMP?EM AS METAS DE OR?AMENTO ADM
//Five Solutions - 28/08/2008 - Conforme Solicita??o de Patr?cial M?lo e Autoriza??o da Diretoria COMAFAL
// Foram retiradas as naturezas 500102, 500103 e 500104?, respectivamente ?Despesas de Viagem da Diretoria, INSS da Diretoria e Plano de Sa?de da Diretoria? das metas de or?amento do Administrativo
/*
Private AdmMetasOrc := "100103,100105,100106,100107,100111,100112,100114,100115,100116,100117,200102,200104,200105,200107,200112,200113,200114,200116,200117,200118,200121,200123,200124,200129,200131,"
        AdmMetasOrc += "200133,200134,200136,200143,200501,200502,200503,200504,200505,200506,200507,200508,200509,200510,500102,500103,500104,600302"
        */
Private AdmMetasOrc := "100103,100105,100106,100107,100111,100112,100114,100115,100116,100117,200102,200104,200105,200107,200112,200113,200114,200116,200117,200118,200121,200123,200124,200129,200131,"
        AdmMetasOrc += "200133,200134,200136,200143,200501,200502,200503,200504,200505,200506,200507,200508,200509,200510,600302"

//GASTOS GERAIS DE NATUREZAS DIVERSAS
Private cGGDiverNat := "700302,800401,900601,DESCONT,NDF,SANGRIA,TROCO"
//METAS OR?AMENTOS DE NATUREZAS DIVERSAS
Private cMtDiverNat := "700302,800401,900601,DESCONT,NDF,SANGRIA,TROCO"

dbSelectArea("SE7")
dbSetOrder(1)


AjstSx1()

Pergunte(cPerg,.F.)

//?????????????????????????????????????????????????????????????????????Ŀ
//? Monta a interface padrao com o usuario...                           ?
//???????????????????????????????????????????????????????????????????????

wnrel := SetPrint(cString,NomeProg,cPerg,@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,.F.,Tamanho,,.F.)

dDataRel := MV_PAR01
nDeptos  := MV_PAR02
lDiversas  := IIF(MV_PAR03 == 1,.T.,.F.)

lComercio := IIF(nDeptos == 1 .Or. nDeptos == 2,.T.,.F.)
lIndustria := IIF(nDeptos == 1 .Or. nDeptos == 3,.T.,.F.)
lAdmFin    := IIF(nDeptos == 1 .Or. nDeptos == 4,.T.,.F.)
lComExter  := IIF(nDeptos == 1 .Or. nDeptos == 5,.T.,.F.)


titulo       := "Mapa Comparativo de Naturezas - Or?ados X Real - Emp/Fil: "+Alltrim(SM0->M0_NOME)+"/"+Alltrim(SM0->M0_FILIAL)

If nLastKey == 27
	Return
Endif

SetDefault(aReturn,cString)

If nLastKey == 27
   Return
Endif

nTipo := If(aReturn[4]==1,15,18)

//?????????????????????????????????????????????????????????????????????Ŀ
//? Processamento. RPTSTATUS monta janela com a regua de processamento. ?
//???????????????????????????????????????????????????????????????????????

RptStatus({|| RunReport(Cabec1,Cabec2,Titulo,nLin) },Titulo)
Return

/*/
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????ͻ??
???Fun??o    ?RUNREPORT ? Autor ? AP6 IDE            ? Data ?  28/05/08   ???
?????????????????????????????????????????????????????????????????????????͹??
???Descri??o ? Funcao auxiliar chamada pela RPTSTATUS. A funcao RPTSTATUS ???
???          ? monta a janela com a regua de processamento.               ???
?????????????????????????????????????????????????????????????????????????͹??
???Uso       ? Programa principal                                         ???
?????????????????????????????????????????????????????????????????????????ͼ??
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
/*/

Static Function RunReport(Cabec1,Cabec2,Titulo,nLin)

Local nOrdem

dbSelectArea(cString)
dbSetOrder(1)

//?????????????????????????????????????????????????????????????????????Ŀ
//? SETREGUA -> Indica quantos registros serao processados para a regua ?
//???????????????????????????????????????????????????????????????????????
//SetRegua(RecCount())

If lAdmFin

   //Cria??o do Arquivo Tempor?rio
   If SELECT("OXR") > 0
      DbSelectArea("OXR")
      DbCloseArea()
   EndIf

   cArqInd := "ORCXREA.CDX"
   While File(cArqInd)
      DELETE FILE &cArqInd
   EndDo

   aCampos := {}
   //Campos para guardar Metas Or?adas para cada natureza
   Aadd(aCampos, {"NATUREZ","C",6,0})
   Aadd(aCampos, {"DESC","C",30,0})
   Aadd(aCampos, {"MES1","N",14,02})
   Aadd(aCampos, {"MES2","N",14,02})
   Aadd(aCampos, {"MES3","N",14,02})
   Aadd(aCampos, {"MES4","N",14,02})
   Aadd(aCampos, {"MES5","N",14,02})
   Aadd(aCampos, {"MES6","N",14,02})
   Aadd(aCampos, {"MES7","N",14,02})
   Aadd(aCampos, {"MES8","N",14,02})
   Aadd(aCampos, {"MES9","N",14,02})
   Aadd(aCampos, {"MES10","N",14,02})
   Aadd(aCampos, {"MES11","N",14,02})
   Aadd(aCampos, {"MES12","N",14,02})
   //Campos para guardar os Gastos Realizados para cada natureza
   Aadd(aCampos, {"REA1","N",14,02})
   Aadd(aCampos, {"REA2","N",14,02})
   Aadd(aCampos, {"REA3","N",14,02})
   Aadd(aCampos, {"REA4","N",14,02})
   Aadd(aCampos, {"REA5","N",14,02})
   Aadd(aCampos, {"REA6","N",14,02})
   Aadd(aCampos, {"REA7","N",14,02})
   Aadd(aCampos, {"REA8","N",14,02})
   Aadd(aCampos, {"REA9","N",14,02})
   Aadd(aCampos, {"REA10","N",14,02})
   Aadd(aCampos, {"REA11","N",14,02})
   Aadd(aCampos, {"REA12","N",14,02})
   //Campos para guardar os Gastos A Realizar para cada natureza
   Aadd(aCampos, {"AREA1","N",14,02})
   Aadd(aCampos, {"AREA2","N",14,02})
   Aadd(aCampos, {"AREA3","N",14,02})
   Aadd(aCampos, {"AREA4","N",14,02})
   Aadd(aCampos, {"AREA5","N",14,02})
   Aadd(aCampos, {"AREA6","N",14,02})
   Aadd(aCampos, {"AREA7","N",14,02})
   Aadd(aCampos, {"AREA8","N",14,02})
   Aadd(aCampos, {"AREA9","N",14,02})
   Aadd(aCampos, {"AREA10","N",14,02})
   Aadd(aCampos, {"AREA11","N",14,02})
   Aadd(aCampos, {"AREA12","N",14,02})
   //Campo para identificar que a natureza do Adm tem Meta de Or?amento
   Aadd(aCampos, {"MTORCADM","C",01,00})


   DbCreate("ORCXREA.DBF",aCampos)

   DbUseArea(.T.,"DBFCDX","ORCXREA.DBF","OXR")

   INDEX ON OXR->NATUREZ TO ORCXREA

   /************************
   *
   * GASTOS GERAIS ADM. (OR?ADO)
   *
   ************************/
   
   //ProcRegua(3)
   
   //For nMes := 1 To 3
      
   //   IncProc("Gastos Metas/Orc Comercio(Or?ado) Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQuery := " SELECT SED.ED_CODIGO,SE7.E7_ANO,SE7.E7_VALJAN1,SE7.E7_VALFEV1,SE7.E7_VALMAR1,SE7.E7_VALABR1," 
      cQuery += " SE7.E7_VALMAI1,SE7.E7_VALJUN1,SE7.E7_VALJUL1,SE7.E7_VALAGO1,SE7.E7_VALSET1,SE7.E7_VALOUT1,"
      cQuery += " SE7.E7_VALNOV1,SE7.E7_VALDEZ1"
      cQuery += " FROM "+RetSQLName("SED")+" SED LEFT OUTER JOIN "+RetSqlName("SE7")+" SE7 "
      cQuery += "                                ON SED.ED_CODIGO = SE7.E7_NATUREZ"
      cQuery += "                               AND SE7.E7_ANO = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      cQuery += "                               AND SE7.E7_FILIAL = '"+xFilial("SE7")+"'"
      cQuery += "                               AND SE7.D_E_L_E_T_ = ' '"
      cQuery += " WHERE SED.D_E_L_E_T_ = ' '"
      cQuery += "   AND SUBSTRING(SED.ED_CODIGO,1,6) IN "+FormatIn(cGstGerADM,",")      
      
      MemoWrite("C:\TEMP\ROrcXRea_Or?aGerADM.SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "ORGAD"
      
      TcSetField("ORGAD","E7_VALJAN1","N",14,02)
      TcSetField("ORGAD","E7_VALFEV1","N",14,02)
      TcSetField("ORGAD","E7_VALMAR1","N",14,02)
      TcSetField("ORGAD","E7_VALABR1","N",14,02)
      TcSetField("ORGAD","E7_VALMAI1","N",14,02)
      TcSetField("ORGAD","E7_VALJUN1","N",14,02)
      TcSetField("ORGAD","E7_VALJUL1","N",14,02)
      TcSetField("ORGAD","E7_VALAGO1","N",14,02)
      TcSetField("ORGAD","E7_VALSET1","N",14,02)
      TcSetField("ORGAD","E7_VALOUT1","N",14,02)
      TcSetField("ORGAD","E7_VALNOV1","N",14,02)
      TcSetField("ORGAD","E7_VALDEZ1","N",14,02)
      
      DbSelectArea("ORGAD")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(ORGAD->ED_CODIGO)
            RecLock("OXR",.F.)
               OXR->MES1 += ORGAD->E7_VALJAN1
               OXR->MES2 += ORGAD->E7_VALFEV1               
               OXR->MES3 += ORGAD->E7_VALMAR1
               OXR->MES4 += ORGAD->E7_VALABR1
               OXR->MES5 += ORGAD->E7_VALMAI1
               OXR->MES6 += ORGAD->E7_VALJUN1
               OXR->MES7 += ORGAD->E7_VALJUL1
               OXR->MES8 += ORGAD->E7_VALAGO1
               OXR->MES9 += ORGAD->E7_VALSET1
               OXR->MES10 += ORGAD->E7_VALOUT1
               OXR->MES11 += ORGAD->E7_VALNOV1
               OXR->MES12 += ORGAD->E7_VALDEZ1
              MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := ORGAD->ED_CODIGO
               OXR->MES1 := ORGAD->E7_VALJAN1
               OXR->MES2 := ORGAD->E7_VALFEV1               
               OXR->MES3 := ORGAD->E7_VALMAR1
               OXR->MES4 := ORGAD->E7_VALABR1
               OXR->MES5 := ORGAD->E7_VALMAI1
               OXR->MES6 := ORGAD->E7_VALJUN1
               OXR->MES7 := ORGAD->E7_VALJUL1
               OXR->MES8 := ORGAD->E7_VALAGO1
               OXR->MES9 := ORGAD->E7_VALSET1
               OXR->MES10 := ORGAD->E7_VALOUT1
               OXR->MES11 := ORGAD->E7_VALNOV1
               OXR->MES12 := ORGAD->E7_VALDEZ1
              MsUnLock()
         EndIf
         
         DbSelectArea("ORGAD")
         DbSkip()
         
      EndDo
      DbSelectArea("ORGAD")
      DbCloseArea()

   //Next nMes

   /************************
   *
   * GASTOS GERAIS ADM. (REALIZADO)
   *
   ************************/
   
   //ProcRegua(12)
   //For nMes := 1 To 12 //1 A 3 - Gastos Gerais Adm.
                      
   //   IncProc(" Calc.Gastos Gerais Adm. Proc. "+Alltrim(Str(nMes))+" / 12")
      cQuery := " SELECT SUBSTRING(SE5.E5_DATA,1,6) AS MESGST,SUBSTRING(SE5.E5_NATUREZ,1,6) AS NATUREZ,SUM(SE5.E5_VALOR) AS VLGGADM "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,4) = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      
      cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cGstGerADM,",")

      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"

      //Implementado em 08/07/2008 - Solicita??o: Patricia
      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCF','PA ','NDF')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"
      cQuery += "  AND SE5.E5_MOTBX <> 'DEV'"


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
      cQuery += " GROUP BY SUBSTRING(SE5.E5_DATA,1,6),SUBSTRING(SE5.E5_NATUREZ,1,6)"
      cQuery += " ORDER BY SUBSTRING(SE5.E5_NATUREZ,1,6)"
      
      MemoWrite("C:\TEMP\ROrcXRea_GastGerADM.SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "GADM"
      
      TcSetField("GADM","VLGGADM","N",14,02)
      
      DbSelectArea("GADM")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(GADM->NATUREZ)
            RecLock("OXR",.F.)
               If Substr(GADM->MESGST,5,2) == "01"
                  OXR->REA1 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "02"
                  OXR->REA2 += GADM->VLGGADM               
               ElseIf Substr(GADM->MESGST,5,2) == "03"
                  OXR->REA3 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "04"
                  OXR->REA4 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "05"
                  OXR->REA5 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "06"
                  OXR->REA6 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "07"
                  OXR->REA7 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "08"
                  OXR->REA8 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "09"
                  OXR->REA9 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "10"
                  OXR->REA10 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "11"
                  OXR->REA11 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "12"
                  OXR->REA12 += GADM->VLGGADM
               EndIf
            MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := GADM->NATUREZ
               If Substr(GADM->MESGST,5,2) == "01"
                  OXR->REA1 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "02"
                  OXR->REA2 := GADM->VLGGADM               
               ElseIf Substr(GADM->MESGST,5,2) == "03"
                  OXR->REA3 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "04"
                  OXR->REA4 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "05"
                  OXR->REA5 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "06"
                  OXR->REA6 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "07"
                  OXR->REA7 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "08"
                  OXR->REA8 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "09"
                  OXR->REA9 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "10"
                  OXR->REA10 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "11"
                  OXR->REA11 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "12"
                  OXR->REA12 := GADM->VLGGADM
               EndIf
            MsUnLock()
         EndIf
         
         DbSelectArea("GADM")
         DbSkip()
         
      EndDo
      DbSelectArea("GADM")
      DbCloseArea()
      
   //Next nMes   


*******************************************************************************************************************
*******************************************************************************************************************

   /************************
   *
   * GASTOS GERAIS ADM. (A REALIZAR)
   *
   ************************/

   //For nMes := 1 To 3
      
   //   IncProc("Calc. Pagamentos em Aberto, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryAP := "SELECT SUBSTRING(SE2.E2_NATUREZ,1,6) NATUREZ,SUBSTRING(SE2.E2_VENCREA,1,6) MESAREA,SUM(SE2.E2_SALDO) AS PAGABERTO"
      cQryAP += " FROM "+RetSQLName("SE2")+" SE2"
      cQryAP += " WHERE SE2.E2_FILIAL = '"+xFilial("SE2")+"'"

      cQryAP += "  AND SE2.E2_TIPO NOT IN ('NCF','PA ','NDF')"
      cQryAP += "  AND SUBSTRING(SE2.E2_TIPO,3,1) <> '-'"

      cQryAP += "  AND  SUBSTRING(SE2.E2_VENCREA,1,4) = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      cQryAP += "  AND SUBSTRING(SE2.E2_NATUREZ,1,6) IN "+FormatIn(cGstGerADM,",")      
      
      cQryAP += "  AND SE2.E2_SALDO > 0 " //Titulos com algum saldo em aberto

      cQryAP += " AND SE2.D_E_L_E_T_ <> '*'"
      cQryAP += " GROUP BY SUBSTRING(SE2.E2_NATUREZ,1,6),SUBSTRING(SE2.E2_VENCREA,1,6)"
      cQryAP += " ORDER BY SUBSTRING(SE2.E2_NATUREZ,1,6)"

      
      //MemoWrite("C:\TEMP\PagamentosemAberto"+Alltrim(Str(nMes))+".SQL",cQryAP)
      
      TCQuery cQryAP NEW ALIAS "TTPG"

      TcSetField("TTPG","PAGABERTO","N",14,02)
      
      DbSelectArea("TTPG")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(TTPG->NATUREZ)
            RecLock("OXR",.F.)
               If Substr(TTPG->MESAREA,5,2) == "01"
                  OXR->AREA1 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "02"
                  OXR->AREA2 += TTPG->PAGABERTO               
               ElseIf Substr(TTPG->MESAREA,5,2) == "03"
                  OXR->AREA3 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "04"
                  OXR->AREA4 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "05"
                  OXR->AREA5 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "06"
                  OXR->AREA6 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "07"
                  OXR->AREA7 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "08"
                  OXR->AREA8 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "09"
                  OXR->AREA9 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "10"
                  OXR->AREA10 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "11"
                  OXR->AREA11 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "12"
                  OXR->AREA12 += TTPG->PAGABERTO
               EndIf
            MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := TTPG->NATUREZ
               If Substr(TTPG->MESAREA,5,2) == "01"
                  OXR->AREA1 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "02"
                  OXR->AREA2 := TTPG->PAGABERTO               
               ElseIf Substr(TTPG->MESAREA,5,2) == "03"
                  OXR->AREA3 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "04"
                  OXR->AREA4 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "05"
                  OXR->AREA5 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "06"
                  OXR->AREA6 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "07"
                  OXR->AREA7 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "08"
                  OXR->AREA8 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "09"
                  OXR->AREA9 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "10"
                  OXR->AREA10 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "11"
                  OXR->AREA11 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "12"
                  OXR->AREA12 := TTPG->PAGABERTO
               EndIf
            MsUnLock()
         EndIf
         
         DbSelectArea("TTPG")
         DbSkip()
         
      EndDo
      DbSelectArea("TTPG")
      DbCloseArea()      
      
   //Next nMes
*******************************************************************************************************************
*******************************************************************************************************************

   
   //Selecionando e Marcando Naturezas que cont?m Metas de Or?amentos
   cQryMTADM := "SELECT ED_CODIGO "
   cQryMTADM += " FROM "+RetSQLName("SED")
   cQryMTADM += " WHERE SUBSTRING(ED_CODIGO,1,6) IN "+FormatIn(AdmMetasOrc,",")
   cQryMTADM += " AND D_E_L_E_T_ <> '*'"
   
   TCQuery cQryMTADM New Alias "MTADM"
   
   DbSelectArea("MTADM")
   While !Eof()
      DbSelectArea("OXR")
      If DbSeek(SUBSTR(MTADM->ED_CODIGO,1,6))
         RecLock("OXR",.F.)
            OXR->MTORCADM := "S"
         MsUnLock()
      EndIf
      DbSelectArea("MTADM")
      DbSkip()
   EndDo
   DbSelectArea("MTADM")
   DbCloseArea()

   

   For nVez := 1 To 2
     
      If nVez == 1
         Cabec1       := "                                                                                                        G A S T O S   G E R A I S   D O   A D M I N I S T R A T I V O                                                       "
      Else
         Cabec1       := "                                                                                                   M E T A S   DE   O R ? A M E N T O   D O   A D M I N I S T R A T I V O                                                   "
      EndIf
      
      DbSelectArea("OXR")
      DbGoTop()
      While !EOF()

      
         If nVez == 2 .And. OXR->MTORCADM <> "S" //Na impress?o do segundo bloco de naturezas, s? apresenta natures
            DbSelectArea("OXR")                  //que comp?em Metas de Or?amentos.
            DbSkip()
            Loop
         EndIf
      
         //?????????????????????????????????????????????????????????????????????Ŀ
         //? Verifica o cancelamento pelo usuario...                             ?
         //???????????????????????????????????????????????????????????????????????
         
         If lAbortPrint
            @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
            Exit
         Endif

         //?????????????????????????????????????????????????????????????????????Ŀ
         //? Impressao do cabecalho do relatorio. . .                            ?
         //???????????????????????????????????????????????????????????????????????

         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
    
         //          10        20        30        40        50        60        70        80        90       100       110       120      130       140        150       160       170       180       190       200      210        220
         // 01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
         //"                                                                                                        G A S T O S   G E R A I S   D O   A D M I N I S T R A T I V O                                                       "
         //"                                                                                                   M E T A S   DE   O R ? A M E N T O   D O   A D M I N I S T R A T I V O                                                   "
         //"Natureza   Descri??o                             Janeiro      Fevereiro          Mar?o          Abril           Maio          Junho          Julho         Agosto       Setembro        Outubro       Novembro       Dezembro"
         // xxxxxxxxxX xxxxxxxxxXxxxxxxxxxXxxx Or?ado 999,999,999.99 999,999,999.99 999,999,999.99 999,999,999.99 999,999,999.99 999,999,999.99 999,999,999.99 999,999,999.99 999,999,999.99 999,999,999.99 999,999,999.99 999,999,999.99
         //                                      Real
         //                                   ARealiz
         //                                     Saldo
         //                                   %Utiliz

         @nLin,000 PSAY Substr(OXR->NATUREZ,1,6)
         @nLin,011 PSAY Substr(Posicione("SED",1,xFilial("SED")+Substr(OXR->NATUREZ,1,6),"ED_DESCRIC"),1,23)//30) 
         @nLin,035 PSAY "Or?ado"
         @nLin,042 PSAY OXR->MES1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->MES2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->MES3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->MES4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->MES5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->MES6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->MES7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->MES8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->MES9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->MES10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->MES11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->MES12 Picture "@E 999,999,999.99"
   
         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,037 PSAY "Real"
         @nLin,042 PSAY OXR->REA1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->REA2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->REA3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->REA4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->REA5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->REA6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->REA7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->REA8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->REA9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->REA10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->REA11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->REA12 Picture "@E 999,999,999.99"

         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,034 PSAY "ARealiz"
         @nLin,042 PSAY OXR->AREA1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->AREA2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->AREA3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->AREA4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->AREA5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->AREA6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->AREA7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->AREA8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->AREA9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->AREA10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->AREA11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->AREA12 Picture "@E 999,999,999.99"

         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,036 PSAY "Saldo"
         @nLin,042 PSAY (OXR->MES1 - (OXR->REA1+OXR->AREA1))  Picture "@E 999,999,999.99"
         @nLin,057 PSAY (OXR->MES2 - (OXR->REA2+OXR->AREA2))  Picture "@E 999,999,999.99"
         @nLin,072 PSAY (OXR->MES3 - (OXR->REA3+OXR->AREA3))  Picture "@E 999,999,999.99"
         @nLin,087 PSAY (OXR->MES4 - (OXR->REA4+OXR->AREA4))  Picture "@E 999,999,999.99"
         @nLin,102 PSAY (OXR->MES5 - (OXR->REA5+OXR->AREA5))  Picture "@E 999,999,999.99"
         @nLin,117 PSAY (OXR->MES6 - (OXR->REA6+OXR->AREA6))  Picture "@E 999,999,999.99"
         @nLin,132 PSAY (OXR->MES7 - (OXR->REA7+OXR->AREA7))  Picture "@E 999,999,999.99"
         @nLin,147 PSAY (OXR->MES8 - (OXR->REA8+OXR->AREA8))  Picture "@E 999,999,999.99"
         @nLin,162 PSAY (OXR->MES9 - (OXR->REA9+OXR->AREA9))  Picture "@E 999,999,999.99"
         @nLin,177 PSAY (OXR->MES10 - (OXR->REA10+OXR->AREA10)) Picture "@E 999,999,999.99"
         @nLin,192 PSAY (OXR->MES11 - (OXR->REA11+OXR->AREA11)) Picture "@E 999,999,999.99"
         @nLin,207 PSAY (OXR->MES12 - (OXR->REA12+OXR->AREA12)) Picture "@E 999,999,999.99"

         nLin := nLin + 1 // Avanca a linha de impressao
         If  nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,034 PSAY "%Utiliz"
         @nLin,042 PSAY IIf(OXR->MES1  > 0, ((OXR->REA1 /OXR->MES1) * 100),0) Picture "@E 999,999,999.99"
         @nLin,057 PSAY IIf(OXR->MES2  > 0, ((OXR->REA2 /OXR->MES2) * 100),0) Picture "@E 999,999,999.99"
         @nLin,072 PSAY IIf(OXR->MES3  > 0, ((OXR->REA3 /OXR->MES3) * 100),0) Picture "@E 999,999,999.99"
         @nLin,087 PSAY IIf(OXR->MES4  > 0, ((OXR->REA4 /OXR->MES4) * 100),0) Picture "@E 999,999,999.99"
         @nLin,102 PSAY IIf(OXR->MES5  > 0, ((OXR->REA5 /OXR->MES5) * 100),0) Picture "@E 999,999,999.99"
         @nLin,117 PSAY IIf(OXR->MES6  > 0, ((OXR->REA6 /OXR->MES6) * 100),0) Picture "@E 999,999,999.99"
         @nLin,132 PSAY IIf(OXR->MES7  > 0, ((OXR->REA7 /OXR->MES7) * 100),0) Picture "@E 999,999,999.99"
         @nLin,147 PSAY IIf(OXR->MES8  > 0, ((OXR->REA8 /OXR->MES8) * 100),0) Picture "@E 999,999,999.99"
         @nLin,162 PSAY IIf(OXR->MES9  > 0, ((OXR->REA9 /OXR->MES9) * 100),0) Picture "@E 999,999,999.99"
         @nLin,177 PSAY IIf(OXR->MES10 > 0, ((OXR->REA10/OXR->MES10)* 100),0) Picture "@E 999,999,999.99"
         @nLin,192 PSAY IIf(OXR->MES11 > 0, ((OXR->REA11/OXR->MES11)* 100),0) Picture "@E 999,999,999.99"
         @nLin,207 PSAY IIf(OXR->MES12 > 0, ((OXR->REA12/OXR->MES12)* 100),0) Picture "@E 999,999,999.99"

         //Acumula Totalizadores dos Or?ados de Gastos Gerais Adm
         nOrc1GADM += OXR->MES1
         nOrc2GADM += OXR->MES2
         nOrc3GADM += OXR->MES3
         nOrc4GADM += OXR->MES4
         nOrc5GADM += OXR->MES5
         nOrc6GADM += OXR->MES6
         nOrc7GADM += OXR->MES7
         nOrc8GADM += OXR->MES8
         nOrc9GADM += OXR->MES9
         nOrc10GADM += OXR->MES10
         nOrc11GADM += OXR->MES11
         nOrc12GADM += OXR->MES12
   
         //Acumula Totalizadores dos Realizados de Gastos Gerais Adm
         nRea1GADM  += OXR->REA1
         nRea2GADM  += OXR->REA2
         nRea3GADM  += OXR->REA3
         nRea4GADM  += OXR->REA4
         nRea5GADM  += OXR->REA5
         nRea6GADM  += OXR->REA6
         nRea7GADM  += OXR->REA7
         nRea8GADM  += OXR->REA8
         nRea9GADM  += OXR->REA9
         nRea10GADM += OXR->REA10
         nRea11GADM += OXR->REA11
         nRea12GADM += OXR->REA12

         //Acumula Totalizadores dos A Realizar de Gastos Gerais Adm
         naRea1GADM  += OXR->AREA1
         naRea2GADM  += OXR->AREA2
         naRea3GADM  += OXR->AREA3
         naRea4GADM  += OXR->AREA4
         naRea5GADM  += OXR->AREA5
         naRea6GADM  += OXR->AREA6
         naRea7GADM  += OXR->AREA7
         naRea8GADM  += OXR->AREA8
         naRea9GADM  += OXR->AREA9
         naRea10GADM += OXR->AREA10
         naRea11GADM += OXR->AREA11
         naRea12GADM += OXR->AREA12


         nSld1GG     += (OXR->MES1 - (OXR->REA1+OXR->AREA1))
         nSld2GG     += (OXR->MES2 - (OXR->REA2+OXR->AREA2))
         nSld3GG     += (OXR->MES3 - (OXR->REA3+OXR->AREA3))
         nSld4GG     += (OXR->MES4 - (OXR->REA4+OXR->AREA4))
         nSld5GG     += (OXR->MES5 - (OXR->REA5+OXR->AREA5))
         nSld6GG     += (OXR->MES6 - (OXR->REA6+OXR->AREA6))
         nSld7GG     += (OXR->MES7 - (OXR->REA7+OXR->AREA7))
         nSld8GG     += (OXR->MES8 - (OXR->REA8+OXR->AREA8))
         nSld9GG     += (OXR->MES9 - (OXR->REA9+OXR->AREA9))
         nSld10GG    += (OXR->MES10 - (OXR->REA10+OXR->AREA10))
         nSld11GG    += (OXR->MES11 - (OXR->REA11+OXR->AREA11))
         nSld12GG    += (OXR->MES12 - (OXR->REA12+OXR->AREA12))


         If nVez == 1
            //Acumula Totalizadores Global dos Or?ados de Gastos Gerais/Metas
            nGlbOr1GG += OXR->MES1
            nGlbOr2GG += OXR->MES2
            nGlbOr3GG += OXR->MES3
            nGlbOr4GG += OXR->MES4
            nGlbOr5GG += OXR->MES5
            nGlbOr6GG += OXR->MES6
            nGlbOr7GG += OXR->MES7
            nGlbOr8GG += OXR->MES8
            nGlbOr9GG += OXR->MES9
            nGlbOr10GG += OXR->MES10
            nGlbOr11GG += OXR->MES11
            nGlbOr12GG += OXR->MES12

           //Acumula Totalizadores Global dos Realizados de Gastos Gerais/Metas
           nGbRea1GG += OXR->REA1
           nGbRea2GG += OXR->REA2
           nGbRea3GG += OXR->REA3
           nGbRea4GG += OXR->REA4
           nGbRea5GG += OXR->REA5
           nGbRea6GG += OXR->REA6
           nGbRea7GG += OXR->REA7
           nGbRea8GG += OXR->REA8
           nGbRea9GG += OXR->REA9
           nGbRea10GG += OXR->REA10
           nGbRea11GG += OXR->REA11
           nGbRea12GG += OXR->REA12

           //Acumula Totalizadores Global dos A Realizar de Gastos Gerais/Metas
           nGbaRea1GG += OXR->AREA1
           nGbaRea2GG += OXR->AREA2
           nGbaRea3GG += OXR->AREA3
           nGbaRea4GG += OXR->AREA4
           nGbaRea5GG += OXR->AREA5
           nGbaRea6GG += OXR->AREA6
           nGbaRea7GG += OXR->AREA7
           nGbaRea8GG += OXR->AREA8
           nGbaRea9GG += OXR->AREA9
           nGbaRea10GG += OXR->AREA10
           nGbaRea11GG += OXR->AREA11
           nGbaRea12GG += OXR->AREA12

           //Acumula Totalizadores dos Saldos de Gastos Gerais/Metas
           nGbSld1GG     += (OXR->MES1 - (OXR->REA1+OXR->AREA1))
           nGbSld2GG     += (OXR->MES2 - (OXR->REA2+OXR->AREA2))
           nGbSld3GG     += (OXR->MES3 - (OXR->REA3+OXR->AREA3))
           nGbSld4GG     += (OXR->MES4 - (OXR->REA4+OXR->AREA4))
           nGbSld5GG     += (OXR->MES5 - (OXR->REA5+OXR->AREA5))
           nGbSld6GG     += (OXR->MES6 - (OXR->REA6+OXR->AREA6))
           nGbSld7GG     += (OXR->MES7 - (OXR->REA7+OXR->AREA7))
           nGbSld8GG     += (OXR->MES8 - (OXR->REA8+OXR->AREA8))
           nGbSld9GG     += (OXR->MES9 - (OXR->REA9+OXR->AREA9))
           nGbSld10GG    += (OXR->MES10 - (OXR->REA10+OXR->AREA10))
           nGbSld11GG    += (OXR->MES11 - (OXR->REA11+OXR->AREA11))
           nGbSld12GG    += (OXR->MES12 - (OXR->REA12+OXR->AREA12))

         EndIf
         
         nLin := nLin + 2 // Avanca a linha de impressao

         DbSkip() // Avanca o ponteiro do registro no arquivo
   
      EndDo

      /*************************************************
      * Imprime Totalizadores do Gastos Gerias Adm/Fin (Or?ado / Realizado)
      **************************************************/

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Or?ado de Gastos Gerais Adm/Fin "
      Else
        @nLin,000 PSAY "Tot.Metas Or?ado Gastos Ger.Adm/Fin "
      EndIf
      @nLin,042 PSAY nOrc1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nOrc2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nOrc3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nOrc4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nOrc5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nOrc6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nOrc7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nOrc8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nOrc9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nOrc10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY nOrc11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY nOrc12GADM Picture "@E 999,999,999.99"

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Realizado de Gastos Gerais Adm/Fin "
      Else
         @nLin,000 PSAY "Tot.Metas Realiz.Gastos Ger.Adm/Fin "
      EndIf
   
      @nLin,042 PSAY nRea1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nRea2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nRea3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nRea4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nRea5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nRea6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nRea7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nRea8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nRea9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nRea10GADM Picture "@E 999,999,999.99"
      @nLin,195 PSAY nRea11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY nRea12GADM Picture "@E 999,999,999.99"

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt A Realiz.Gastos Gerais Adm/Fin "
      Else
         @nLin,000 PSAY "Tot.Metas A Realiz.Gastos Ger.Adm/Fin. "
      EndIf
   
      @nLin,042 PSAY naRea1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY naRea2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY naRea3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY naRea4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY naRea5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY naRea6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY naRea7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY naRea8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY naRea9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY naRea10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY naRea11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY naRea12GADM Picture "@E 999,999,999.99"


      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Saldo de Gastos Gerais Adm/Fin "
      Else
         @nLin,000 PSAY "Tot.Metas Saldo Gastos Ger.Adm/Fin. "
      EndIf
   
      @nLin,042 PSAY nSld1GG  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nSld2GG  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nSld3GG  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nSld4GG  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nSld5GG  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nSld6GG  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nSld7GG  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nSld8GG  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nSld9GG  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nSld10GG Picture "@E 999,999,999.99"
      @nLin,192 PSAY nSld11GG Picture "@E 999,999,999.99"
      @nLin,207 PSAY nSld12GG Picture "@E 999,999,999.99"
      
      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt % Utiliz.Gastos Gerais Adm/Fin"
      Else
         @nLin,000 PSAY "Tot.Metas % Utiliz.Gastos Ger.Adm/Fin."
      EndIf
   
      @nLin,042 PSAY IIf(nOrc1GADM  > 0, ((nRea1GADM /nOrc1GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,057 PSAY IIf(nOrc2GADM  > 0, ((nRea2GADM /nOrc2GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,072 PSAY IIf(nOrc3GADM  > 0, ((nRea3GADM /nOrc3GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,087 PSAY IIf(nOrc4GADM  > 0, ((nRea4GADM /nOrc4GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,102 PSAY IIf(nOrc5GADM  > 0, ((nRea5GADM /nOrc5GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,117 PSAY IIf(nOrc6GADM  > 0, ((nRea6GADM /nOrc6GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,132 PSAY IIf(nOrc7GADM  > 0, ((nRea7GADM /nOrc7GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,147 PSAY IIf(nOrc8GADM  > 0, ((nRea8GADM /nOrc8GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,162 PSAY IIf(nOrc9GADM  > 0, ((nRea9GADM /nOrc9GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,177 PSAY IIf(nOrc10GADM  > 0, ((nRea10GADM /nOrc10GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,192 PSAY IIf(nOrc11GADM  > 0, ((nRea11GADM /nOrc11GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,207 PSAY IIf(nOrc12GADM  > 0, ((nRea12GADM /nOrc12GADM) *100),0) Picture "@E 999,999,999.99"
   
      nLin := 56 //For?a mudan?a de p?gina
      
      //Zera Acumuladores para apresentar valores das naturezas apenas que comp?em metas de or?amentos

      //Or?ado Gastos Gerais Adm
      nOrc1GADM := 0
      nOrc2GADM := 0
      nOrc3GADM := 0
      nOrc4GADM := 0
      nOrc5GADM := 0
      nOrc6GADM := 0
      nOrc7GADM := 0
      nOrc8GADM := 0
      nOrc9GADM := 0
      nOrc10GADM := 0
      nOrc11GADM := 0
      nOrc12GADM := 0

      //Realizado Gastos Gerais Adm
      nRea1GADM := 0
      nRea2GADM := 0
      nRea3GADM := 0
      nRea4GADM := 0
      nRea5GADM := 0
      nRea6GADM := 0
      nRea7GADM := 0
      nRea8GADM := 0
      nRea9GADM := 0
      nRea10GADM := 0
      nRea11GADM := 0
      nRea12GADM := 0

       //A Realizar Gastos Gerais Adm
      naRea1GADM := 0
      naRea2GADM := 0
      naRea3GADM := 0
      naRea4GADM := 0
      naRea5GADM := 0
      naRea6GADM := 0
      naRea7GADM := 0
      naRea8GADM := 0
      naRea9GADM := 0
      naRea10GADM := 0
      naRea11GADM := 0
      naRea12GADM := 0

 
      //Saldo(Or?ado - (Rrealizado + A Realizar))
      nSld1GG     := 0
      nSld2GG     := 0
      nSld3GG     := 0
      nSld4GG     := 0
      nSld5GG     := 0
      nSld6GG     := 0
      nSld7GG     := 0
      nSld8GG     := 0
      nSld9GG     := 0
      nSld10GG    := 0
      nSld11GG    := 0
      nSld12GG    := 0

   Next nVez

EndIf

If lComercio


   //Cria??o do Arquivo Tempor?rio
   If SELECT("OXR") > 0
      DbSelectArea("OXR")
      DbCloseArea()
   EndIf

   cArqInd := "ORCXREA.CDX"
   While File(cArqInd)
      DELETE FILE &cArqInd
   EndDo

   aCampos := {}
   //Campos para guardar Metas Or?adas para cada natureza
   Aadd(aCampos, {"NATUREZ","C",6,0})
   Aadd(aCampos, {"DESC","C",30,0})
   Aadd(aCampos, {"MES1","N",14,02})
   Aadd(aCampos, {"MES2","N",14,02})
   Aadd(aCampos, {"MES3","N",14,02})
   Aadd(aCampos, {"MES4","N",14,02})
   Aadd(aCampos, {"MES5","N",14,02})
   Aadd(aCampos, {"MES6","N",14,02})
   Aadd(aCampos, {"MES7","N",14,02})
   Aadd(aCampos, {"MES8","N",14,02})
   Aadd(aCampos, {"MES9","N",14,02})
   Aadd(aCampos, {"MES10","N",14,02})
   Aadd(aCampos, {"MES11","N",14,02})
   Aadd(aCampos, {"MES12","N",14,02})
   //Campos para guardar os Gastos Realizados para cada natureza
   Aadd(aCampos, {"REA1","N",14,02})
   Aadd(aCampos, {"REA2","N",14,02})
   Aadd(aCampos, {"REA3","N",14,02})
   Aadd(aCampos, {"REA4","N",14,02})
   Aadd(aCampos, {"REA5","N",14,02})
   Aadd(aCampos, {"REA6","N",14,02})
   Aadd(aCampos, {"REA7","N",14,02})
   Aadd(aCampos, {"REA8","N",14,02})
   Aadd(aCampos, {"REA9","N",14,02})
   Aadd(aCampos, {"REA10","N",14,02})
   Aadd(aCampos, {"REA11","N",14,02})
   Aadd(aCampos, {"REA12","N",14,02})
   //Campos para guardar os Gastos A Realizar para cada natureza
   Aadd(aCampos, {"AREA1","N",14,02})
   Aadd(aCampos, {"AREA2","N",14,02})
   Aadd(aCampos, {"AREA3","N",14,02})
   Aadd(aCampos, {"AREA4","N",14,02})
   Aadd(aCampos, {"AREA5","N",14,02})
   Aadd(aCampos, {"AREA6","N",14,02})
   Aadd(aCampos, {"AREA7","N",14,02})
   Aadd(aCampos, {"AREA8","N",14,02})
   Aadd(aCampos, {"AREA9","N",14,02})
   Aadd(aCampos, {"AREA10","N",14,02})
   Aadd(aCampos, {"AREA11","N",14,02})
   Aadd(aCampos, {"AREA12","N",14,02})
   //Campo para identificar que a natureza do Adm tem Meta de Or?amento
   Aadd(aCampos, {"MTORCADM","C",01,00})


   DbCreate("ORCXREA.DBF",aCampos)

   DbUseArea(.T.,"DBFCDX","ORCXREA.DBF","OXR")

   INDEX ON OXR->NATUREZ TO ORCXREA

   /************************
   *
   * GASTOS GERAIS COMERCIO (OR?ADO)
   *
   ************************/
   
      cQuery := " SELECT SED.ED_CODIGO,SE7.E7_ANO,SE7.E7_VALJAN1,SE7.E7_VALFEV1,SE7.E7_VALMAR1,SE7.E7_VALABR1," 
      cQuery += " SE7.E7_VALMAI1,SE7.E7_VALJUN1,SE7.E7_VALJUL1,SE7.E7_VALAGO1,SE7.E7_VALSET1,SE7.E7_VALOUT1,"
      cQuery += " SE7.E7_VALNOV1,SE7.E7_VALDEZ1"
      cQuery += " FROM "+RetSQLName("SED")+" SED LEFT OUTER JOIN "+RetSqlName("SE7")+" SE7 "
      cQuery += "                                ON SED.ED_CODIGO = SE7.E7_NATUREZ"
      cQuery += "                               AND SE7.E7_ANO = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      cQuery += "                               AND SE7.E7_FILIAL = '"+xFilial("SE7")+"'"
      cQuery += "                               AND SE7.D_E_L_E_T_ = ' '"
      cQuery += " WHERE SED.D_E_L_E_T_ = ' '"
      cQuery += "   AND SUBSTRING(SED.ED_CODIGO,1,6) IN "+FormatIn(cGGNatComl,",")      
      
      MemoWrite("C:\TEMP\ROrcXRea_Or?aGerADM.SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "ORGAD"
      
      TcSetField("ORGAD","E7_VALJAN1","N",14,02)
      TcSetField("ORGAD","E7_VALFEV1","N",14,02)
      TcSetField("ORGAD","E7_VALMAR1","N",14,02)
      TcSetField("ORGAD","E7_VALABR1","N",14,02)
      TcSetField("ORGAD","E7_VALMAI1","N",14,02)
      TcSetField("ORGAD","E7_VALJUN1","N",14,02)
      TcSetField("ORGAD","E7_VALJUL1","N",14,02)
      TcSetField("ORGAD","E7_VALAGO1","N",14,02)
      TcSetField("ORGAD","E7_VALSET1","N",14,02)
      TcSetField("ORGAD","E7_VALOUT1","N",14,02)
      TcSetField("ORGAD","E7_VALNOV1","N",14,02)
      TcSetField("ORGAD","E7_VALDEZ1","N",14,02)
      
      DbSelectArea("ORGAD")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(ORGAD->ED_CODIGO)
            RecLock("OXR",.F.)
               OXR->MES1 += ORGAD->E7_VALJAN1
               OXR->MES2 += ORGAD->E7_VALFEV1               
               OXR->MES3 += ORGAD->E7_VALMAR1
               OXR->MES4 += ORGAD->E7_VALABR1
               OXR->MES5 += ORGAD->E7_VALMAI1
               OXR->MES6 += ORGAD->E7_VALJUN1
               OXR->MES7 += ORGAD->E7_VALJUL1
               OXR->MES8 += ORGAD->E7_VALAGO1
               OXR->MES9 += ORGAD->E7_VALSET1
               OXR->MES10 += ORGAD->E7_VALOUT1
               OXR->MES11 += ORGAD->E7_VALNOV1
               OXR->MES12 += ORGAD->E7_VALDEZ1
              MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := ORGAD->ED_CODIGO
               OXR->MES1 := ORGAD->E7_VALJAN1
               OXR->MES2 := ORGAD->E7_VALFEV1               
               OXR->MES3 := ORGAD->E7_VALMAR1
               OXR->MES4 := ORGAD->E7_VALABR1
               OXR->MES5 := ORGAD->E7_VALMAI1
               OXR->MES6 := ORGAD->E7_VALJUN1
               OXR->MES7 := ORGAD->E7_VALJUL1
               OXR->MES8 := ORGAD->E7_VALAGO1
               OXR->MES9 := ORGAD->E7_VALSET1
               OXR->MES10 := ORGAD->E7_VALOUT1
               OXR->MES11 := ORGAD->E7_VALNOV1
               OXR->MES12 := ORGAD->E7_VALDEZ1
              MsUnLock()
         EndIf
         
         DbSelectArea("ORGAD")
         DbSkip()
         
      EndDo
      DbSelectArea("ORGAD")
      DbCloseArea()


   /************************
   *
   * GASTOS GERAIS COMERCIO (REALIZADO)
   *
   ************************/
   
      cQuery := " SELECT SUBSTRING(SE5.E5_DATA,1,6) AS MESGST,SUBSTRING(SE5.E5_NATUREZ,1,6) AS NATUREZ,SUM(SE5.E5_VALOR) AS VLGGADM "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,4) = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      
      cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cGGNatComl,",")

      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      
      //Implementado em 08/07/2008 - Solicita??o: Patricia
      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCF','PA ','NDF')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"
      cQuery += "  AND SE5.E5_MOTBX <> 'DEV'"      
      
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
      cQuery += " GROUP BY SUBSTRING(SE5.E5_DATA,1,6),SUBSTRING(SE5.E5_NATUREZ,1,6)"
      cQuery += " ORDER BY SUBSTRING(SE5.E5_NATUREZ,1,6)"
      
      MemoWrite("C:\TEMP\ROrcXRea_GastGerADM.SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "GADM"
      
      TcSetField("GADM","VLGGADM","N",14,02)
      
      DbSelectArea("GADM")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(GADM->NATUREZ)
            RecLock("OXR",.F.)
               If Substr(GADM->MESGST,5,2) == "01"
                  OXR->REA1 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "02"
                  OXR->REA2 += GADM->VLGGADM               
               ElseIf Substr(GADM->MESGST,5,2) == "03"
                  OXR->REA3 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "04"
                  OXR->REA4 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "05"
                  OXR->REA5 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "06"
                  OXR->REA6 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "07"
                  OXR->REA7 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "08"
                  OXR->REA8 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "09"
                  OXR->REA9 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "10"
                  OXR->REA10 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "11"
                  OXR->REA11 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "12"
                  OXR->REA12 += GADM->VLGGADM
               EndIf
            MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := GADM->NATUREZ
               If Substr(GADM->MESGST,5,2) == "01"
                  OXR->REA1 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "02"
                  OXR->REA2 := GADM->VLGGADM               
               ElseIf Substr(GADM->MESGST,5,2) == "03"
                  OXR->REA3 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "04"
                  OXR->REA4 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "05"
                  OXR->REA5 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "06"
                  OXR->REA6 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "07"
                  OXR->REA7 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "08"
                  OXR->REA8 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "09"
                  OXR->REA9 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "10"
                  OXR->REA10 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "11"
                  OXR->REA11 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "12"
                  OXR->REA12 := GADM->VLGGADM
               EndIf
            MsUnLock()
         EndIf
         
         DbSelectArea("GADM")
         DbSkip()
         
      EndDo
      DbSelectArea("GADM")
      DbCloseArea()
      

*******************************************************************************************************************
*******************************************************************************************************************

   /************************
   *
   * GASTOS GERAIS COMERCIO (A REALIZAR)
   *
   ************************/

   //For nMes := 1 To 3
      
   //   IncProc("Calc. Pagamentos em Aberto, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryAP := "SELECT SUBSTRING(SE2.E2_NATUREZ,1,6) NATUREZ,SUBSTRING(SE2.E2_VENCREA,1,6) MESAREA,SUM(SE2.E2_SALDO) AS PAGABERTO"
      cQryAP += " FROM "+RetSQLName("SE2")+" SE2"
      cQryAP += " WHERE SE2.E2_FILIAL = '"+xFilial("SE2")+"'"

      cQryAP += "  AND SE2.E2_TIPO NOT IN ('NCF','PA ','NDF')"
      cQryAP += "  AND SUBSTRING(SE2.E2_TIPO,3,1) <> '-'"

      cQryAP += "  AND  SUBSTRING(SE2.E2_VENCREA,1,4) = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      cQryAP += "  AND SUBSTRING(SE2.E2_NATUREZ,1,6) IN "+FormatIn(cGGNatComl,",")      
      
      cQryAP += "  AND SE2.E2_SALDO > 0 " //Titulos com algum saldo em aberto

      cQryAP += " AND SE2.D_E_L_E_T_ <> '*'"
      cQryAP += " GROUP BY SUBSTRING(SE2.E2_NATUREZ,1,6),SUBSTRING(SE2.E2_VENCREA,1,6)"
      cQryAP += " ORDER BY SUBSTRING(SE2.E2_NATUREZ,1,6)"

      
      //MemoWrite("C:\TEMP\PagamentosemAberto"+Alltrim(Str(nMes))+".SQL",cQryAP)
      
      TCQuery cQryAP NEW ALIAS "TTPG"

      TcSetField("TTPG","PAGABERTO","N",14,02)
      
      DbSelectArea("TTPG")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(TTPG->NATUREZ)
            RecLock("OXR",.F.)
               If Substr(TTPG->MESAREA,5,2) == "01"
                  OXR->AREA1 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "02"
                  OXR->AREA2 += TTPG->PAGABERTO               
               ElseIf Substr(TTPG->MESAREA,5,2) == "03"
                  OXR->AREA3 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "04"
                  OXR->AREA4 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "05"
                  OXR->AREA5 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "06"
                  OXR->AREA6 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "07"
                  OXR->AREA7 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "08"
                  OXR->AREA8 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "09"
                  OXR->AREA9 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "10"
                  OXR->AREA10 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "11"
                  OXR->AREA11 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "12"
                  OXR->AREA12 += TTPG->PAGABERTO
               EndIf
            MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := TTPG->NATUREZ
               If Substr(TTPG->MESAREA,5,2) == "01"
                  OXR->AREA1 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "02"
                  OXR->AREA2 := TTPG->PAGABERTO               
               ElseIf Substr(TTPG->MESAREA,5,2) == "03"
                  OXR->AREA3 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "04"
                  OXR->AREA4 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "05"
                  OXR->AREA5 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "06"
                  OXR->AREA6 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "07"
                  OXR->AREA7 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "08"
                  OXR->AREA8 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "09"
                  OXR->AREA9 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "10"
                  OXR->AREA10 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "11"
                  OXR->AREA11 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "12"
                  OXR->AREA12 := TTPG->PAGABERTO
               EndIf
            MsUnLock()
         EndIf
         
         DbSelectArea("TTPG")
         DbSkip()
         
      EndDo
      DbSelectArea("TTPG")
      DbCloseArea()      
      
   //Next nMes
*******************************************************************************************************************
*******************************************************************************************************************
   
   //Selecionando e Marcando Naturezas que cont?m Metas de Or?amentos
   cQryMTADM := "SELECT ED_CODIGO "
   cQryMTADM += " FROM "+RetSQLName("SED")
   cQryMTADM += " WHERE SUBSTRING(ED_CODIGO,1,6) IN "+FormatIn(cNtGMtOrCml,",")
   cQryMTADM += " AND D_E_L_E_T_ <> '*'"
   
   TCQuery cQryMTADM New Alias "MTADM"
   
   DbSelectArea("MTADM")
   While !Eof()
      DbSelectArea("OXR")
      If DbSeek(SUBSTR(MTADM->ED_CODIGO,1,6))
         RecLock("OXR",.F.)
            OXR->MTORCADM := "S"
         MsUnLock()
      EndIf
      DbSelectArea("MTADM")
      DbSkip()
   EndDo
   DbSelectArea("MTADM")
   DbCloseArea()

   

   For nVez := 1 To 2
     
      If nVez == 1
         Cabec1       := "                                                                                                        G A S T O S   G E R A I S   D O   C O M E R C I O                                                                   "
      Else
         Cabec1       := "                                                                                                   M E T A S   DE   O R ? A M E N T O   D O   C O M E R C I O                                                               "
      EndIf
      
      DbSelectArea("OXR")
      DbGoTop()
      While !EOF()

      
         If nVez == 2 .And. OXR->MTORCADM <> "S" //Na impress?o do segundo bloco de naturezas, s? apresenta natures
            DbSelectArea("OXR")                  //que comp?em Metas de Or?amentos.
            DbSkip()
            Loop
         EndIf
      
         //?????????????????????????????????????????????????????????????????????Ŀ
         //? Verifica o cancelamento pelo usuario...                             ?
         //???????????????????????????????????????????????????????????????????????
         
         If lAbortPrint
            @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
            Exit
         Endif

         //?????????????????????????????????????????????????????????????????????Ŀ
         //? Impressao do cabecalho do relatorio. . .                            ?
         //???????????????????????????????????????????????????????????????????????

         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
    
         //          10        20        30        40        50        60        70        80        90       100       110       120      130       140        150       160       170       180       190       200      210        220
         // 01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
         //"                                                                                                        G A S T O S   G E R A I S   D O   C O M E R C I O                                                                   "
         //"                                                                                                   M E T A S   DE   O R ? A M E N T O   D O   C O M E R C I O                                                               "
         //"Natureza   Descri??o                                       Janeiro     Fevereiro         Mar?o         Abril          Maio         Junho         Julho        Agosto      Setembro       Outubro      Novembro      Dezembro"
         // xxxxxxxxxX xxxxxxxxxXxxxxxxxxxXxxxxxxxxxX      Or?ado 9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99
         //                                             Realizado
         //                                            A Realizar
         //                                                  Real
         //                                            % Utiliza??o

         @nLin,000 PSAY Substr(OXR->NATUREZ,1,6)
         @nLin,011 PSAY Substr(Posicione("SED",1,xFilial("SED")+Substr(OXR->NATUREZ,1,6),"ED_DESCRIC"),1,23)//30) 
         @nLin,035 PSAY "Or?ado"
         @nLin,042 PSAY OXR->MES1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->MES2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->MES3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->MES4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->MES5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->MES6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->MES7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->MES8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->MES9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->MES10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->MES11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->MES12 Picture "@E 999,999,999.99"
   
         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,037 PSAY "Real"
         @nLin,042 PSAY OXR->REA1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->REA2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->REA3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->REA4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->REA5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->REA6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->REA7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->REA8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->REA9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->REA10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->REA11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->REA12 Picture "@E 999,999,999.99"


         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,034 PSAY "ARealiz"
         @nLin,042 PSAY OXR->AREA1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->AREA2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->AREA3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->AREA4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->AREA5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->AREA6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->AREA7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->AREA8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->AREA9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->AREA10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->AREA11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->AREA12 Picture "@E 999,999,999.99"


         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,036 PSAY "Saldo"
         @nLin,042 PSAY (OXR->MES1 - (OXR->REA1+OXR->AREA1))  Picture "@E 999,999,999.99"
         @nLin,057 PSAY (OXR->MES2 - (OXR->REA2+OXR->AREA2))  Picture "@E 999,999,999.99"
         @nLin,072 PSAY (OXR->MES3 - (OXR->REA3+OXR->AREA3))  Picture "@E 999,999,999.99"
         @nLin,087 PSAY (OXR->MES4 - (OXR->REA4+OXR->AREA4))  Picture "@E 999,999,999.99"
         @nLin,102 PSAY (OXR->MES5 - (OXR->REA5+OXR->AREA5))  Picture "@E 999,999,999.99"
         @nLin,117 PSAY (OXR->MES6 - (OXR->REA6+OXR->AREA6))  Picture "@E 999,999,999.99"
         @nLin,132 PSAY (OXR->MES7 - (OXR->REA7+OXR->AREA7))  Picture "@E 999,999,999.99"
         @nLin,147 PSAY (OXR->MES8 - (OXR->REA8+OXR->AREA8))  Picture "@E 999,999,999.99"
         @nLin,162 PSAY (OXR->MES9 - (OXR->REA9+OXR->AREA9))  Picture "@E 999,999,999.99"
         @nLin,177 PSAY (OXR->MES10 - (OXR->REA10+OXR->AREA10)) Picture "@E 999,999,999.99"
         @nLin,192 PSAY (OXR->MES11 - (OXR->REA11+OXR->AREA11)) Picture "@E 999,999,999.99"
         @nLin,207 PSAY (OXR->MES12 - (OXR->REA12+OXR->AREA12)) Picture "@E 999,999,999.99"

         nLin := nLin + 1 // Avanca a linha de impressao
         If  nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,034 PSAY "%Utiliz"
         @nLin,042 PSAY IIf(OXR->MES1  > 0, ((OXR->REA1 /OXR->MES1) * 100),0) Picture "@E 999,999,999.99"
         @nLin,057 PSAY IIf(OXR->MES2  > 0, ((OXR->REA2 /OXR->MES2) * 100),0) Picture "@E 999,999,999.99"
         @nLin,072 PSAY IIf(OXR->MES3  > 0, ((OXR->REA3 /OXR->MES3) * 100),0) Picture "@E 999,999,999.99"
         @nLin,087 PSAY IIf(OXR->MES4  > 0, ((OXR->REA4 /OXR->MES4) * 100),0) Picture "@E 999,999,999.99"
         @nLin,102 PSAY IIf(OXR->MES5  > 0, ((OXR->REA5 /OXR->MES5) * 100),0) Picture "@E 999,999,999.99"
         @nLin,117 PSAY IIf(OXR->MES6  > 0, ((OXR->REA6 /OXR->MES6) * 100),0) Picture "@E 999,999,999.99"
         @nLin,132 PSAY IIf(OXR->MES7  > 0, ((OXR->REA7 /OXR->MES7) * 100),0) Picture "@E 999,999,999.99"
         @nLin,147 PSAY IIf(OXR->MES8  > 0, ((OXR->REA8 /OXR->MES8) * 100),0) Picture "@E 999,999,999.99"
         @nLin,162 PSAY IIf(OXR->MES9  > 0, ((OXR->REA9 /OXR->MES9) * 100),0) Picture "@E 999,999,999.99"
         @nLin,177 PSAY IIf(OXR->MES10 > 0, ((OXR->REA10/OXR->MES10)* 100),0) Picture "@E 999,999,999.99"
         @nLin,192 PSAY IIf(OXR->MES11 > 0, ((OXR->REA11/OXR->MES11)* 100),0) Picture "@E 999,999,999.99"
         @nLin,207 PSAY IIf(OXR->MES12 > 0, ((OXR->REA12/OXR->MES12)* 100),0) Picture "@E 999,999,999.99"

         //Acumula Totalizadores dos Or?ados de Gastos Gerais/Metas
         nOrc1GADM += OXR->MES1
         nOrc2GADM += OXR->MES2
         nOrc3GADM += OXR->MES3
         nOrc4GADM += OXR->MES4
         nOrc5GADM += OXR->MES5
         nOrc6GADM += OXR->MES6
         nOrc7GADM += OXR->MES7
         nOrc8GADM += OXR->MES8
         nOrc9GADM += OXR->MES9
         nOrc10GADM += OXR->MES10
         nOrc11GADM += OXR->MES11
         nOrc12GADM += OXR->MES12
   
         //Acumula Totalizadores dos Realizados de Gastos Gerais/Metas
         nRea1GADM  += OXR->REA1
         nRea2GADM  += OXR->REA2
         nRea3GADM  += OXR->REA3
         nRea4GADM  += OXR->REA4
         nRea5GADM  += OXR->REA5
         nRea6GADM  += OXR->REA6
         nRea7GADM  += OXR->REA7
         nRea8GADM  += OXR->REA8
         nRea9GADM  += OXR->REA9
         nRea10GADM += OXR->REA10
         nRea11GADM += OXR->REA11
         nRea12GADM += OXR->REA12

         //Acumula Totalizadores dos A Realizar de Gastos Gerais/Metas
         naRea1GADM  += OXR->AREA1
         naRea2GADM  += OXR->AREA2
         naRea3GADM  += OXR->AREA3
         naRea4GADM  += OXR->AREA4
         naRea5GADM  += OXR->AREA5
         naRea6GADM  += OXR->AREA6
         naRea7GADM  += OXR->AREA7
         naRea8GADM  += OXR->AREA8
         naRea9GADM  += OXR->AREA9
         naRea10GADM += OXR->AREA10
         naRea11GADM += OXR->AREA11
         naRea12GADM += OXR->AREA12

         nSld1GG     += (OXR->MES1 - (OXR->REA1+OXR->AREA1))
         nSld2GG     += (OXR->MES2 - (OXR->REA2+OXR->AREA2))
         nSld3GG     += (OXR->MES3 - (OXR->REA3+OXR->AREA3))
         nSld4GG     += (OXR->MES4 - (OXR->REA4+OXR->AREA4))
         nSld5GG     += (OXR->MES5 - (OXR->REA5+OXR->AREA5))
         nSld6GG     += (OXR->MES6 - (OXR->REA6+OXR->AREA6))
         nSld7GG     += (OXR->MES7 - (OXR->REA7+OXR->AREA7))
         nSld8GG     += (OXR->MES8 - (OXR->REA8+OXR->AREA8))
         nSld9GG     += (OXR->MES9 - (OXR->REA9+OXR->AREA9))
         nSld10GG    += (OXR->MES10 - (OXR->REA10+OXR->AREA10))
         nSld11GG    += (OXR->MES11 - (OXR->REA11+OXR->AREA11))
         nSld12GG    += (OXR->MES12 - (OXR->REA12+OXR->AREA12))

         If nVez == 1
            //Acumula Totalizadores Global dos Or?ados de Gastos Gerais/Metas
            nGlbOr1GG += OXR->MES1
            nGlbOr2GG += OXR->MES2
            nGlbOr3GG += OXR->MES3
            nGlbOr4GG += OXR->MES4
            nGlbOr5GG += OXR->MES5
            nGlbOr6GG += OXR->MES6
            nGlbOr7GG += OXR->MES7
            nGlbOr8GG += OXR->MES8
            nGlbOr9GG += OXR->MES9
            nGlbOr10GG += OXR->MES10
            nGlbOr11GG += OXR->MES11
            nGlbOr12GG += OXR->MES12

           //Acumula Totalizadores Global dos Realizados de Gastos Gerais/Metas
           nGbRea1GG += OXR->REA1
           nGbRea2GG += OXR->REA2
           nGbRea3GG += OXR->REA3
           nGbRea4GG += OXR->REA4
           nGbRea5GG += OXR->REA5
           nGbRea6GG += OXR->REA6
           nGbRea7GG += OXR->REA7
           nGbRea8GG += OXR->REA8
           nGbRea9GG += OXR->REA9
           nGbRea10GG += OXR->REA10
           nGbRea11GG += OXR->REA11
           nGbRea12GG += OXR->REA12

           //Acumula Totalizadores Global dos A Realizar de Gastos Gerais/Metas
           nGbaRea1GG += OXR->AREA1
           nGbaRea2GG += OXR->AREA2
           nGbaRea3GG += OXR->AREA3
           nGbaRea4GG += OXR->AREA4
           nGbaRea5GG += OXR->AREA5
           nGbaRea6GG += OXR->AREA6
           nGbaRea7GG += OXR->AREA7
           nGbaRea8GG += OXR->AREA8
           nGbaRea9GG += OXR->AREA9
           nGbaRea10GG += OXR->AREA10
           nGbaRea11GG += OXR->AREA11
           nGbaRea12GG += OXR->AREA12

           //Acumula Totalizadores dos Saldos de Gastos Gerais/Metas
           nGbSld1GG     += (OXR->MES1 - (OXR->REA1+OXR->AREA1))
           nGbSld2GG     += (OXR->MES2 - (OXR->REA2+OXR->AREA2))
           nGbSld3GG     += (OXR->MES3 - (OXR->REA3+OXR->AREA3))
           nGbSld4GG     += (OXR->MES4 - (OXR->REA4+OXR->AREA4))
           nGbSld5GG     += (OXR->MES5 - (OXR->REA5+OXR->AREA5))
           nGbSld6GG     += (OXR->MES6 - (OXR->REA6+OXR->AREA6))
           nGbSld7GG     += (OXR->MES7 - (OXR->REA7+OXR->AREA7))
           nGbSld8GG     += (OXR->MES8 - (OXR->REA8+OXR->AREA8))
           nGbSld9GG     += (OXR->MES9 - (OXR->REA9+OXR->AREA9))
           nGbSld10GG    += (OXR->MES10 - (OXR->REA10+OXR->AREA10))
           nGbSld11GG    += (OXR->MES11 - (OXR->REA11+OXR->AREA11))
           nGbSld12GG    += (OXR->MES12 - (OXR->REA12+OXR->AREA12))
           
         EndIf


         nLin := nLin + 2 // Avanca a linha de impressao

         DbSkip() // Avanca o ponteiro do registro no arquivo
   
      EndDo

      /*************************************************
      * Imprime Totalizadores do Gastos Gerias Adm/Fin (Or?ado / Realizado)
      **************************************************/

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Or?ado de Gastos Gerais Com?rcio "
      Else
        @nLin,000 PSAY "Tot.Metas Or?ado Gastos Ger.Com.... "
      EndIf
      @nLin,042 PSAY nOrc1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nOrc2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nOrc3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nOrc4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nOrc5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nOrc6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nOrc7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nOrc8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nOrc9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nOrc10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY nOrc11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY nOrc12GADM Picture "@E 999,999,999.99"

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Realizado de Gastos Gerais Com?rcio "
      Else
         @nLin,000 PSAY "Tot.Metas Realiz.Gastos Ger.Com?rcio "
      EndIf
   
      @nLin,042 PSAY nRea1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nRea2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nRea3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nRea4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nRea5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nRea6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nRea7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nRea8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nRea9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nRea10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY nRea11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY nRea12GADM Picture "@E 999,999,999.99"


      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt A Realizar de Gastos Gerais Com?rcio "
      Else
         @nLin,000 PSAY "Tot.Metas A Realiz.Gastos Ger.Com?rcio "
      EndIf
   
      @nLin,042 PSAY naRea1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY naRea2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY naRea3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY naRea4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY naRea5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY naRea6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY naRea7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY naRea8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY naRea9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY naRea10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY naRea11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY naRea12GADM Picture "@E 999,999,999.99"
      

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Saldo de Gastos Gerais Adm/Fin "
      Else
         @nLin,000 PSAY "Tot.Metas Saldo Gastos Ger.Adm/Fin "
      EndIf
   
      @nLin,042 PSAY nSld1GG  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nSld2GG  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nSld3GG  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nSld4GG  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nSld5GG  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nSld6GG  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nSld7GG  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nSld8GG  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nSld9GG  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nSld10GG Picture "@E 999,999,999.99"
      @nLin,192 PSAY nSld11GG Picture "@E 999,999,999.99"
      @nLin,207 PSAY nSld12GG Picture "@E 999,999,999.99"
      
      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt % Utiliza??o de Gastos Gerais Com?rcio"
      Else
         @nLin,000 PSAY "Tot.Metas % Utiliz.Gastos Ger.Com?rcio"
      EndIf
   
      @nLin,042 PSAY IIf(nOrc1GADM  > 0, ((nRea1GADM /nOrc1GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,057 PSAY IIf(nOrc2GADM  > 0, ((nRea2GADM /nOrc2GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,072 PSAY IIf(nOrc3GADM  > 0, ((nRea3GADM /nOrc3GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,087 PSAY IIf(nOrc4GADM  > 0, ((nRea4GADM /nOrc4GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,102 PSAY IIf(nOrc5GADM  > 0, ((nRea5GADM /nOrc5GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,117 PSAY IIf(nOrc6GADM  > 0, ((nRea6GADM /nOrc6GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,132 PSAY IIf(nOrc7GADM  > 0, ((nRea7GADM /nOrc7GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,147 PSAY IIf(nOrc8GADM  > 0, ((nRea8GADM /nOrc8GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,162 PSAY IIf(nOrc9GADM  > 0, ((nRea9GADM /nOrc9GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,177 PSAY IIf(nOrc10GADM  > 0, ((nRea10GADM /nOrc10GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,192 PSAY IIf(nOrc11GADM  > 0, ((nRea11GADM /nOrc11GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,207 PSAY IIf(nOrc12GADM  > 0, ((nRea12GADM /nOrc12GADM) *100),0) Picture "@E 999,999,999.99"
   
      nLin := 56 //For?a mudan?a de p?gina
      
      //Zera Acumuladores para apresentar valores das naturezas apenas que comp?em metas de or?amentos

      //Or?ado Gastos Gerais Adm
      nOrc1GADM := 0
      nOrc2GADM := 0
      nOrc3GADM := 0
      nOrc4GADM := 0
      nOrc5GADM := 0
      nOrc6GADM := 0
      nOrc7GADM := 0
      nOrc8GADM := 0
      nOrc9GADM := 0
      nOrc10GADM := 0
      nOrc11GADM := 0
      nOrc12GADM := 0

      //Realizado Gastos Gerais Adm
      nRea1GADM := 0
      nRea2GADM := 0
      nRea3GADM := 0
      nRea4GADM := 0
      nRea5GADM := 0
      nRea6GADM := 0
      nRea7GADM := 0
      nRea8GADM := 0
      nRea9GADM := 0
      nRea10GADM := 0
      nRea11GADM := 0
      nRea12GADM := 0

      //A Realizar Gastos Gerais Adm
      naRea1GADM := 0
      naRea2GADM := 0
      naRea3GADM := 0
      naRea4GADM := 0
      naRea5GADM := 0
      naRea6GADM := 0
      naRea7GADM := 0
      naRea8GADM := 0
      naRea9GADM := 0
      naRea10GADM := 0
      naRea11GADM := 0
      naRea12GADM := 0
      
      //Saldo(Or?ado - (Rrealizado + A Realizar))
      nSld1GG     := 0
      nSld2GG     := 0
      nSld3GG     := 0
      nSld4GG     := 0
      nSld5GG     := 0
      nSld6GG     := 0
      nSld7GG     := 0
      nSld8GG     := 0
      nSld9GG     := 0
      nSld10GG    := 0
      nSld11GG    := 0
      nSld12GG    := 0
      
   Next nVez

EndIf

If lComExter



   //Cria??o do Arquivo Tempor?rio
   If SELECT("OXR") > 0
      DbSelectArea("OXR")
      DbCloseArea()
   EndIf

   cArqInd := "ORCXREA.CDX"
   While File(cArqInd)
      DELETE FILE &cArqInd
   EndDo

   aCampos := {}
   //Campos para guardar Metas Or?adas para cada natureza
   Aadd(aCampos, {"NATUREZ","C",6,0})
   Aadd(aCampos, {"DESC","C",30,0})
   Aadd(aCampos, {"MES1","N",14,02})
   Aadd(aCampos, {"MES2","N",14,02})
   Aadd(aCampos, {"MES3","N",14,02})
   Aadd(aCampos, {"MES4","N",14,02})
   Aadd(aCampos, {"MES5","N",14,02})
   Aadd(aCampos, {"MES6","N",14,02})
   Aadd(aCampos, {"MES7","N",14,02})
   Aadd(aCampos, {"MES8","N",14,02})
   Aadd(aCampos, {"MES9","N",14,02})
   Aadd(aCampos, {"MES10","N",14,02})
   Aadd(aCampos, {"MES11","N",14,02})
   Aadd(aCampos, {"MES12","N",14,02})
   //Campos para guardar os Gastos Realizados para cada natureza
   Aadd(aCampos, {"REA1","N",14,02})
   Aadd(aCampos, {"REA2","N",14,02})
   Aadd(aCampos, {"REA3","N",14,02})
   Aadd(aCampos, {"REA4","N",14,02})
   Aadd(aCampos, {"REA5","N",14,02})
   Aadd(aCampos, {"REA6","N",14,02})
   Aadd(aCampos, {"REA7","N",14,02})
   Aadd(aCampos, {"REA8","N",14,02})
   Aadd(aCampos, {"REA9","N",14,02})
   Aadd(aCampos, {"REA10","N",14,02})
   Aadd(aCampos, {"REA11","N",14,02})
   Aadd(aCampos, {"REA12","N",14,02})
   //Campos para guardar os Gastos A Realizar para cada natureza
   Aadd(aCampos, {"AREA1","N",14,02})
   Aadd(aCampos, {"AREA2","N",14,02})
   Aadd(aCampos, {"AREA3","N",14,02})
   Aadd(aCampos, {"AREA4","N",14,02})
   Aadd(aCampos, {"AREA5","N",14,02})
   Aadd(aCampos, {"AREA6","N",14,02})
   Aadd(aCampos, {"AREA7","N",14,02})
   Aadd(aCampos, {"AREA8","N",14,02})
   Aadd(aCampos, {"AREA9","N",14,02})
   Aadd(aCampos, {"AREA10","N",14,02})
   Aadd(aCampos, {"AREA11","N",14,02})
   Aadd(aCampos, {"AREA12","N",14,02})
   //Campo para identificar que a natureza do Adm tem Meta de Or?amento
   Aadd(aCampos, {"MTORCADM","C",01,00})


   DbCreate("ORCXREA.DBF",aCampos)

   DbUseArea(.T.,"DBFCDX","ORCXREA.DBF","OXR")

   INDEX ON OXR->NATUREZ TO ORCXREA

   /************************
   *
   * GASTOS GERAIS COMERCIO EXTERIOR (OR?ADO)
   *
   ************************/
   
      cQuery := " SELECT SED.ED_CODIGO,SE7.E7_ANO,SE7.E7_VALJAN1,SE7.E7_VALFEV1,SE7.E7_VALMAR1,SE7.E7_VALABR1," 
      cQuery += " SE7.E7_VALMAI1,SE7.E7_VALJUN1,SE7.E7_VALJUL1,SE7.E7_VALAGO1,SE7.E7_VALSET1,SE7.E7_VALOUT1,"
      cQuery += " SE7.E7_VALNOV1,SE7.E7_VALDEZ1"
      cQuery += " FROM "+RetSQLName("SED")+" SED LEFT OUTER JOIN "+RetSqlName("SE7")+" SE7 "
      cQuery += "                                ON SED.ED_CODIGO = SE7.E7_NATUREZ"
      cQuery += "                               AND SE7.E7_ANO = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      cQuery += "                               AND SE7.E7_FILIAL = '"+xFilial("SE7")+"'"
      cQuery += "                               AND SE7.D_E_L_E_T_ = ' '"
      cQuery += " WHERE SED.D_E_L_E_T_ = ' '"
      cQuery += "   AND SUBSTRING(SED.ED_CODIGO,1,6) IN "+FormatIn(cGGCENat,",")      
      
      MemoWrite("C:\TEMP\ROrcXRea_Or?aGerADM.SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "ORGAD"
      
      TcSetField("ORGAD","E7_VALJAN1","N",14,02)
      TcSetField("ORGAD","E7_VALFEV1","N",14,02)
      TcSetField("ORGAD","E7_VALMAR1","N",14,02)
      TcSetField("ORGAD","E7_VALABR1","N",14,02)
      TcSetField("ORGAD","E7_VALMAI1","N",14,02)
      TcSetField("ORGAD","E7_VALJUN1","N",14,02)
      TcSetField("ORGAD","E7_VALJUL1","N",14,02)
      TcSetField("ORGAD","E7_VALAGO1","N",14,02)
      TcSetField("ORGAD","E7_VALSET1","N",14,02)
      TcSetField("ORGAD","E7_VALOUT1","N",14,02)
      TcSetField("ORGAD","E7_VALNOV1","N",14,02)
      TcSetField("ORGAD","E7_VALDEZ1","N",14,02)
      
      DbSelectArea("ORGAD")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(ORGAD->ED_CODIGO)
            RecLock("OXR",.F.)
               OXR->MES1 += ORGAD->E7_VALJAN1
               OXR->MES2 += ORGAD->E7_VALFEV1               
               OXR->MES3 += ORGAD->E7_VALMAR1
               OXR->MES4 += ORGAD->E7_VALABR1
               OXR->MES5 += ORGAD->E7_VALMAI1
               OXR->MES6 += ORGAD->E7_VALJUN1
               OXR->MES7 += ORGAD->E7_VALJUL1
               OXR->MES8 += ORGAD->E7_VALAGO1
               OXR->MES9 += ORGAD->E7_VALSET1
               OXR->MES10 += ORGAD->E7_VALOUT1
               OXR->MES11 += ORGAD->E7_VALNOV1
               OXR->MES12 += ORGAD->E7_VALDEZ1
              MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := ORGAD->ED_CODIGO
               OXR->MES1 := ORGAD->E7_VALJAN1
               OXR->MES2 := ORGAD->E7_VALFEV1               
               OXR->MES3 := ORGAD->E7_VALMAR1
               OXR->MES4 := ORGAD->E7_VALABR1
               OXR->MES5 := ORGAD->E7_VALMAI1
               OXR->MES6 := ORGAD->E7_VALJUN1
               OXR->MES7 := ORGAD->E7_VALJUL1
               OXR->MES8 := ORGAD->E7_VALAGO1
               OXR->MES9 := ORGAD->E7_VALSET1
               OXR->MES10 := ORGAD->E7_VALOUT1
               OXR->MES11 := ORGAD->E7_VALNOV1
               OXR->MES12 := ORGAD->E7_VALDEZ1
              MsUnLock()
         EndIf
         
         DbSelectArea("ORGAD")
         DbSkip()
         
      EndDo
      DbSelectArea("ORGAD")
      DbCloseArea()


   /************************
   *
   * GASTOS GERAIS COMERCIO EXTERIOR (REALIZADO)
   *
   ************************/
   
      cQuery := " SELECT SUBSTRING(SE5.E5_DATA,1,6) AS MESGST,SUBSTRING(SE5.E5_NATUREZ,1,6) AS NATUREZ,SUM(SE5.E5_VALOR) AS VLGGADM "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,4) = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      
      cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cGGCENat,",")

      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      
      //Implementado em 08/07/2008 - Solicita??o: Patricia
      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCF','PA ','NDF')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"
      cQuery += "  AND SE5.E5_MOTBX <> 'DEV'"      
      
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
      cQuery += " GROUP BY SUBSTRING(SE5.E5_DATA,1,6),SUBSTRING(SE5.E5_NATUREZ,1,6)"
      cQuery += " ORDER BY SUBSTRING(SE5.E5_NATUREZ,1,6)"
      
      MemoWrite("C:\TEMP\ROrcXRea_GastGerADM.SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "GADM"
      
      TcSetField("GADM","VLGGADM","N",14,02)
      
      DbSelectArea("GADM")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(GADM->NATUREZ)
            RecLock("OXR",.F.)
               If Substr(GADM->MESGST,5,2) == "01"
                  OXR->REA1 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "02"
                  OXR->REA2 += GADM->VLGGADM               
               ElseIf Substr(GADM->MESGST,5,2) == "03"
                  OXR->REA3 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "04"
                  OXR->REA4 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "05"
                  OXR->REA5 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "06"
                  OXR->REA6 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "07"
                  OXR->REA7 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "08"
                  OXR->REA8 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "09"
                  OXR->REA9 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "10"
                  OXR->REA10 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "11"
                  OXR->REA11 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "12"
                  OXR->REA12 += GADM->VLGGADM
               EndIf
            MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := GADM->NATUREZ
               If Substr(GADM->MESGST,5,2) == "01"
                  OXR->REA1 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "02"
                  OXR->REA2 := GADM->VLGGADM               
               ElseIf Substr(GADM->MESGST,5,2) == "03"
                  OXR->REA3 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "04"
                  OXR->REA4 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "05"
                  OXR->REA5 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "06"
                  OXR->REA6 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "07"
                  OXR->REA7 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "08"
                  OXR->REA8 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "09"
                  OXR->REA9 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "10"
                  OXR->REA10 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "11"
                  OXR->REA11 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "12"
                  OXR->REA12 := GADM->VLGGADM
               EndIf
            MsUnLock()
         EndIf
         
         DbSelectArea("GADM")
         DbSkip()
         
      EndDo
      DbSelectArea("GADM")
      DbCloseArea()
      
   
*******************************************************************************************************************
*******************************************************************************************************************

   /************************
   *
   * GASTOS GERAIS COMERCIO EXTERIOR (A REALIZAR)
   *
   ************************/

   //For nMes := 1 To 3
      
   //   IncProc("Calc. Pagamentos em Aberto, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryAP := "SELECT SUBSTRING(SE2.E2_NATUREZ,1,6) NATUREZ,SUBSTRING(SE2.E2_VENCREA,1,6) MESAREA,SUM(SE2.E2_SALDO) AS PAGABERTO"
      cQryAP += " FROM "+RetSQLName("SE2")+" SE2"
      cQryAP += " WHERE SE2.E2_FILIAL = '"+xFilial("SE2")+"'"

      cQryAP += "  AND SE2.E2_TIPO NOT IN ('NCF','PA ','NDF')"
      cQryAP += "  AND SUBSTRING(SE2.E2_TIPO,3,1) <> '-'"

      cQryAP += "  AND  SUBSTRING(SE2.E2_VENCREA,1,4) = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      cQryAP += "  AND SUBSTRING(SE2.E2_NATUREZ,1,6) IN "+FormatIn(cGGCENat,",")      
      
      cQryAP += "  AND SE2.E2_SALDO > 0 " //Titulos com algum saldo em aberto

      cQryAP += " AND SE2.D_E_L_E_T_ <> '*'"
      cQryAP += " GROUP BY SUBSTRING(SE2.E2_NATUREZ,1,6),SUBSTRING(SE2.E2_VENCREA,1,6)"
      cQryAP += " ORDER BY SUBSTRING(SE2.E2_NATUREZ,1,6)"

      
      //MemoWrite("C:\TEMP\PagamentosemAberto"+Alltrim(Str(nMes))+".SQL",cQryAP)
      
      TCQuery cQryAP NEW ALIAS "TTPG"

      TcSetField("TTPG","PAGABERTO","N",14,02)
      
      DbSelectArea("TTPG")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(TTPG->NATUREZ)
            RecLock("OXR",.F.)
               If Substr(TTPG->MESAREA,5,2) == "01"
                  OXR->AREA1 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "02"
                  OXR->AREA2 += TTPG->PAGABERTO               
               ElseIf Substr(TTPG->MESAREA,5,2) == "03"
                  OXR->AREA3 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "04"
                  OXR->AREA4 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "05"
                  OXR->AREA5 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "06"
                  OXR->AREA6 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "07"
                  OXR->AREA7 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "08"
                  OXR->AREA8 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "09"
                  OXR->AREA9 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "10"
                  OXR->AREA10 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "11"
                  OXR->AREA11 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "12"
                  OXR->AREA12 += TTPG->PAGABERTO
               EndIf
            MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := TTPG->NATUREZ
               If Substr(TTPG->MESAREA,5,2) == "01"
                  OXR->AREA1 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "02"
                  OXR->AREA2 := TTPG->PAGABERTO               
               ElseIf Substr(TTPG->MESAREA,5,2) == "03"
                  OXR->AREA3 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "04"
                  OXR->AREA4 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "05"
                  OXR->AREA5 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "06"
                  OXR->AREA6 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "07"
                  OXR->AREA7 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "08"
                  OXR->AREA8 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "09"
                  OXR->AREA9 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "10"
                  OXR->AREA10 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "11"
                  OXR->AREA11 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "12"
                  OXR->AREA12 := TTPG->PAGABERTO
               EndIf
            MsUnLock()
         EndIf
         
         DbSelectArea("TTPG")
         DbSkip()
         
      EndDo
      DbSelectArea("TTPG")
      DbCloseArea()      
      
   //Next nMes
*******************************************************************************************************************
*******************************************************************************************************************

   //Selecionando e Marcando Naturezas que cont?m Metas de Or?amentos
   cQryMTADM := "SELECT ED_CODIGO "
   cQryMTADM += " FROM "+RetSQLName("SED")
   cQryMTADM += " WHERE SUBSTRING(ED_CODIGO,1,6) IN "+FormatIn(cMtCENat,",")
   cQryMTADM += " AND D_E_L_E_T_ <> '*'"
   
   TCQuery cQryMTADM New Alias "MTADM"
   
   DbSelectArea("MTADM")
   While !Eof()
      DbSelectArea("OXR")
      If DbSeek(SUBSTR(MTADM->ED_CODIGO,1,6))
         RecLock("OXR",.F.)
            OXR->MTORCADM := "S"
         MsUnLock()
      EndIf
      DbSelectArea("MTADM")
      DbSkip()
   EndDo
   DbSelectArea("MTADM")
   DbCloseArea()

   

   For nVez := 1 To 2
     
      If nVez == 1
         Cabec1       := "                                                                                                    G A S T O S   G E R A I S   D O   C O M E R C I O   E X T E R I O R                                                     "
      Else
         Cabec1       := "                                                                                                M E T A S   DE   O R ? A M E N T O   D O   C O M E R C I O  E X T E R I O R                                                 "
      EndIf
      
      DbSelectArea("OXR")
      DbGoTop()
      While !EOF()

      
         If nVez == 2 .And. OXR->MTORCADM <> "S" //Na impress?o do segundo bloco de naturezas, s? apresenta natures
            DbSelectArea("OXR")                  //que comp?em Metas de Or?amentos.
            DbSkip()
            Loop
         EndIf
      
         //?????????????????????????????????????????????????????????????????????Ŀ
         //? Verifica o cancelamento pelo usuario...                             ?
         //???????????????????????????????????????????????????????????????????????
         
         If lAbortPrint
            @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
            Exit
         Endif

         //?????????????????????????????????????????????????????????????????????Ŀ
         //? Impressao do cabecalho do relatorio. . .                            ?
         //???????????????????????????????????????????????????????????????????????

         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
    
         //          10        20        30        40        50        60        70        80        90       100       110       120      130       140        150       160       170       180       190       200      210        220
         // 01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
         //"                                                                                                    G A S T O S   G E R A I S   D O   C O M E R C I O   E X T E R I O R                                                     "
         //"                                                                                                M E T A S   DE   O R ? A M E N T O   D O   C O M E R C I O  E X T E R I O R                                                 "
         //"Natureza   Descri??o                                       Janeiro     Fevereiro         Mar?o         Abril          Maio         Junho         Julho        Agosto      Setembro       Outubro      Novembro      Dezembro"
         // xxxxxxxxxX xxxxxxxxxXxxxxxxxxxXxxxxxxxxxX      Or?ado 9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99
         //                                             Realizado
         //                                            A Realizar
         //                                                  Real
         //                                            % Utiliza??o

         @nLin,000 PSAY Substr(OXR->NATUREZ,1,6)
         @nLin,011 PSAY Substr(Posicione("SED",1,xFilial("SED")+Substr(OXR->NATUREZ,1,6),"ED_DESCRIC"),1,23)//30) 
         @nLin,035 PSAY "Or?ado"
         @nLin,042 PSAY OXR->MES1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->MES2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->MES3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->MES4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->MES5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->MES6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->MES7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->MES8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->MES9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->MES10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->MES11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->MES12 Picture "@E 999,999,999.99"
   
         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,037 PSAY "Real"
         @nLin,042 PSAY OXR->REA1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->REA2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->REA3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->REA4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->REA5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->REA6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->REA7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->REA8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->REA9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->REA10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->REA11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->REA12 Picture "@E 999,999,999.99"

         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,034 PSAY "ARealiz"
         @nLin,042 PSAY OXR->AREA1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->AREA2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->AREA3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->AREA4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->AREA5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->AREA6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->AREA7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->AREA8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->AREA9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->AREA10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->AREA11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->AREA12 Picture "@E 999,999,999.99"
         

         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,036 PSAY "Saldo"
         @nLin,042 PSAY (OXR->MES1 - (OXR->REA1+OXR->AREA1))  Picture "@E 999,999,999.99"
         @nLin,057 PSAY (OXR->MES2 - (OXR->REA2+OXR->AREA2))  Picture "@E 999,999,999.99"
         @nLin,072 PSAY (OXR->MES3 - (OXR->REA3+OXR->AREA3))  Picture "@E 999,999,999.99"
         @nLin,087 PSAY (OXR->MES4 - (OXR->REA4+OXR->AREA4))  Picture "@E 999,999,999.99"
         @nLin,102 PSAY (OXR->MES5 - (OXR->REA5+OXR->AREA5))  Picture "@E 999,999,999.99"
         @nLin,117 PSAY (OXR->MES6 - (OXR->REA6+OXR->AREA6))  Picture "@E 999,999,999.99"
         @nLin,132 PSAY (OXR->MES7 - (OXR->REA7+OXR->AREA7))  Picture "@E 999,999,999.99"
         @nLin,147 PSAY (OXR->MES8 - (OXR->REA8+OXR->AREA8))  Picture "@E 999,999,999.99"
         @nLin,162 PSAY (OXR->MES9 - (OXR->REA9+OXR->AREA9))  Picture "@E 999,999,999.99"
         @nLin,177 PSAY (OXR->MES10 - (OXR->REA10+OXR->AREA10)) Picture "@E 999,999,999.99"
         @nLin,192 PSAY (OXR->MES11 - (OXR->REA11+OXR->AREA11)) Picture "@E 999,999,999.99"
         @nLin,207 PSAY (OXR->MES12 - (OXR->REA12+OXR->AREA12)) Picture "@E 999,999,999.99"
         
         nLin := nLin + 1 // Avanca a linha de impressao
         If  nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,034 PSAY "%Utiliz"
         @nLin,042 PSAY IIf(OXR->MES1  > 0, ((OXR->REA1 /OXR->MES1) * 100),0) Picture "@E 999,999,999.99"
         @nLin,057 PSAY IIf(OXR->MES2  > 0, ((OXR->REA2 /OXR->MES2) * 100),0) Picture "@E 999,999,999.99"
         @nLin,072 PSAY IIf(OXR->MES3  > 0, ((OXR->REA3 /OXR->MES3) * 100),0) Picture "@E 999,999,999.99"
         @nLin,087 PSAY IIf(OXR->MES4  > 0, ((OXR->REA4 /OXR->MES4) * 100),0) Picture "@E 999,999,999.99"
         @nLin,102 PSAY IIf(OXR->MES5  > 0, ((OXR->REA5 /OXR->MES5) * 100),0) Picture "@E 999,999,999.99"
         @nLin,117 PSAY IIf(OXR->MES6  > 0, ((OXR->REA6 /OXR->MES6) * 100),0) Picture "@E 999,999,999.99"
         @nLin,132 PSAY IIf(OXR->MES7  > 0, ((OXR->REA7 /OXR->MES7) * 100),0) Picture "@E 999,999,999.99"
         @nLin,147 PSAY IIf(OXR->MES8  > 0, ((OXR->REA8 /OXR->MES8) * 100),0) Picture "@E 999,999,999.99"
         @nLin,162 PSAY IIf(OXR->MES9  > 0, ((OXR->REA9 /OXR->MES9) * 100),0) Picture "@E 999,999,999.99"
         @nLin,177 PSAY IIf(OXR->MES10 > 0, ((OXR->REA10/OXR->MES10)* 100),0) Picture "@E 999,999,999.99"
         @nLin,192 PSAY IIf(OXR->MES11 > 0, ((OXR->REA11/OXR->MES11)* 100),0) Picture "@E 999,999,999.99"
         @nLin,207 PSAY IIf(OXR->MES12 > 0, ((OXR->REA12/OXR->MES12)* 100),0) Picture "@E 999,999,999.99"

         //Acumula Totalizadores dos Or?ados de Gastos Gerais Adm
         nOrc1GADM += OXR->MES1
         nOrc2GADM += OXR->MES2
         nOrc3GADM += OXR->MES3
         nOrc4GADM += OXR->MES4
         nOrc5GADM += OXR->MES5
         nOrc6GADM += OXR->MES6
         nOrc7GADM += OXR->MES7
         nOrc8GADM += OXR->MES8
         nOrc9GADM += OXR->MES9
         nOrc10GADM += OXR->MES10
         nOrc11GADM += OXR->MES11
         nOrc12GADM += OXR->MES12
   
   
         //Acumula Totalizadores dos Realizados de Gastos Gerais Adm
         nRea1GADM  += OXR->REA1
         nRea2GADM  += OXR->REA2
         nRea3GADM  += OXR->REA3
         nRea4GADM  += OXR->REA4
         nRea5GADM  += OXR->REA5
         nRea6GADM  += OXR->REA6
         nRea7GADM  += OXR->REA7
         nRea8GADM  += OXR->REA8
         nRea9GADM  += OXR->REA9
         nRea10GADM += OXR->REA10
         nRea11GADM += OXR->REA11
         nRea12GADM += OXR->REA12

         //Acumula Totalizadores dos A Realizar de Gastos Gerais Adm
         naRea1GADM  += OXR->AREA1
         naRea2GADM  += OXR->AREA2
         naRea3GADM  += OXR->AREA3
         naRea4GADM  += OXR->AREA4
         naRea5GADM  += OXR->AREA5
         naRea6GADM  += OXR->AREA6
         naRea7GADM  += OXR->AREA7
         naRea8GADM  += OXR->AREA8
         naRea9GADM  += OXR->AREA9
         naRea10GADM += OXR->AREA10
         naRea11GADM += OXR->AREA11
         naRea12GADM += OXR->AREA12

         nSld1GG     += (OXR->MES1 - (OXR->REA1+OXR->AREA1))
         nSld2GG     += (OXR->MES2 - (OXR->REA2+OXR->AREA2))
         nSld3GG     += (OXR->MES3 - (OXR->REA3+OXR->AREA3))
         nSld4GG     += (OXR->MES4 - (OXR->REA4+OXR->AREA4))
         nSld5GG     += (OXR->MES5 - (OXR->REA5+OXR->AREA5))
         nSld6GG     += (OXR->MES6 - (OXR->REA6+OXR->AREA6))
         nSld7GG     += (OXR->MES7 - (OXR->REA7+OXR->AREA7))
         nSld8GG     += (OXR->MES8 - (OXR->REA8+OXR->AREA8))
         nSld9GG     += (OXR->MES9 - (OXR->REA9+OXR->AREA9))
         nSld10GG    += (OXR->MES10 - (OXR->REA10+OXR->AREA10))
         nSld11GG    += (OXR->MES11 - (OXR->REA11+OXR->AREA11))
         nSld12GG    += (OXR->MES12 - (OXR->REA12+OXR->AREA12))

         If nVez == 1
            //Acumula Totalizadores Global dos Or?ados de Gastos Gerais/Metas
            nGlbOr1GG += OXR->MES1
            nGlbOr2GG += OXR->MES2
            nGlbOr3GG += OXR->MES3
            nGlbOr4GG += OXR->MES4
            nGlbOr5GG += OXR->MES5
            nGlbOr6GG += OXR->MES6
            nGlbOr7GG += OXR->MES7
            nGlbOr8GG += OXR->MES8
            nGlbOr9GG += OXR->MES9
            nGlbOr10GG += OXR->MES10
            nGlbOr11GG += OXR->MES11
            nGlbOr12GG += OXR->MES12

           //Acumula Totalizadores Global dos Realizados de Gastos Gerais/Metas
           nGbRea1GG += OXR->REA1
           nGbRea2GG += OXR->REA2
           nGbRea3GG += OXR->REA3
           nGbRea4GG += OXR->REA4
           nGbRea5GG += OXR->REA5
           nGbRea6GG += OXR->REA6
           nGbRea7GG += OXR->REA7
           nGbRea8GG += OXR->REA8
           nGbRea9GG += OXR->REA9
           nGbRea10GG += OXR->REA10
           nGbRea11GG += OXR->REA11
           nGbRea12GG += OXR->REA12
           
           //Acumula Totalizadores Global dos A Realizar de Gastos Gerais/Metas
           nGbaRea1GG += OXR->AREA1
           nGbaRea2GG += OXR->AREA2
           nGbaRea3GG += OXR->AREA3
           nGbaRea4GG += OXR->AREA4
           nGbaRea5GG += OXR->AREA5
           nGbaRea6GG += OXR->AREA6
           nGbaRea7GG += OXR->AREA7
           nGbaRea8GG += OXR->AREA8
           nGbaRea9GG += OXR->AREA9
           nGbaRea10GG += OXR->AREA10
           nGbaRea11GG += OXR->AREA11
           nGbaRea12GG += OXR->AREA12           
           
           //Acumula Totalizadores dos Saldos de Gastos Gerais/Metas
           nGbSld1GG     += (OXR->MES1 - (OXR->REA1+OXR->AREA1))
           nGbSld2GG     += (OXR->MES2 - (OXR->REA2+OXR->AREA2))
           nGbSld3GG     += (OXR->MES3 - (OXR->REA3+OXR->AREA3))
           nGbSld4GG     += (OXR->MES4 - (OXR->REA4+OXR->AREA4))
           nGbSld5GG     += (OXR->MES5 - (OXR->REA5+OXR->AREA5))
           nGbSld6GG     += (OXR->MES6 - (OXR->REA6+OXR->AREA6))
           nGbSld7GG     += (OXR->MES7 - (OXR->REA7+OXR->AREA7))
           nGbSld8GG     += (OXR->MES8 - (OXR->REA8+OXR->AREA8))
           nGbSld9GG     += (OXR->MES9 - (OXR->REA9+OXR->AREA9))
           nGbSld10GG    += (OXR->MES10 - (OXR->REA10+OXR->AREA10))
           nGbSld11GG    += (OXR->MES11 - (OXR->REA11+OXR->AREA11))
           nGbSld12GG    += (OXR->MES12 - (OXR->REA12+OXR->AREA12))           
           
         EndIf

         nLin := nLin + 2 // Avanca a linha de impressao

         DbSkip() // Avanca o ponteiro do registro no arquivo
   
      EndDo

      /*************************************************
      * Imprime Totalizadores do Gastos Gerias Adm/Fin (Or?ado / Realizado)
      **************************************************/

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Or?ado de Gastos Ger.Com.Exterior "
      Else
        @nLin,000 PSAY "Tot.Metas Or?ado Gastos Ger.Com.Ext. "
      EndIf
      @nLin,042 PSAY nOrc1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nOrc2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nOrc3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nOrc4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nOrc5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nOrc6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nOrc7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nOrc8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nOrc9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nOrc10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY nOrc11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY nOrc12GADM Picture "@E 999,999,999.99"

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Realiz.Gastos Gerais Com.Ext. "
      Else
         @nLin,000 PSAY "Tot.Metas Realiz.Gastos Ger.Com.Ext."
      EndIf
   
      @nLin,042 PSAY nRea1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nRea2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nRea3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nRea4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nRea5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nRea6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nRea7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nRea8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nRea9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nRea10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY nRea11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY nRea12GADM Picture "@E 999,999,999.99"

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt A Realizar Gastos Ger.Com.Ext. "
      Else
         @nLin,000 PSAY "Tot.Metas A Realiz.Gastos Ger.Com.Ext."
      EndIf
   
      @nLin,042 PSAY naRea1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY naRea2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY naRea3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY naRea4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY naRea5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY naRea6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY naRea7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY naRea8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY naRea9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY naRea10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY naRea11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY naRea12GADM Picture "@E 999,999,999.99"


      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Saldo de Gastos Gerais Adm/Fin "
      Else
         @nLin,000 PSAY "Tot.Metas Saldo Gastos Gerais Adm/Fin. "
      EndIf
   
      @nLin,042 PSAY nSld1GG  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nSld2GG  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nSld3GG  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nSld4GG  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nSld5GG  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nSld6GG  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nSld7GG  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nSld8GG  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nSld9GG  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nSld10GG Picture "@E 999,999,999.99"
      @nLin,192 PSAY nSld11GG Picture "@E 999,999,999.99"
      @nLin,207 PSAY nSld12GG Picture "@E 999,999,999.99"
      
      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt % Utiliz.Gastos Gerais Com.Ext."
      Else
         @nLin,000 PSAY "Tot.Metas % Utiliz.Gastos Ger.Com.Ext."
      EndIf
   
      @nLin,042 PSAY IIf(nOrc1GADM  > 0, ((nRea1GADM /nOrc1GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,057 PSAY IIf(nOrc2GADM  > 0, ((nRea2GADM /nOrc2GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,072 PSAY IIf(nOrc3GADM  > 0, ((nRea3GADM /nOrc3GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,087 PSAY IIf(nOrc4GADM  > 0, ((nRea4GADM /nOrc4GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,102 PSAY IIf(nOrc5GADM  > 0, ((nRea5GADM /nOrc5GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,117 PSAY IIf(nOrc6GADM  > 0, ((nRea6GADM /nOrc6GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,132 PSAY IIf(nOrc7GADM  > 0, ((nRea7GADM /nOrc7GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,147 PSAY IIf(nOrc8GADM  > 0, ((nRea8GADM /nOrc8GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,162 PSAY IIf(nOrc9GADM  > 0, ((nRea9GADM /nOrc9GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,177 PSAY IIf(nOrc10GADM  > 0, ((nRea10GADM /nOrc10GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,192 PSAY IIf(nOrc11GADM  > 0, ((nRea11GADM /nOrc11GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,207 PSAY IIf(nOrc12GADM  > 0, ((nRea12GADM /nOrc12GADM) *100),0) Picture "@E 999,999,999.99"
   
      nLin := 56 //For?a mudan?a de p?gina
      
      //Zera Acumuladores para apresentar valores das naturezas apenas que comp?em metas de or?amentos

      //Or?ado Gastos Gerais Adm
      nOrc1GADM := 0
      nOrc2GADM := 0
      nOrc3GADM := 0
      nOrc4GADM := 0
      nOrc5GADM := 0
      nOrc6GADM := 0
      nOrc7GADM := 0
      nOrc8GADM := 0
      nOrc9GADM := 0
      nOrc10GADM := 0
      nOrc11GADM := 0
      nOrc12GADM := 0

      //Realizado Gastos Gerais Adm
      nRea1GADM := 0
      nRea2GADM := 0
      nRea3GADM := 0
      nRea4GADM := 0
      nRea5GADM := 0
      nRea6GADM := 0
      nRea7GADM := 0
      nRea8GADM := 0
      nRea9GADM := 0
      nRea10GADM := 0
      nRea11GADM := 0
      nRea12GADM := 0

      //A Realizar Gastos Gerais Adm
      naRea1GADM := 0
      naRea2GADM := 0
      naRea3GADM := 0
      naRea4GADM := 0
      naRea5GADM := 0
      naRea6GADM := 0
      naRea7GADM := 0
      naRea8GADM := 0
      naRea9GADM := 0
      naRea10GADM := 0
      naRea11GADM := 0
      naRea12GADM := 0

      //Saldo(Or?ado - (Rrealizado + A Realizar))
      nSld1GG     := 0
      nSld2GG     := 0
      nSld3GG     := 0
      nSld4GG     := 0
      nSld5GG     := 0
      nSld6GG     := 0
      nSld7GG     := 0
      nSld8GG     := 0
      nSld9GG     := 0
      nSld10GG    := 0
      nSld11GG    := 0
      nSld12GG    := 0
      
   Next nVez

EndIf


If lIndustria

   //Cria??o do Arquivo Tempor?rio
   If SELECT("OXR") > 0
      DbSelectArea("OXR")
      DbCloseArea()
   EndIf

   cArqInd := "ORCXREA.CDX"
   While File(cArqInd)
      DELETE FILE &cArqInd
   EndDo

   aCampos := {}
   //Campos para guardar Metas Or?adas para cada natureza
   Aadd(aCampos, {"NATUREZ","C",6,0})
   Aadd(aCampos, {"DESC","C",30,0})
   Aadd(aCampos, {"MES1","N",14,02})
   Aadd(aCampos, {"MES2","N",14,02})
   Aadd(aCampos, {"MES3","N",14,02})
   Aadd(aCampos, {"MES4","N",14,02})
   Aadd(aCampos, {"MES5","N",14,02})
   Aadd(aCampos, {"MES6","N",14,02})
   Aadd(aCampos, {"MES7","N",14,02})
   Aadd(aCampos, {"MES8","N",14,02})
   Aadd(aCampos, {"MES9","N",14,02})
   Aadd(aCampos, {"MES10","N",14,02})
   Aadd(aCampos, {"MES11","N",14,02})
   Aadd(aCampos, {"MES12","N",14,02})
   //Campos para guardar os Gastos Realizados para cada natureza
   Aadd(aCampos, {"REA1","N",14,02})
   Aadd(aCampos, {"REA2","N",14,02})
   Aadd(aCampos, {"REA3","N",14,02})
   Aadd(aCampos, {"REA4","N",14,02})
   Aadd(aCampos, {"REA5","N",14,02})
   Aadd(aCampos, {"REA6","N",14,02})
   Aadd(aCampos, {"REA7","N",14,02})
   Aadd(aCampos, {"REA8","N",14,02})
   Aadd(aCampos, {"REA9","N",14,02})
   Aadd(aCampos, {"REA10","N",14,02})
   Aadd(aCampos, {"REA11","N",14,02})
   Aadd(aCampos, {"REA12","N",14,02})
   //Campos para guardar os Gastos A Realizar para cada natureza
   Aadd(aCampos, {"AREA1","N",14,02})
   Aadd(aCampos, {"AREA2","N",14,02})
   Aadd(aCampos, {"AREA3","N",14,02})
   Aadd(aCampos, {"AREA4","N",14,02})
   Aadd(aCampos, {"AREA5","N",14,02})
   Aadd(aCampos, {"AREA6","N",14,02})
   Aadd(aCampos, {"AREA7","N",14,02})
   Aadd(aCampos, {"AREA8","N",14,02})
   Aadd(aCampos, {"AREA9","N",14,02})
   Aadd(aCampos, {"AREA10","N",14,02})
   Aadd(aCampos, {"AREA11","N",14,02})
   Aadd(aCampos, {"AREA12","N",14,02})
   //Campo para identificar que a natureza do Adm tem Meta de Or?amento
   Aadd(aCampos, {"MTORCADM","C",01,00})


   DbCreate("ORCXREA.DBF",aCampos)

   DbUseArea(.T.,"DBFCDX","ORCXREA.DBF","OXR")

   INDEX ON OXR->NATUREZ TO ORCXREA

   /************************
   *
   * GASTOS GERAIS INDUSTRIA (OR?ADO)
   *
   ************************/
   
      cQuery := " SELECT SED.ED_CODIGO,SE7.E7_ANO,SE7.E7_VALJAN1,SE7.E7_VALFEV1,SE7.E7_VALMAR1,SE7.E7_VALABR1," 
      cQuery += " SE7.E7_VALMAI1,SE7.E7_VALJUN1,SE7.E7_VALJUL1,SE7.E7_VALAGO1,SE7.E7_VALSET1,SE7.E7_VALOUT1,"
      cQuery += " SE7.E7_VALNOV1,SE7.E7_VALDEZ1"
      cQuery += " FROM "+RetSQLName("SED")+" SED LEFT OUTER JOIN "+RetSqlName("SE7")+" SE7 "
      cQuery += "                                ON SED.ED_CODIGO = SE7.E7_NATUREZ"
      cQuery += "                               AND SE7.E7_ANO = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      cQuery += "                               AND SE7.E7_FILIAL = '"+xFilial("SE7")+"'"
      cQuery += "                               AND SE7.D_E_L_E_T_ = ' '"
      cQuery += " WHERE SED.D_E_L_E_T_ = ' '"
      cQuery += "   AND SUBSTRING(SED.ED_CODIGO,1,6) IN "+FormatIn(cGASTGERIND,",")      
      
      MemoWrite("C:\TEMP\ROrcXRea_Or?aGerADM.SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "ORGAD"
      
      TcSetField("ORGAD","E7_VALJAN1","N",14,02)
      TcSetField("ORGAD","E7_VALFEV1","N",14,02)
      TcSetField("ORGAD","E7_VALMAR1","N",14,02)
      TcSetField("ORGAD","E7_VALABR1","N",14,02)
      TcSetField("ORGAD","E7_VALMAI1","N",14,02)
      TcSetField("ORGAD","E7_VALJUN1","N",14,02)
      TcSetField("ORGAD","E7_VALJUL1","N",14,02)
      TcSetField("ORGAD","E7_VALAGO1","N",14,02)
      TcSetField("ORGAD","E7_VALSET1","N",14,02)
      TcSetField("ORGAD","E7_VALOUT1","N",14,02)
      TcSetField("ORGAD","E7_VALNOV1","N",14,02)
      TcSetField("ORGAD","E7_VALDEZ1","N",14,02)
      
      DbSelectArea("ORGAD")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(ORGAD->ED_CODIGO)
            RecLock("OXR",.F.)
               OXR->MES1 += ORGAD->E7_VALJAN1
               OXR->MES2 += ORGAD->E7_VALFEV1               
               OXR->MES3 += ORGAD->E7_VALMAR1
               OXR->MES4 += ORGAD->E7_VALABR1
               OXR->MES5 += ORGAD->E7_VALMAI1
               OXR->MES6 += ORGAD->E7_VALJUN1
               OXR->MES7 += ORGAD->E7_VALJUL1
               OXR->MES8 += ORGAD->E7_VALAGO1
               OXR->MES9 += ORGAD->E7_VALSET1
               OXR->MES10 += ORGAD->E7_VALOUT1
               OXR->MES11 += ORGAD->E7_VALNOV1
               OXR->MES12 += ORGAD->E7_VALDEZ1
              MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := ORGAD->ED_CODIGO
               OXR->MES1 := ORGAD->E7_VALJAN1
               OXR->MES2 := ORGAD->E7_VALFEV1               
               OXR->MES3 := ORGAD->E7_VALMAR1
               OXR->MES4 := ORGAD->E7_VALABR1
               OXR->MES5 := ORGAD->E7_VALMAI1
               OXR->MES6 := ORGAD->E7_VALJUN1
               OXR->MES7 := ORGAD->E7_VALJUL1
               OXR->MES8 := ORGAD->E7_VALAGO1
               OXR->MES9 := ORGAD->E7_VALSET1
               OXR->MES10 := ORGAD->E7_VALOUT1
               OXR->MES11 := ORGAD->E7_VALNOV1
               OXR->MES12 := ORGAD->E7_VALDEZ1
              MsUnLock()
         EndIf
         
         DbSelectArea("ORGAD")
         DbSkip()
         
      EndDo
      DbSelectArea("ORGAD")
      DbCloseArea()


   /************************
   *
   * GASTOS GERAIS INDUSTRIA (REALIZADO)
   *
   ************************/
   
      cQuery := " SELECT SUBSTRING(SE5.E5_DATA,1,6) AS MESGST,SUBSTRING(SE5.E5_NATUREZ,1,6) AS NATUREZ,SUM(SE5.E5_VALOR) AS VLGGADM "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,4) = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      
      cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cGASTGERIND,",")

      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      
      //Implementado em 08/07/2008 - Solicita??o: Patricia
      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCF','PA ','NDF')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"
      cQuery += "  AND SE5.E5_MOTBX <> 'DEV'"      
      
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
      cQuery += " GROUP BY SUBSTRING(SE5.E5_DATA,1,6),SUBSTRING(SE5.E5_NATUREZ,1,6)"
      cQuery += " ORDER BY SUBSTRING(SE5.E5_NATUREZ,1,6)"
      
      MemoWrite("C:\TEMP\ROrcXRea_GastGerADM.SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "GADM"
      
      TcSetField("GADM","VLGGADM","N",14,02)
      
      DbSelectArea("GADM")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(GADM->NATUREZ)
            RecLock("OXR",.F.)
               If Substr(GADM->MESGST,5,2) == "01"
                  OXR->REA1 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "02"
                  OXR->REA2 += GADM->VLGGADM               
               ElseIf Substr(GADM->MESGST,5,2) == "03"
                  OXR->REA3 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "04"
                  OXR->REA4 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "05"
                  OXR->REA5 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "06"
                  OXR->REA6 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "07"
                  OXR->REA7 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "08"
                  OXR->REA8 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "09"
                  OXR->REA9 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "10"
                  OXR->REA10 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "11"
                  OXR->REA11 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "12"
                  OXR->REA12 += GADM->VLGGADM
               EndIf
            MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := GADM->NATUREZ
               If Substr(GADM->MESGST,5,2) == "01"
                  OXR->REA1 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "02"
                  OXR->REA2 := GADM->VLGGADM               
               ElseIf Substr(GADM->MESGST,5,2) == "03"
                  OXR->REA3 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "04"
                  OXR->REA4 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "05"
                  OXR->REA5 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "06"
                  OXR->REA6 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "07"
                  OXR->REA7 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "08"
                  OXR->REA8 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "09"
                  OXR->REA9 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "10"
                  OXR->REA10 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "11"
                  OXR->REA11 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "12"
                  OXR->REA12 := GADM->VLGGADM
               EndIf
            MsUnLock()
         EndIf
         
         DbSelectArea("GADM")
         DbSkip()
         
      EndDo
      DbSelectArea("GADM")
      DbCloseArea()
      
   
*******************************************************************************************************************
*******************************************************************************************************************

   /************************
   *
   * GASTOS GERAIS IND?STRIA (A REALIZAR)
   *
   ************************/

   //For nMes := 1 To 3
      
   //   IncProc("Calc. Pagamentos em Aberto, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryAP := "SELECT SUBSTRING(SE2.E2_NATUREZ,1,6) NATUREZ,SUBSTRING(SE2.E2_VENCREA,1,6) MESAREA,SUM(SE2.E2_SALDO) AS PAGABERTO"
      cQryAP += " FROM "+RetSQLName("SE2")+" SE2"
      cQryAP += " WHERE SE2.E2_FILIAL = '"+xFilial("SE2")+"'"

      cQryAP += "  AND SE2.E2_TIPO NOT IN ('NCF','PA ','NDF')"
      cQryAP += "  AND SUBSTRING(SE2.E2_TIPO,3,1) <> '-'"

      cQryAP += "  AND  SUBSTRING(SE2.E2_VENCREA,1,4) = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      cQryAP += "  AND SUBSTRING(SE2.E2_NATUREZ,1,6) IN "+FormatIn(cGASTGERIND,",")      
      
      cQryAP += "  AND SE2.E2_SALDO > 0 " //Titulos com algum saldo em aberto

      cQryAP += " AND SE2.D_E_L_E_T_ <> '*'"
      cQryAP += " GROUP BY SUBSTRING(SE2.E2_NATUREZ,1,6),SUBSTRING(SE2.E2_VENCREA,1,6)"
      cQryAP += " ORDER BY SUBSTRING(SE2.E2_NATUREZ,1,6)"

      
      //MemoWrite("C:\TEMP\PagamentosemAberto"+Alltrim(Str(nMes))+".SQL",cQryAP)
      
      TCQuery cQryAP NEW ALIAS "TTPG"

      TcSetField("TTPG","PAGABERTO","N",14,02)
      
      DbSelectArea("TTPG")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(TTPG->NATUREZ)
            RecLock("OXR",.F.)
               If Substr(TTPG->MESAREA,5,2) == "01"
                  OXR->AREA1 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "02"
                  OXR->AREA2 += TTPG->PAGABERTO               
               ElseIf Substr(TTPG->MESAREA,5,2) == "03"
                  OXR->AREA3 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "04"
                  OXR->AREA4 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "05"
                  OXR->AREA5 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "06"
                  OXR->AREA6 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "07"
                  OXR->AREA7 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "08"
                  OXR->AREA8 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "09"
                  OXR->AREA9 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "10"
                  OXR->AREA10 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "11"
                  OXR->AREA11 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "12"
                  OXR->AREA12 += TTPG->PAGABERTO
               EndIf
            MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := TTPG->NATUREZ
               If Substr(TTPG->MESAREA,5,2) == "01"
                  OXR->AREA1 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "02"
                  OXR->AREA2 := TTPG->PAGABERTO               
               ElseIf Substr(TTPG->MESAREA,5,2) == "03"
                  OXR->AREA3 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "04"
                  OXR->AREA4 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "05"
                  OXR->AREA5 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "06"
                  OXR->AREA6 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "07"
                  OXR->AREA7 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "08"
                  OXR->AREA8 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "09"
                  OXR->AREA9 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "10"
                  OXR->AREA10 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "11"
                  OXR->AREA11 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "12"
                  OXR->AREA12 := TTPG->PAGABERTO
               EndIf
            MsUnLock()
         EndIf
         
         DbSelectArea("TTPG")
         DbSkip()
         
      EndDo
      DbSelectArea("TTPG")
      DbCloseArea()      
      
   //Next nMes
*******************************************************************************************************************
*******************************************************************************************************************

   //Selecionando e Marcando Naturezas que cont?m Metas de Or?amentos
   cQryMTADM := "SELECT ED_CODIGO "
   cQryMTADM += " FROM "+RetSQLName("SED")
   cQryMTADM += " WHERE SUBSTRING(ED_CODIGO,1,6) IN "+FormatIn(cNtGstMtOrInd,",")
   cQryMTADM += " AND D_E_L_E_T_ <> '*'"
   
   TCQuery cQryMTADM New Alias "MTADM"
   
   DbSelectArea("MTADM")
   While !Eof()
      DbSelectArea("OXR")
      If DbSeek(SUBSTR(MTADM->ED_CODIGO,1,6))
         RecLock("OXR",.F.)
            OXR->MTORCADM := "S"
         MsUnLock()
      EndIf
      DbSelectArea("MTADM")
      DbSkip()
   EndDo
   DbSelectArea("MTADM")
   DbCloseArea()

   

   For nVez := 1 To 2
     
      If nVez == 1
         Cabec1       := "                                                                                                        G A S T O S   G E R A I S   D A   I N D U S T R I A                                                                 "
      Else
         Cabec1       := "                                                                                                   M E T A S   DE   O R ? A M E N T O   D A   I N D U S T R I A                                                             "
      EndIf
      
      DbSelectArea("OXR")
      DbGoTop()
      While !EOF()

      
         If nVez == 2 .And. OXR->MTORCADM <> "S" //Na impress?o do segundo bloco de naturezas, s? apresenta natures
            DbSelectArea("OXR")                  //que comp?em Metas de Or?amentos.
            DbSkip()
            Loop
         EndIf
      
         //?????????????????????????????????????????????????????????????????????Ŀ
         //? Verifica o cancelamento pelo usuario...                             ?
         //???????????????????????????????????????????????????????????????????????
         
         If lAbortPrint
            @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
            Exit
         Endif

         //?????????????????????????????????????????????????????????????????????Ŀ
         //? Impressao do cabecalho do relatorio. . .                            ?
         //???????????????????????????????????????????????????????????????????????

         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
    
         //          10        20        30        40        50        60        70        80        90       100       110       120      130       140        150       160       170       180       190       200      210        220
         // 01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
         //"                                                                                                        G A S T O S   G E R A I S   D A   I N D U S T R I A                                                                 "
         //"                                                                                                   M E T A S   DE   O R ? A M E N T O   D A   I N D U S T R I A                                                             "
         //"Natureza   Descri??o                                       Janeiro     Fevereiro         Mar?o         Abril          Maio         Junho         Julho        Agosto      Setembro       Outubro      Novembro      Dezembro"
         // xxxxxxxxxX xxxxxxxxxXxxxxxxxxxXxxxxxxxxxX      Or?ado 9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99
         //                                             Realizado
         //                                            A Realizar
         //                                                  Real
         //                                            % Utiliza??o

         @nLin,000 PSAY Substr(OXR->NATUREZ,1,6)
         @nLin,011 PSAY Substr(Posicione("SED",1,xFilial("SED")+Substr(OXR->NATUREZ,1,6),"ED_DESCRIC"),1,23)//30) 
         @nLin,035 PSAY "Or?ado"
         @nLin,042 PSAY OXR->MES1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->MES2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->MES3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->MES4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->MES5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->MES6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->MES7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->MES8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->MES9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->MES10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->MES11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->MES12 Picture "@E 999,999,999.99"
   
         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,037 PSAY "Real"
         @nLin,042 PSAY OXR->REA1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->REA2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->REA3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->REA4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->REA5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->REA6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->REA7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->REA8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->REA9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->REA10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->REA11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->REA12 Picture "@E 999,999,999.99"

         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,034 PSAY "ARealiz"
         @nLin,042 PSAY OXR->AREA1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->AREA2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->AREA3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->AREA4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->AREA5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->AREA6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->AREA7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->AREA8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->AREA9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->AREA10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->AREA11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->AREA12 Picture "@E 999,999,999.99"


         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,036 PSAY "Saldo"
         @nLin,042 PSAY (OXR->MES1 - (OXR->REA1+OXR->AREA1))  Picture "@E 999,999,999.99"
         @nLin,057 PSAY (OXR->MES2 - (OXR->REA2+OXR->AREA2))  Picture "@E 999,999,999.99"
         @nLin,072 PSAY (OXR->MES3 - (OXR->REA3+OXR->AREA3))  Picture "@E 999,999,999.99"
         @nLin,087 PSAY (OXR->MES4 - (OXR->REA4+OXR->AREA4))  Picture "@E 999,999,999.99"
         @nLin,102 PSAY (OXR->MES5 - (OXR->REA5+OXR->AREA5))  Picture "@E 999,999,999.99"
         @nLin,117 PSAY (OXR->MES6 - (OXR->REA6+OXR->AREA6))  Picture "@E 999,999,999.99"
         @nLin,132 PSAY (OXR->MES7 - (OXR->REA7+OXR->AREA7))  Picture "@E 999,999,999.99"
         @nLin,147 PSAY (OXR->MES8 - (OXR->REA8+OXR->AREA8))  Picture "@E 999,999,999.99"
         @nLin,162 PSAY (OXR->MES9 - (OXR->REA9+OXR->AREA9))  Picture "@E 999,999,999.99"
         @nLin,177 PSAY (OXR->MES10 - (OXR->REA10+OXR->AREA10)) Picture "@E 999,999,999.99"
         @nLin,192 PSAY (OXR->MES11 - (OXR->REA11+OXR->AREA11)) Picture "@E 999,999,999.99"
         @nLin,207 PSAY (OXR->MES12 - (OXR->REA12+OXR->AREA12)) Picture "@E 999,999,999.99"
         
         nLin := nLin + 1 // Avanca a linha de impressao
         If  nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,034 PSAY "%Utiliz"
         @nLin,042 PSAY IIf(OXR->MES1  > 0, ((OXR->REA1 /OXR->MES1) * 100),0) Picture "@E 999,999,999.99"
         @nLin,057 PSAY IIf(OXR->MES2  > 0, ((OXR->REA2 /OXR->MES2) * 100),0) Picture "@E 999,999,999.99"
         @nLin,072 PSAY IIf(OXR->MES3  > 0, ((OXR->REA3 /OXR->MES3) * 100),0) Picture "@E 999,999,999.99"
         @nLin,087 PSAY IIf(OXR->MES4  > 0, ((OXR->REA4 /OXR->MES4) * 100),0) Picture "@E 999,999,999.99"
         @nLin,102 PSAY IIf(OXR->MES5  > 0, ((OXR->REA5 /OXR->MES5) * 100),0) Picture "@E 999,999,999.99"
         @nLin,117 PSAY IIf(OXR->MES6  > 0, ((OXR->REA6 /OXR->MES6) * 100),0) Picture "@E 999,999,999.99"
         @nLin,132 PSAY IIf(OXR->MES7  > 0, ((OXR->REA7 /OXR->MES7) * 100),0) Picture "@E 999,999,999.99"
         @nLin,147 PSAY IIf(OXR->MES8  > 0, ((OXR->REA8 /OXR->MES8) * 100),0) Picture "@E 999,999,999.99"
         @nLin,162 PSAY IIf(OXR->MES9  > 0, ((OXR->REA9 /OXR->MES9) * 100),0) Picture "@E 999,999,999.99"
         @nLin,177 PSAY IIf(OXR->MES10 > 0, ((OXR->REA10/OXR->MES10)* 100),0) Picture "@E 999,999,999.99"
         @nLin,192 PSAY IIf(OXR->MES11 > 0, ((OXR->REA11/OXR->MES11)* 100),0) Picture "@E 999,999,999.99"
         @nLin,207 PSAY IIf(OXR->MES12 > 0, ((OXR->REA12/OXR->MES12)* 100),0) Picture "@E 999,999,999.99"

         //Acumula Totalizadores dos Or?ados de Gastos Gerais Adm
         nOrc1GADM += OXR->MES1
         nOrc2GADM += OXR->MES2
         nOrc3GADM += OXR->MES3
         nOrc4GADM += OXR->MES4
         nOrc5GADM += OXR->MES5
         nOrc6GADM += OXR->MES6
         nOrc7GADM += OXR->MES7
         nOrc8GADM += OXR->MES8
         nOrc9GADM += OXR->MES9
         nOrc10GADM += OXR->MES10
         nOrc11GADM += OXR->MES11
         nOrc12GADM += OXR->MES12
   
   
         //Acumula Totalizadores dos Realizados de Gastos Gerais Adm
         nRea1GADM  += OXR->REA1
         nRea2GADM  += OXR->REA2
         nRea3GADM  += OXR->REA3
         nRea4GADM  += OXR->REA4
         nRea5GADM  += OXR->REA5
         nRea6GADM  += OXR->REA6
         nRea7GADM  += OXR->REA7
         nRea8GADM  += OXR->REA8
         nRea9GADM  += OXR->REA9
         nRea10GADM += OXR->REA10
         nRea11GADM += OXR->REA11
         nRea12GADM += OXR->REA12

         //Acumula Totalizadores dos A Realizar de Gastos Gerais Adm
         naRea1GADM  += OXR->AREA1
         naRea2GADM  += OXR->AREA2
         naRea3GADM  += OXR->AREA3
         naRea4GADM  += OXR->AREA4
         naRea5GADM  += OXR->AREA5
         naRea6GADM  += OXR->AREA6
         naRea7GADM  += OXR->AREA7
         naRea8GADM  += OXR->AREA8
         naRea9GADM  += OXR->AREA9
         naRea10GADM += OXR->AREA10
         naRea11GADM += OXR->AREA11
         naRea12GADM += OXR->AREA12

         nSld1GG     += (OXR->MES1 - (OXR->REA1+OXR->AREA1))
         nSld2GG     += (OXR->MES2 - (OXR->REA2+OXR->AREA2))
         nSld3GG     += (OXR->MES3 - (OXR->REA3+OXR->AREA3))
         nSld4GG     += (OXR->MES4 - (OXR->REA4+OXR->AREA4))
         nSld5GG     += (OXR->MES5 - (OXR->REA5+OXR->AREA5))
         nSld6GG     += (OXR->MES6 - (OXR->REA6+OXR->AREA6))
         nSld7GG     += (OXR->MES7 - (OXR->REA7+OXR->AREA7))
         nSld8GG     += (OXR->MES8 - (OXR->REA8+OXR->AREA8))
         nSld9GG     += (OXR->MES9 - (OXR->REA9+OXR->AREA9))
         nSld10GG    += (OXR->MES10 - (OXR->REA10+OXR->AREA10))
         nSld11GG    += (OXR->MES11 - (OXR->REA11+OXR->AREA11))
         nSld12GG    += (OXR->MES12 - (OXR->REA12+OXR->AREA12))
         
         If nVez == 1
            //Acumula Totalizadores Global dos Or?ados de Gastos Gerais/Metas
            nGlbOr1GG += OXR->MES1
            nGlbOr2GG += OXR->MES2
            nGlbOr3GG += OXR->MES3
            nGlbOr4GG += OXR->MES4
            nGlbOr5GG += OXR->MES5
            nGlbOr6GG += OXR->MES6
            nGlbOr7GG += OXR->MES7
            nGlbOr8GG += OXR->MES8
            nGlbOr9GG += OXR->MES9
            nGlbOr10GG += OXR->MES10
            nGlbOr11GG += OXR->MES11
            nGlbOr12GG += OXR->MES12

           //Acumula Totalizadores Global dos Realizados de Gastos Gerais/Metas
           nGbRea1GG += OXR->REA1
           nGbRea2GG += OXR->REA2
           nGbRea3GG += OXR->REA3
           nGbRea4GG += OXR->REA4
           nGbRea5GG += OXR->REA5
           nGbRea6GG += OXR->REA6
           nGbRea7GG += OXR->REA7
           nGbRea8GG += OXR->REA8
           nGbRea9GG += OXR->REA9
           nGbRea10GG += OXR->REA10
           nGbRea11GG += OXR->REA11
           nGbRea12GG += OXR->REA12
           
           //Acumula Totalizadores Global dos A Realizar de Gastos Gerais/Metas
           nGbaRea1GG += OXR->AREA1
           nGbaRea2GG += OXR->AREA2
           nGbaRea3GG += OXR->AREA3
           nGbaRea4GG += OXR->AREA4
           nGbaRea5GG += OXR->AREA5
           nGbaRea6GG += OXR->AREA6
           nGbaRea7GG += OXR->AREA7
           nGbaRea8GG += OXR->AREA8
           nGbaRea9GG += OXR->AREA9
           nGbaRea10GG += OXR->AREA10
           nGbaRea11GG += OXR->AREA11
           nGbaRea12GG += OXR->AREA12
           
           //Acumula Totalizadores dos Saldos de Gastos Gerais/Metas
           nGbSld1GG     += (OXR->MES1 - (OXR->REA1+OXR->AREA1))
           nGbSld2GG     += (OXR->MES2 - (OXR->REA2+OXR->AREA2))
           nGbSld3GG     += (OXR->MES3 - (OXR->REA3+OXR->AREA3))
           nGbSld4GG     += (OXR->MES4 - (OXR->REA4+OXR->AREA4))
           nGbSld5GG     += (OXR->MES5 - (OXR->REA5+OXR->AREA5))
           nGbSld6GG     += (OXR->MES6 - (OXR->REA6+OXR->AREA6))
           nGbSld7GG     += (OXR->MES7 - (OXR->REA7+OXR->AREA7))
           nGbSld8GG     += (OXR->MES8 - (OXR->REA8+OXR->AREA8))
           nGbSld9GG     += (OXR->MES9 - (OXR->REA9+OXR->AREA9))
           nGbSld10GG    += (OXR->MES10 - (OXR->REA10+OXR->AREA10))
           nGbSld11GG    += (OXR->MES11 - (OXR->REA11+OXR->AREA11))
           nGbSld12GG    += (OXR->MES12 - (OXR->REA12+OXR->AREA12))           
                      
         EndIf
         
         nLin := nLin + 2 // Avanca a linha de impressao

         DbSkip() // Avanca o ponteiro do registro no arquivo
   
      EndDo

      /*************************************************
      * Imprime Totalizadores do Gastos Gerias Adm/Fin (Or?ado / Realizado)
      **************************************************/

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Or?ado de Gastos Gerais Ind?stria "
      Else
        @nLin,000 PSAY "Tot.Metas Or?ado Gastos Ger.Ind?st. "
      EndIf
      @nLin,042 PSAY nOrc1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nOrc2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nOrc3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nOrc4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nOrc5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nOrc6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nOrc7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nOrc8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nOrc9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nOrc10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY nOrc11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY nOrc12GADM Picture "@E 999,999,999.99"

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Realizado de Gastos Ger.Ind?stria "
      Else
         @nLin,000 PSAY "Tot.Metas Realiz.Gastos Ger. Ind?stria "
      EndIf
   
      @nLin,042 PSAY nRea1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nRea2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nRea3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nRea4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nRea5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nRea6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nRea7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nRea8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nRea9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nRea10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY nRea11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY nRea12GADM Picture "@E 999,999,999.99"


      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt A Realizar Gastos Ger.Ind?stria "
      Else
         @nLin,000 PSAY "Tot.Metas A Realiz.Gastos Ger.Ind?stria "
      EndIf
   
      @nLin,042 PSAY naRea1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY naRea2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY naRea3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY naRea4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY naRea5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY naRea6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY naRea7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY naRea8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY naRea9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY naRea10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY naRea11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY naRea12GADM Picture "@E 999,999,999.99"
               
      

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Saldo de Gastos Gerais Adm/Fin... "
      Else
         @nLin,000 PSAY "Tot.Metas Saldo Gastos Ger.Adm/Fin... "
      EndIf
   
      @nLin,042 PSAY nSld1GG  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nSld2GG  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nSld3GG  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nSld4GG  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nSld5GG  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nSld6GG  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nSld7GG  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nSld8GG  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nSld9GG  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nSld10GG Picture "@E 999,999,999.99"
      @nLin,192 PSAY nSld11GG Picture "@E 999,999,999.99"
      @nLin,207 PSAY nSld12GG Picture "@E 999,999,999.99"

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt % Utiliz.Gastos Gerais Ind?stria"
      Else
         @nLin,000 PSAY "Tot.Metas % Utiliz.Gastos Gerais Ind."
      EndIf
   
      @nLin,042 PSAY IIf(nOrc1GADM  > 0, ((nRea1GADM /nOrc1GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,057 PSAY IIf(nOrc2GADM  > 0, ((nRea2GADM /nOrc2GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,072 PSAY IIf(nOrc3GADM  > 0, ((nRea3GADM /nOrc3GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,087 PSAY IIf(nOrc4GADM  > 0, ((nRea4GADM /nOrc4GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,102 PSAY IIf(nOrc5GADM  > 0, ((nRea5GADM /nOrc5GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,117 PSAY IIf(nOrc6GADM  > 0, ((nRea6GADM /nOrc6GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,132 PSAY IIf(nOrc7GADM  > 0, ((nRea7GADM /nOrc7GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,147 PSAY IIf(nOrc8GADM  > 0, ((nRea8GADM /nOrc8GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,162 PSAY IIf(nOrc9GADM  > 0, ((nRea9GADM /nOrc9GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,177 PSAY IIf(nOrc10GADM  > 0, ((nRea10GADM /nOrc10GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,192 PSAY IIf(nOrc11GADM  > 0, ((nRea11GADM /nOrc11GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,207 PSAY IIf(nOrc12GADM  > 0, ((nRea12GADM /nOrc12GADM) *100),0) Picture "@E 999,999,999.99"
   
      nLin := 56 //For?a mudan?a de p?gina
      
      //Zera Acumuladores para apresentar valores das naturezas apenas que comp?em metas de or?amentos

      //Or?ado Gastos Gerais Adm
      nOrc1GADM := 0
      nOrc2GADM := 0
      nOrc3GADM := 0
      nOrc4GADM := 0
      nOrc5GADM := 0
      nOrc6GADM := 0
      nOrc7GADM := 0
      nOrc8GADM := 0
      nOrc9GADM := 0
      nOrc10GADM := 0
      nOrc11GADM := 0
      nOrc12GADM := 0

      //Realizado Gastos Gerais Adm
      nRea1GADM := 0
      nRea2GADM := 0
      nRea3GADM := 0
      nRea4GADM := 0
      nRea5GADM := 0
      nRea6GADM := 0
      nRea7GADM := 0
      nRea8GADM := 0
      nRea9GADM := 0
      nRea10GADM := 0
      nRea11GADM := 0
      nRea12GADM := 0

      //A Realizar Gastos Gerais Adm
      naRea1GADM := 0
      naRea2GADM := 0
      naRea3GADM := 0
      naRea4GADM := 0
      naRea5GADM := 0
      naRea6GADM := 0
      naRea7GADM := 0
      naRea8GADM := 0
      naRea9GADM := 0
      naRea10GADM := 0
      naRea11GADM := 0
      naRea12GADM := 0

      //Saldo(Or?ado - (Rrealizado + A Realizar))
      nSld1GG     := 0
      nSld2GG     := 0
      nSld3GG     := 0
      nSld4GG     := 0
      nSld5GG     := 0
      nSld6GG     := 0
      nSld7GG     := 0
      nSld8GG     := 0
      nSld9GG     := 0
      nSld10GG    := 0
      nSld11GG    := 0
      nSld12GG    := 0

   Next nVez

EndIf

If lDiversas



   //Cria??o do Arquivo Tempor?rio
   If SELECT("OXR") > 0
      DbSelectArea("OXR")
      DbCloseArea()
   EndIf

   cArqInd := "ORCXREA.CDX"
   While File(cArqInd)
      DELETE FILE &cArqInd
   EndDo

   aCampos := {}
   //Campos para guardar Metas Or?adas para cada natureza
   Aadd(aCampos, {"NATUREZ","C",6,0})
   Aadd(aCampos, {"DESC","C",30,0})
   Aadd(aCampos, {"MES1","N",14,02})
   Aadd(aCampos, {"MES2","N",14,02})
   Aadd(aCampos, {"MES3","N",14,02})
   Aadd(aCampos, {"MES4","N",14,02})
   Aadd(aCampos, {"MES5","N",14,02})
   Aadd(aCampos, {"MES6","N",14,02})
   Aadd(aCampos, {"MES7","N",14,02})
   Aadd(aCampos, {"MES8","N",14,02})
   Aadd(aCampos, {"MES9","N",14,02})
   Aadd(aCampos, {"MES10","N",14,02})
   Aadd(aCampos, {"MES11","N",14,02})
   Aadd(aCampos, {"MES12","N",14,02})
   //Campos para guardar os Gastos Realizados para cada natureza
   Aadd(aCampos, {"REA1","N",14,02})
   Aadd(aCampos, {"REA2","N",14,02})
   Aadd(aCampos, {"REA3","N",14,02})
   Aadd(aCampos, {"REA4","N",14,02})
   Aadd(aCampos, {"REA5","N",14,02})
   Aadd(aCampos, {"REA6","N",14,02})
   Aadd(aCampos, {"REA7","N",14,02})
   Aadd(aCampos, {"REA8","N",14,02})
   Aadd(aCampos, {"REA9","N",14,02})
   Aadd(aCampos, {"REA10","N",14,02})
   Aadd(aCampos, {"REA11","N",14,02})
   Aadd(aCampos, {"REA12","N",14,02})
   //Campos para guardar os Gastos A Realizar para cada natureza
   Aadd(aCampos, {"AREA1","N",14,02})
   Aadd(aCampos, {"AREA2","N",14,02})
   Aadd(aCampos, {"AREA3","N",14,02})
   Aadd(aCampos, {"AREA4","N",14,02})
   Aadd(aCampos, {"AREA5","N",14,02})
   Aadd(aCampos, {"AREA6","N",14,02})
   Aadd(aCampos, {"AREA7","N",14,02})
   Aadd(aCampos, {"AREA8","N",14,02})
   Aadd(aCampos, {"AREA9","N",14,02})
   Aadd(aCampos, {"AREA10","N",14,02})
   Aadd(aCampos, {"AREA11","N",14,02})
   Aadd(aCampos, {"AREA12","N",14,02})
   //Campo para identificar que a natureza do Adm tem Meta de Or?amento
   Aadd(aCampos, {"MTORCADM","C",01,00})


   DbCreate("ORCXREA.DBF",aCampos)

   DbUseArea(.T.,"DBFCDX","ORCXREA.DBF","OXR")

   INDEX ON OXR->NATUREZ TO ORCXREA

   /************************
   *
   * GASTOS GERAIS NATUREZAS DIVERSAS (OR?ADO)
   *
   ************************/
   
      cQuery := " SELECT SED.ED_CODIGO,SE7.E7_ANO,SE7.E7_VALJAN1,SE7.E7_VALFEV1,SE7.E7_VALMAR1,SE7.E7_VALABR1," 
      cQuery += " SE7.E7_VALMAI1,SE7.E7_VALJUN1,SE7.E7_VALJUL1,SE7.E7_VALAGO1,SE7.E7_VALSET1,SE7.E7_VALOUT1,"
      cQuery += " SE7.E7_VALNOV1,SE7.E7_VALDEZ1"
      cQuery += " FROM "+RetSQLName("SED")+" SED LEFT OUTER JOIN "+RetSqlName("SE7")+" SE7 "
      cQuery += "                                ON SED.ED_CODIGO = SE7.E7_NATUREZ"
      cQuery += "                               AND SE7.E7_ANO = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      cQuery += "                               AND SE7.E7_FILIAL = '"+xFilial("SE7")+"'"
      cQuery += "                               AND SE7.D_E_L_E_T_ = ' '"
      cQuery += " WHERE SED.D_E_L_E_T_ = ' '"
      cQuery += "   AND SUBSTRING(SED.ED_CODIGO,1,6) IN "+FormatIn(cGGDiverNat,",")      
      
      MemoWrite("C:\TEMP\ROrcXRea_Or?aGerADM.SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "ORGAD"
      
      TcSetField("ORGAD","E7_VALJAN1","N",14,02)
      TcSetField("ORGAD","E7_VALFEV1","N",14,02)
      TcSetField("ORGAD","E7_VALMAR1","N",14,02)
      TcSetField("ORGAD","E7_VALABR1","N",14,02)
      TcSetField("ORGAD","E7_VALMAI1","N",14,02)
      TcSetField("ORGAD","E7_VALJUN1","N",14,02)
      TcSetField("ORGAD","E7_VALJUL1","N",14,02)
      TcSetField("ORGAD","E7_VALAGO1","N",14,02)
      TcSetField("ORGAD","E7_VALSET1","N",14,02)
      TcSetField("ORGAD","E7_VALOUT1","N",14,02)
      TcSetField("ORGAD","E7_VALNOV1","N",14,02)
      TcSetField("ORGAD","E7_VALDEZ1","N",14,02)
      
      DbSelectArea("ORGAD")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(ORGAD->ED_CODIGO)
            RecLock("OXR",.F.)
               OXR->MES1 += ORGAD->E7_VALJAN1
               OXR->MES2 += ORGAD->E7_VALFEV1               
               OXR->MES3 += ORGAD->E7_VALMAR1
               OXR->MES4 += ORGAD->E7_VALABR1
               OXR->MES5 += ORGAD->E7_VALMAI1
               OXR->MES6 += ORGAD->E7_VALJUN1
               OXR->MES7 += ORGAD->E7_VALJUL1
               OXR->MES8 += ORGAD->E7_VALAGO1
               OXR->MES9 += ORGAD->E7_VALSET1
               OXR->MES10 += ORGAD->E7_VALOUT1
               OXR->MES11 += ORGAD->E7_VALNOV1
               OXR->MES12 += ORGAD->E7_VALDEZ1
              MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := ORGAD->ED_CODIGO
               OXR->MES1 := ORGAD->E7_VALJAN1
               OXR->MES2 := ORGAD->E7_VALFEV1               
               OXR->MES3 := ORGAD->E7_VALMAR1
               OXR->MES4 := ORGAD->E7_VALABR1
               OXR->MES5 := ORGAD->E7_VALMAI1
               OXR->MES6 := ORGAD->E7_VALJUN1
               OXR->MES7 := ORGAD->E7_VALJUL1
               OXR->MES8 := ORGAD->E7_VALAGO1
               OXR->MES9 := ORGAD->E7_VALSET1
               OXR->MES10 := ORGAD->E7_VALOUT1
               OXR->MES11 := ORGAD->E7_VALNOV1
               OXR->MES12 := ORGAD->E7_VALDEZ1
              MsUnLock()
         EndIf
         
         DbSelectArea("ORGAD")
         DbSkip()
         
      EndDo
      DbSelectArea("ORGAD")
      DbCloseArea()


   /************************
   *
   * GASTOS GERAIS NATUREZAS DIVERSAS (REALIZADO)
   *
   ************************/
   
      cQuery := " SELECT SUBSTRING(SE5.E5_DATA,1,6) AS MESGST,SUBSTRING(SE5.E5_NATUREZ,1,6) AS NATUREZ,SUM(SE5.E5_VALOR) AS VLGGADM "
      cQuery += " FROM " + RetSqlName("SE5")+" SE5 "
      cQuery += " WHERE SE5.E5_RECPAG = 'P'"
      
      cQuery += "    AND  SUBSTRING(SE5.E5_DATA,1,4) = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      
      cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) IN "+FormatIn(cGGDiverNat,",")

      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
      
      //Implementado em 08/07/2008 - Solicita??o: Patricia
      cQuery += "  AND SE5.E5_TIPO NOT IN ('NCF','PA ','NDF')"
      cQuery += "  AND SUBSTRING(SE5.E5_TIPO,3,1) <> '-'"
      cQuery += "  AND SE5.E5_MOTBX <> 'DEV'"      
      
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
      cQuery += " GROUP BY SUBSTRING(SE5.E5_DATA,1,6),SUBSTRING(SE5.E5_NATUREZ,1,6)"
      cQuery += " ORDER BY SUBSTRING(SE5.E5_NATUREZ,1,6)"
      
      MemoWrite("C:\TEMP\ROrcXRea_GastGerADM.SQL",cQuery)
      
      TCQuery cQuery NEW ALIAS "GADM"
      
      TcSetField("GADM","VLGGADM","N",14,02)
      
      DbSelectArea("GADM")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(GADM->NATUREZ)
            RecLock("OXR",.F.)
               If Substr(GADM->MESGST,5,2) == "01"
                  OXR->REA1 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "02"
                  OXR->REA2 += GADM->VLGGADM               
               ElseIf Substr(GADM->MESGST,5,2) == "03"
                  OXR->REA3 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "04"
                  OXR->REA4 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "05"
                  OXR->REA5 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "06"
                  OXR->REA6 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "07"
                  OXR->REA7 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "08"
                  OXR->REA8 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "09"
                  OXR->REA9 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "10"
                  OXR->REA10 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "11"
                  OXR->REA11 += GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "12"
                  OXR->REA12 += GADM->VLGGADM
               EndIf
            MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := GADM->NATUREZ
               If Substr(GADM->MESGST,5,2) == "01"
                  OXR->REA1 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "02"
                  OXR->REA2 := GADM->VLGGADM               
               ElseIf Substr(GADM->MESGST,5,2) == "03"
                  OXR->REA3 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "04"
                  OXR->REA4 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "05"
                  OXR->REA5 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "06"
                  OXR->REA6 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "07"
                  OXR->REA7 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "08"
                  OXR->REA8 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "09"
                  OXR->REA9 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "10"
                  OXR->REA10 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "11"
                  OXR->REA11 := GADM->VLGGADM
               ElseIf Substr(GADM->MESGST,5,2) == "12"
                  OXR->REA12 := GADM->VLGGADM
               EndIf
            MsUnLock()
         EndIf
         
         DbSelectArea("GADM")
         DbSkip()
         
      EndDo
      DbSelectArea("GADM")
      DbCloseArea()
      
   
*******************************************************************************************************************
*******************************************************************************************************************

   /************************
   *
   * GASTOS GERAIS NATUREZAS DIVERSAS  (A REALIZAR)
   *
   ************************/

   //For nMes := 1 To 3
      
   //   IncProc("Calc. Pagamentos em Aberto, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryAP := "SELECT SUBSTRING(SE2.E2_NATUREZ,1,6) NATUREZ,SUBSTRING(SE2.E2_VENCREA,1,6) MESAREA,SUM(SE2.E2_SALDO) AS PAGABERTO"
      cQryAP += " FROM "+RetSQLName("SE2")+" SE2"
      cQryAP += " WHERE SE2.E2_FILIAL = '"+xFilial("SE2")+"'"

      cQryAP += "  AND SE2.E2_TIPO NOT IN ('NCF','PA ','NDF')"
      cQryAP += "  AND SUBSTRING(SE2.E2_TIPO,3,1) <> '-'"

      cQryAP += "  AND  SUBSTRING(SE2.E2_VENCREA,1,4) = '" +Substr(DTOS(dDataRel),1,4)+ "'"
      cQryAP += "  AND SUBSTRING(SE2.E2_NATUREZ,1,6) IN "+FormatIn(cGGDiverNat,",")      
      
      cQryAP += "  AND SE2.E2_SALDO > 0 " //Titulos com algum saldo em aberto

      cQryAP += " AND SE2.D_E_L_E_T_ <> '*'"
      cQryAP += " GROUP BY SUBSTRING(SE2.E2_NATUREZ,1,6),SUBSTRING(SE2.E2_VENCREA,1,6)"
      cQryAP += " ORDER BY SUBSTRING(SE2.E2_NATUREZ,1,6)"

      
      //MemoWrite("C:\TEMP\PagamentosemAberto"+Alltrim(Str(nMes))+".SQL",cQryAP)
      
      TCQuery cQryAP NEW ALIAS "TTPG"

      TcSetField("TTPG","PAGABERTO","N",14,02)
      
      DbSelectArea("TTPG")
      While !Eof()
         
         DbSelectArea("OXR")
         If DbSeek(TTPG->NATUREZ)
            RecLock("OXR",.F.)
               If Substr(TTPG->MESAREA,5,2) == "01"
                  OXR->AREA1 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "02"
                  OXR->AREA2 += TTPG->PAGABERTO               
               ElseIf Substr(TTPG->MESAREA,5,2) == "03"
                  OXR->AREA3 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "04"
                  OXR->AREA4 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "05"
                  OXR->AREA5 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "06"
                  OXR->AREA6 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "07"
                  OXR->AREA7 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "08"
                  OXR->AREA8 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "09"
                  OXR->AREA9 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "10"
                  OXR->AREA10 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "11"
                  OXR->AREA11 += TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "12"
                  OXR->AREA12 += TTPG->PAGABERTO
               EndIf
            MsUnLock()
         Else
            RecLock("OXR",.T.)
               OXR->NATUREZ := TTPG->NATUREZ
               If Substr(TTPG->MESAREA,5,2) == "01"
                  OXR->AREA1 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "02"
                  OXR->AREA2 := TTPG->PAGABERTO               
               ElseIf Substr(TTPG->MESAREA,5,2) == "03"
                  OXR->AREA3 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "04"
                  OXR->AREA4 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "05"
                  OXR->AREA5 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "06"
                  OXR->AREA6 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "07"
                  OXR->AREA7 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "08"
                  OXR->AREA8 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "09"
                  OXR->AREA9 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "10"
                  OXR->AREA10 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "11"
                  OXR->AREA11 := TTPG->PAGABERTO
               ElseIf Substr(TTPG->MESAREA,5,2) == "12"
                  OXR->AREA12 := TTPG->PAGABERTO
               EndIf
            MsUnLock()
         EndIf
         
         DbSelectArea("TTPG")
         DbSkip()
         
      EndDo
      DbSelectArea("TTPG")
      DbCloseArea()      
      
   //Next nMes
*******************************************************************************************************************
*******************************************************************************************************************

   //Selecionando e Marcando Naturezas que cont?m Metas de Or?amentos
   cQryMTADM := "SELECT ED_CODIGO "
   cQryMTADM += " FROM "+RetSQLName("SED")
   cQryMTADM += " WHERE SUBSTRING(ED_CODIGO,1,6) IN "+FormatIn(cMtDiverNat,",")
   cQryMTADM += " AND D_E_L_E_T_ <> '*'"
   
   TCQuery cQryMTADM New Alias "MTADM"
   
   DbSelectArea("MTADM")
   While !Eof()
      DbSelectArea("OXR")
      If DbSeek(SUBSTR(MTADM->ED_CODIGO,1,6))
         RecLock("OXR",.F.)
            OXR->MTORCADM := "S"
         MsUnLock()
      EndIf
      DbSelectArea("MTADM")
      DbSkip()
   EndDo
   DbSelectArea("MTADM")
   DbCloseArea()

   

   For nVez := 1 To 1 //2 - Especificamente p/ Naturezas Diversas, o sistema s? apresentar? Gastos Gerais.
     
      If nVez == 1
         Cabec1       := "                                                                                                    G A S T O S   G E R A I S   D E   N A T U R E Z A S   D I V E R S A S                                                   "
      Else
         Cabec1       := "                                                                                                M E T A S   DE   O R ? A M E N T O   D O   N A T U R E Z A S   D I V E R S A S                                              "
      EndIf
      
      DbSelectArea("OXR")
      DbGoTop()
      While !EOF()

      
         If nVez == 2 .And. OXR->MTORCADM <> "S" //Na impress?o do segundo bloco de naturezas, s? apresenta natures
            DbSelectArea("OXR")                  //que comp?em Metas de Or?amentos.
            DbSkip()
            Loop
         EndIf
      
         //?????????????????????????????????????????????????????????????????????Ŀ
         //? Verifica o cancelamento pelo usuario...                             ?
         //???????????????????????????????????????????????????????????????????????
         
         If lAbortPrint
            @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
            Exit
         Endif

         //?????????????????????????????????????????????????????????????????????Ŀ
         //? Impressao do cabecalho do relatorio. . .                            ?
         //???????????????????????????????????????????????????????????????????????

         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
    
         //          10        20        30        40        50        60        70        80        90       100       110       120      130       140        150       160       170       180       190       200      210        220
         // 01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
         //"                                                                                                    G A S T O S   G E R A I S   D E   N A T U R E Z A S   D I V E R S A S                                                   "
         //"                                                                                                M E T A S   DE   O R ? A M E N T O   D O   N A T U R E Z A S   D I V E R S A S                                              "
         //"Natureza   Descri??o                                       Janeiro     Fevereiro         Mar?o         Abril          Maio         Junho         Julho        Agosto      Setembro       Outubro      Novembro      Dezembro"
         // xxxxxxxxxX xxxxxxxxxXxxxxxxxxxXxxxxxxxxxX      Or?ado 9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99  9,999,999.99
         //                                             Realizado
         //                                            A Realizar
         //                                                  Real
         //                                            % Utiliza??o

         @nLin,000 PSAY Substr(OXR->NATUREZ,1,6)
         @nLin,011 PSAY Substr(Posicione("SED",1,xFilial("SED")+Substr(OXR->NATUREZ,1,6),"ED_DESCRIC"),1,23)//30) 
         @nLin,035 PSAY "Or?ado"
         @nLin,042 PSAY OXR->MES1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->MES2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->MES3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->MES4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->MES5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->MES6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->MES7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->MES8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->MES9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->MES10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->MES11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->MES12 Picture "@E 999,999,999.99"
   
         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,037 PSAY "Real"
         @nLin,042 PSAY OXR->REA1  Picture "@E 999,999,999.99"
         @nLin,057 PSAY OXR->REA2  Picture "@E 999,999,999.99"
         @nLin,072 PSAY OXR->REA3  Picture "@E 999,999,999.99"
         @nLin,087 PSAY OXR->REA4  Picture "@E 999,999,999.99"
         @nLin,102 PSAY OXR->REA5  Picture "@E 999,999,999.99"
         @nLin,117 PSAY OXR->REA6  Picture "@E 999,999,999.99"
         @nLin,132 PSAY OXR->REA7  Picture "@E 999,999,999.99"
         @nLin,147 PSAY OXR->REA8  Picture "@E 999,999,999.99"
         @nLin,162 PSAY OXR->REA9  Picture "@E 999,999,999.99"
         @nLin,177 PSAY OXR->REA10 Picture "@E 999,999,999.99"
         @nLin,192 PSAY OXR->REA11 Picture "@E 999,999,999.99"
         @nLin,207 PSAY OXR->REA12 Picture "@E 999,999,999.99"

         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,034 PSAY "ARealiz"
         @nLin,042 PSAY OXR->AREA1  Picture "@E 9,999,999.99"
         @nLin,057 PSAY OXR->AREA2  Picture "@E 9,999,999.99"
         @nLin,072 PSAY OXR->AREA3  Picture "@E 9,999,999.99"
         @nLin,087 PSAY OXR->AREA4  Picture "@E 9,999,999.99"
         @nLin,102 PSAY OXR->AREA5  Picture "@E 9,999,999.99"
         @nLin,117 PSAY OXR->AREA6  Picture "@E 9,999,999.99"
         @nLin,132 PSAY OXR->AREA7  Picture "@E 9,999,999.99"
         @nLin,147 PSAY OXR->AREA8  Picture "@E 9,999,999.99"
         @nLin,162 PSAY OXR->AREA9  Picture "@E 9,999,999.99"
         @nLin,177 PSAY OXR->AREA10 Picture "@E 9,999,999.99"
         @nLin,192 PSAY OXR->AREA11 Picture "@E 9,999,999.99"
         @nLin,207 PSAY OXR->AREA12 Picture "@E 9,999,999.99"
         

         nLin := nLin + 1 // Avanca a linha de impressao
         If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,036 PSAY "Saldo"
         @nLin,042 PSAY (OXR->MES1 - (OXR->REA1+OXR->AREA1))  Picture "@E 999,999,999.99"
         @nLin,057 PSAY (OXR->MES2 - (OXR->REA2+OXR->AREA2))  Picture "@E 999,999,999.99"
         @nLin,072 PSAY (OXR->MES3 - (OXR->REA3+OXR->AREA3))  Picture "@E 999,999,999.99"
         @nLin,087 PSAY (OXR->MES4 - (OXR->REA4+OXR->AREA4))  Picture "@E 999,999,999.99"
         @nLin,102 PSAY (OXR->MES5 - (OXR->REA5+OXR->AREA5))  Picture "@E 999,999,999.99"
         @nLin,117 PSAY (OXR->MES6 - (OXR->REA6+OXR->AREA6))  Picture "@E 999,999,999.99"
         @nLin,132 PSAY (OXR->MES7 - (OXR->REA7+OXR->AREA7))  Picture "@E 999,999,999.99"
         @nLin,147 PSAY (OXR->MES8 - (OXR->REA8+OXR->AREA8))  Picture "@E 999,999,999.99"
         @nLin,162 PSAY (OXR->MES9 - (OXR->REA9+OXR->AREA9))  Picture "@E 999,999,999.99"
         @nLin,177 PSAY (OXR->MES10 - (OXR->REA10+OXR->AREA10)) Picture "@E 999,999,999.99"
         @nLin,192 PSAY (OXR->MES11 - (OXR->REA11+OXR->AREA11)) Picture "@E 999,999,999.99"
         @nLin,207 PSAY (OXR->MES12 - (OXR->REA12+OXR->AREA12)) Picture "@E 999,999,999.99"

         nLin := nLin + 1 // Avanca a linha de impressao
         If  nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 9 //8
         Endif
   
         @nLin,034 PSAY "%Utiliz"
         @nLin,042 PSAY IIf(OXR->MES1  > 0, ((OXR->REA1 /OXR->MES1) * 100),0) Picture "@E 999,999,999.99"
         @nLin,057 PSAY IIf(OXR->MES2  > 0, ((OXR->REA2 /OXR->MES2) * 100),0) Picture "@E 999,999,999.99"
         @nLin,072 PSAY IIf(OXR->MES3  > 0, ((OXR->REA3 /OXR->MES3) * 100),0) Picture "@E 999,999,999.99"
         @nLin,087 PSAY IIf(OXR->MES4  > 0, ((OXR->REA4 /OXR->MES4) * 100),0) Picture "@E 999,999,999.99"
         @nLin,102 PSAY IIf(OXR->MES5  > 0, ((OXR->REA5 /OXR->MES5) * 100),0) Picture "@E 999,999,999.99"
         @nLin,117 PSAY IIf(OXR->MES6  > 0, ((OXR->REA6 /OXR->MES6) * 100),0) Picture "@E 999,999,999.99"
         @nLin,132 PSAY IIf(OXR->MES7  > 0, ((OXR->REA7 /OXR->MES7) * 100),0) Picture "@E 999,999,999.99"
         @nLin,147 PSAY IIf(OXR->MES8  > 0, ((OXR->REA8 /OXR->MES8) * 100),0) Picture "@E 999,999,999.99"
         @nLin,162 PSAY IIf(OXR->MES9  > 0, ((OXR->REA9 /OXR->MES9) * 100),0) Picture "@E 999,999,999.99"
         @nLin,177 PSAY IIf(OXR->MES10 > 0, ((OXR->REA10/OXR->MES10)* 100),0) Picture "@E 999,999,999.99"
         @nLin,192 PSAY IIf(OXR->MES11 > 0, ((OXR->REA11/OXR->MES11)* 100),0) Picture "@E 999,999,999.99"
         @nLin,207 PSAY IIf(OXR->MES12 > 0, ((OXR->REA12/OXR->MES12)* 100),0) Picture "@E 999,999,999.99"

         //Acumula Totalizadores dos Or?ados de Gastos Gerais Adm
         nOrc1GADM += OXR->MES1
         nOrc2GADM += OXR->MES2
         nOrc3GADM += OXR->MES3
         nOrc4GADM += OXR->MES4
         nOrc5GADM += OXR->MES5
         nOrc6GADM += OXR->MES6
         nOrc7GADM += OXR->MES7
         nOrc8GADM += OXR->MES8
         nOrc9GADM += OXR->MES9
         nOrc10GADM += OXR->MES10
         nOrc11GADM += OXR->MES11
         nOrc12GADM += OXR->MES12
   
   
         //Acumula Totalizadores dos Realizados de Gastos Gerais Adm
         nRea1GADM  += OXR->REA1
         nRea2GADM  += OXR->REA2
         nRea3GADM  += OXR->REA3
         nRea4GADM  += OXR->REA4
         nRea5GADM  += OXR->REA5
         nRea6GADM  += OXR->REA6
         nRea7GADM  += OXR->REA7
         nRea8GADM  += OXR->REA8
         nRea9GADM  += OXR->REA9
         nRea10GADM += OXR->REA10
         nRea11GADM += OXR->REA11
         nRea12GADM += OXR->REA12

         //Acumula Totalizadores dos A Realizar de Gastos Gerais Adm
         naRea1GADM  += OXR->AREA1
         naRea2GADM  += OXR->AREA2
         naRea3GADM  += OXR->AREA3
         naRea4GADM  += OXR->AREA4
         naRea5GADM  += OXR->AREA5
         naRea6GADM  += OXR->AREA6
         naRea7GADM  += OXR->AREA7
         naRea8GADM  += OXR->AREA8
         naRea9GADM  += OXR->AREA9
         naRea10GADM += OXR->AREA10
         naRea11GADM += OXR->AREA11
         naRea12GADM += OXR->AREA12

         nSld1GG     += (OXR->MES1 - (OXR->REA1+OXR->AREA1))
         nSld2GG     += (OXR->MES2 - (OXR->REA2+OXR->AREA2))
         nSld3GG     += (OXR->MES3 - (OXR->REA3+OXR->AREA3))
         nSld4GG     += (OXR->MES4 - (OXR->REA4+OXR->AREA4))
         nSld5GG     += (OXR->MES5 - (OXR->REA5+OXR->AREA5))
         nSld6GG     += (OXR->MES6 - (OXR->REA6+OXR->AREA6))
         nSld7GG     += (OXR->MES7 - (OXR->REA7+OXR->AREA7))
         nSld8GG     += (OXR->MES8 - (OXR->REA8+OXR->AREA8))
         nSld9GG     += (OXR->MES9 - (OXR->REA9+OXR->AREA9))
         nSld10GG    += (OXR->MES10 - (OXR->REA10+OXR->AREA10))
         nSld11GG    += (OXR->MES11 - (OXR->REA11+OXR->AREA11))
         nSld12GG    += (OXR->MES12 - (OXR->REA12+OXR->AREA12))

         If nVez == 1
            //Acumula Totalizadores Global dos Or?ados de Gastos Gerais/Metas
            nGlbOr1GG += OXR->MES1
            nGlbOr2GG += OXR->MES2
            nGlbOr3GG += OXR->MES3
            nGlbOr4GG += OXR->MES4
            nGlbOr5GG += OXR->MES5
            nGlbOr6GG += OXR->MES6
            nGlbOr7GG += OXR->MES7
            nGlbOr8GG += OXR->MES8
            nGlbOr9GG += OXR->MES9
            nGlbOr10GG += OXR->MES10
            nGlbOr11GG += OXR->MES11
            nGlbOr12GG += OXR->MES12

           //Acumula Totalizadores Global dos Realizados de Gastos Gerais/Metas
           nGbRea1GG += OXR->REA1
           nGbRea2GG += OXR->REA2
           nGbRea3GG += OXR->REA3
           nGbRea4GG += OXR->REA4
           nGbRea5GG += OXR->REA5
           nGbRea6GG += OXR->REA6
           nGbRea7GG += OXR->REA7
           nGbRea8GG += OXR->REA8
           nGbRea9GG += OXR->REA9
           nGbRea10GG += OXR->REA10
           nGbRea11GG += OXR->REA11
           nGbRea12GG += OXR->REA12

           //Acumula Totalizadores Global dos A Realizar de Gastos Gerais/Metas
           nGbaRea1GG += OXR->AREA1
           nGbaRea2GG += OXR->AREA2
           nGbaRea3GG += OXR->AREA3
           nGbaRea4GG += OXR->AREA4
           nGbaRea5GG += OXR->AREA5
           nGbaRea6GG += OXR->AREA6
           nGbaRea7GG += OXR->AREA7
           nGbaRea8GG += OXR->AREA8
           nGbaRea9GG += OXR->AREA9
           nGbaRea10GG += OXR->AREA10
           nGbaRea11GG += OXR->AREA11
           nGbaRea12GG += OXR->AREA12
           
           //Acumula Totalizadores dos Saldos de Gastos Gerais/Metas
           nGbSld1GG     += (OXR->MES1 - (OXR->REA1+OXR->AREA1))
           nGbSld2GG     += (OXR->MES2 - (OXR->REA2+OXR->AREA2))
           nGbSld3GG     += (OXR->MES3 - (OXR->REA3+OXR->AREA3))
           nGbSld4GG     += (OXR->MES4 - (OXR->REA4+OXR->AREA4))
           nGbSld5GG     += (OXR->MES5 - (OXR->REA5+OXR->AREA5))
           nGbSld6GG     += (OXR->MES6 - (OXR->REA6+OXR->AREA6))
           nGbSld7GG     += (OXR->MES7 - (OXR->REA7+OXR->AREA7))
           nGbSld8GG     += (OXR->MES8 - (OXR->REA8+OXR->AREA8))
           nGbSld9GG     += (OXR->MES9 - (OXR->REA9+OXR->AREA9))
           nGbSld10GG    += (OXR->MES10 - (OXR->REA10+OXR->AREA10))
           nGbSld11GG    += (OXR->MES11 - (OXR->REA11+OXR->AREA11))
           nGbSld12GG    += (OXR->MES12 - (OXR->REA12+OXR->AREA12))

         EndIf

         nLin := nLin + 2 // Avanca a linha de impressao

         DbSkip() // Avanca o ponteiro do registro no arquivo
   
      EndDo

      /*************************************************
      * Imprime Totalizadores do Gastos Gerias Adm/Fin (Or?ado / Realizado)
      **************************************************/

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Or?ado de Gastos Gerais Nat.Diversas "
      Else
        @nLin,000 PSAY "Tot.Metas Or?.Gastos Gerais Nat.Diversas "
      EndIf
      @nLin,042 PSAY nOrc1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nOrc2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nOrc3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nOrc4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nOrc5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nOrc6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nOrc7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nOrc8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nOrc9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nOrc10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY nOrc11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY nOrc12GADM Picture "@E 999,999,999.99"

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Realizado Gastos Ger.Nat.Diversas "
      Else
         @nLin,000 PSAY "Tot.Metas Realiz.Gastos Ger.Nat.Div. "
      EndIf
   
      @nLin,042 PSAY nRea1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nRea2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nRea3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nRea4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nRea5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nRea6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nRea7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nRea8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nRea9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nRea10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY nRea11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY nRea12GADM Picture "@E 999,999,999.99"
               

      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt A Realiz.Gastos Gerais Nat.Diversas "
      Else
         @nLin,000 PSAY "Tot.Metas A Realiz.Gastos Ger.Nat.Div. "
      EndIf
   
      @nLin,042 PSAY naRea1GADM  Picture "@E 999,999,999.99"
      @nLin,057 PSAY naRea2GADM  Picture "@E 999,999,999.99"
      @nLin,072 PSAY naRea3GADM  Picture "@E 999,999,999.99"
      @nLin,087 PSAY naRea4GADM  Picture "@E 999,999,999.99"
      @nLin,102 PSAY naRea5GADM  Picture "@E 999,999,999.99"
      @nLin,117 PSAY naRea6GADM  Picture "@E 999,999,999.99"
      @nLin,132 PSAY naRea7GADM  Picture "@E 999,999,999.99"
      @nLin,147 PSAY naRea8GADM  Picture "@E 999,999,999.99"
      @nLin,162 PSAY naRea9GADM  Picture "@E 999,999,999.99"
      @nLin,177 PSAY naRea10GADM Picture "@E 999,999,999.99"
      @nLin,192 PSAY naRea11GADM Picture "@E 999,999,999.99"
      @nLin,207 PSAY naRea12GADM Picture "@E 999,999,999.99"


      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt Saldo de Gastos Gerais Adm/Fin "
      Else
         @nLin,000 PSAY "Tot.Metas Saldo Gastos Ger.Adm/Fin "
      EndIf
   
      @nLin,042 PSAY nSld1GG  Picture "@E 999,999,999.99"
      @nLin,057 PSAY nSld2GG  Picture "@E 999,999,999.99"
      @nLin,072 PSAY nSld3GG  Picture "@E 999,999,999.99"
      @nLin,087 PSAY nSld4GG  Picture "@E 999,999,999.99"
      @nLin,102 PSAY nSld5GG  Picture "@E 999,999,999.99"
      @nLin,117 PSAY nSld6GG  Picture "@E 999,999,999.99"
      @nLin,132 PSAY nSld7GG  Picture "@E 999,999,999.99"
      @nLin,147 PSAY nSld8GG  Picture "@E 999,999,999.99"
      @nLin,162 PSAY nSld9GG  Picture "@E 999,999,999.99"
      @nLin,177 PSAY nSld10GG Picture "@E 999,999,999.99"
      @nLin,192 PSAY nSld11GG Picture "@E 999,999,999.99"
      @nLin,207 PSAY nSld12GG Picture "@E 999,999,999.99"
      
      nLin := nLin + 1 // Avanca a linha de impressao

      If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9 //8
      Endif

      If nVez == 1
         @nLin,000 PSAY "Tt % Utiliz.Gastos Ger.Nat.Diversas"
      Else
         @nLin,000 PSAY "Tot.Metas %Utiliz.Gastos Ger.Nat.Diversas"
      EndIf
   
      @nLin,042 PSAY IIf(nOrc1GADM  > 0, ((nRea1GADM /nOrc1GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,057 PSAY IIf(nOrc2GADM  > 0, ((nRea2GADM /nOrc2GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,072 PSAY IIf(nOrc3GADM  > 0, ((nRea3GADM /nOrc3GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,087 PSAY IIf(nOrc4GADM  > 0, ((nRea4GADM /nOrc4GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,102 PSAY IIf(nOrc5GADM  > 0, ((nRea5GADM /nOrc5GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,117 PSAY IIf(nOrc6GADM  > 0, ((nRea6GADM /nOrc6GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,132 PSAY IIf(nOrc7GADM  > 0, ((nRea7GADM /nOrc7GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,147 PSAY IIf(nOrc8GADM  > 0, ((nRea8GADM /nOrc8GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,162 PSAY IIf(nOrc9GADM  > 0, ((nRea9GADM /nOrc9GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,177 PSAY IIf(nOrc10GADM  > 0, ((nRea10GADM /nOrc10GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,192 PSAY IIf(nOrc11GADM  > 0, ((nRea11GADM /nOrc11GADM) *100),0) Picture "@E 999,999,999.99"
      @nLin,207 PSAY IIf(nOrc12GADM  > 0, ((nRea12GADM /nOrc12GADM) *100),0) Picture "@E 999,999,999.99"
   
      nLin := 56 //For?a mudan?a de p?gina
      
      //Zera Acumuladores para apresentar valores das naturezas apenas que comp?em metas de or?amentos

      //Or?ado Gastos Gerais Adm
      nOrc1GADM := 0
      nOrc2GADM := 0
      nOrc3GADM := 0
      nOrc4GADM := 0
      nOrc5GADM := 0
      nOrc6GADM := 0
      nOrc7GADM := 0
      nOrc8GADM := 0
      nOrc9GADM := 0
      nOrc10GADM := 0
      nOrc11GADM := 0
      nOrc12GADM := 0

      //Realizado Gastos Gerais Adm
      nRea1GADM := 0
      nRea2GADM := 0
      nRea3GADM := 0
      nRea4GADM := 0
      nRea5GADM := 0
      nRea6GADM := 0
      nRea7GADM := 0
      nRea8GADM := 0
      nRea9GADM := 0
      nRea10GADM := 0
      nRea11GADM := 0
      nRea12GADM := 0

      //A Realizar Gastos Gerais Adm
      naRea1GADM := 0
      naRea2GADM := 0
      naRea3GADM := 0
      naRea4GADM := 0
      naRea5GADM := 0
      naRea6GADM := 0
      naRea7GADM := 0
      naRea8GADM := 0
      naRea9GADM := 0
      naRea10GADM := 0
      naRea11GADM := 0
      naRea12GADM := 0

      //Saldo(Or?ado - (Rrealizado + A Realizar))
      nSld1GG     := 0
      nSld2GG     := 0
      nSld3GG     := 0
      nSld4GG     := 0
      nSld5GG     := 0
      nSld6GG     := 0
      nSld7GG     := 0
      nSld8GG     := 0
      nSld9GG     := 0
      nSld10GG    := 0
      nSld11GG    := 0
      nSld12GG    := 0

   Next nVez

EndIf

/*************************************************
* Imprime Totalizadores Global dos Gastos Gerias (Or?ado / Realizado)
**************************************************/

nLin := nLin + 2 // Avanca a linha de impressao

If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
    Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
    nLin := 9 //8
Endif

@nLin,000 PSAY "Tt Global de Gastos Gerais Or?ados "
@nLin,042 PSAY nGlbOr1GG  Picture "@E 999,999,999.99"
@nLin,057 PSAY nGlbOr2GG  Picture "@E 999,999,999.99"
@nLin,072 PSAY nGlbOr3GG  Picture "@E 999,999,999.99"
@nLin,087 PSAY nGlbOr4GG  Picture "@E 999,999,999.99"
@nLin,102 PSAY nGlbOr5GG  Picture "@E 999,999,999.99"
@nLin,117 PSAY nGlbOr6GG  Picture "@E 999,999,999.99"
@nLin,132 PSAY nGlbOr7GG  Picture "@E 999,999,999.99"
@nLin,147 PSAY nGlbOr8GG  Picture "@E 999,999,999.99"
@nLin,162 PSAY nGlbOr9GG  Picture "@E 999,999,999.99"
@nLin,177 PSAY nGlbOr10GG Picture "@E 999,999,999.99"
@nLin,192 PSAY nGlbOr11GG Picture "@E 999,999,999.99"
@nLin,207 PSAY nGlbOr12GG Picture "@E 999,999,999.99"

nLin := nLin + 1 // Avanca a linha de impressao

If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
    Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
    nLin := 9 //8
Endif

@nLin,000 PSAY "Tt Global Gastos Gerais Realizados "
@nLin,042 PSAY nGbRea1GG  Picture "@E 999,999,999.99"
@nLin,057 PSAY nGbRea2GG  Picture "@E 999,999,999.99"
@nLin,072 PSAY nGbRea3GG  Picture "@E 999,999,999.99"
@nLin,087 PSAY nGbRea4GG  Picture "@E 999,999,999.99"
@nLin,102 PSAY nGbRea5GG  Picture "@E 999,999,999.99"
@nLin,117 PSAY nGbRea6GG  Picture "@E 999,999,999.99"
@nLin,132 PSAY nGbRea7GG  Picture "@E 999,999,999.99"
@nLin,147 PSAY nGbRea8GG  Picture "@E 999,999,999.99"
@nLin,162 PSAY nGbRea9GG  Picture "@E 999,999,999.99"
@nLin,177 PSAY nGbRea10GG Picture "@E 999,999,999.99"
@nLin,192 PSAY nGbRea11GG Picture "@E 999,999,999.99"
@nLin,207 PSAY nGbRea12GG Picture "@E 999,999,999.99"

nLin := nLin + 1 // Avanca a linha de impressao

If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
    Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
    nLin := 9 //8
Endif

@nLin,000 PSAY "Tt Global de Gastos Gerais A Realizar "
@nLin,042 PSAY nGbaRea1GG  Picture "@E 999,999,999.99"
@nLin,057 PSAY nGbaRea2GG  Picture "@E 999,999,999.99"
@nLin,072 PSAY nGbaRea3GG  Picture "@E 999,999,999.99"
@nLin,087 PSAY nGbaRea4GG  Picture "@E 999,999,999.99"
@nLin,102 PSAY nGbaRea5GG  Picture "@E 999,999,999.99"
@nLin,117 PSAY nGbaRea6GG  Picture "@E 999,999,999.99"
@nLin,132 PSAY nGbaRea7GG  Picture "@E 999,999,999.99"
@nLin,147 PSAY nGbaRea8GG  Picture "@E 999,999,999.99"
@nLin,162 PSAY nGbaRea9GG  Picture "@E 999,999,999.99"
@nLin,177 PSAY nGbaRea10GG Picture "@E 999,999,999.99"
@nLin,192 PSAY nGbaRea11GG Picture "@E 999,999,999.99"
@nLin,207 PSAY nGbaRea12GG Picture "@E 999,999,999.99"


nLin := nLin + 1 // Avanca a linha de impressao

If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
    Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
    nLin := 9 //8
Endif

@nLin,000 PSAY "Tt Global de Gastos Gerais Saldos "
@nLin,042 PSAY nGbSld1GG  Picture "@E 999,999,999.99"
@nLin,057 PSAY nGbSld2GG  Picture "@E 999,999,999.99"
@nLin,072 PSAY nGbSld3GG  Picture "@E 999,999,999.99"
@nLin,087 PSAY nGbSld4GG  Picture "@E 999,999,999.99"
@nLin,102 PSAY nGbSld5GG  Picture "@E 999,999,999.99"
@nLin,117 PSAY nGbSld6GG  Picture "@E 999,999,999.99"
@nLin,132 PSAY nGbSld7GG  Picture "@E 999,999,999.99"
@nLin,147 PSAY nGbSld8GG  Picture "@E 999,999,999.99"
@nLin,162 PSAY nGbSld9GG  Picture "@E 999,999,999.99"
@nLin,177 PSAY nGbSld10GG Picture "@E 999,999,999.99"
@nLin,192 PSAY nGbSld11GG Picture "@E 999,999,999.99"
@nLin,207 PSAY nGbSld12GG Picture "@E 999,999,999.99"

nLin := nLin + 1 // Avanca a linha de impressao

If nLin > 55 // Salto de P?gina. Neste caso o formulario tem 55 linhas...
   Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
   nLin := 9 //8
Endif

@nLin,000 PSAY "Tt Global % Utiliz. Gastos Gerais"
@nLin,042 PSAY IIf(nGlbOr1GG  > 0, ((nGbRea1GG /nGlbOr1GG) *100),0) Picture "@E 999,999,999.99"
@nLin,057 PSAY IIf(nGlbOr2GG  > 0, ((nGbRea2GG /nGlbOr2GG) *100),0) Picture "@E 999,999,999.99"
@nLin,072 PSAY IIf(nGlbOr3GG  > 0, ((nGbRea3GG /nGlbOr3GG) *100),0) Picture "@E 999,999,999.99"
@nLin,087 PSAY IIf(nGlbOr4GG  > 0, ((nGbRea4GG /nGlbOr4GG) *100),0) Picture "@E 999,999,999.99"
@nLin,102 PSAY IIf(nGlbOr5GG  > 0, ((nGbRea5GG /nGlbOr5GG) *100),0) Picture "@E 999,999,999.99"
@nLin,117 PSAY IIf(nGlbOr6GG  > 0, ((nGbRea6GG /nGlbOr6GG) *100),0) Picture "@E 999,999,999.99"
@nLin,132 PSAY IIf(nGlbOr7GG  > 0, ((nGbRea7GG /nGlbOr7GG) *100),0) Picture "@E 999,999,999.99"
@nLin,147 PSAY IIf(nGlbOr8GG  > 0, ((nGbRea8GG /nGlbOr8GG) *100),0) Picture "@E 999,999,999.99"
@nLin,162 PSAY IIf(nGlbOr9GG  > 0, ((nGbRea9GG /nGlbOr9GG) *100),0) Picture "@E 999,999,999.99"
@nLin,177 PSAY IIf(nGlbOr10GG  > 0, ((nGbRea10GG /nGlbOr10GG) *100),0) Picture "@E 999,999,999.99"
@nLin,192 PSAY IIf(nGlbOr11GG  > 0, ((nGbRea11GG /nGlbOr11GG) *100),0) Picture "@E 999,999,999.99"
@nLin,207 PSAY IIf(nGlbOr12GG  > 0, ((nGbRea12GG /nGlbOr12GG) *100),0) Picture "@E 999,999,999.99"
        

//?????????????????????????????????????????????????????????????????????Ŀ
//? Finaliza a execucao do relatorio...                                 ?
//???????????????????????????????????????????????????????????????????????

SET DEVICE TO SCREEN

//?????????????????????????????????????????????????????????????????????Ŀ
//? Se impressao em disco, chama o gerenciador de impressao...          ?
//???????????????????????????????????????????????????????????????????????

If aReturn[5]==1
   dbCommitAll()
   SET PRINTER TO
   OurSpool(wnrel)
Endif

MS_FLUSH()

Return

Static Function AjstSx1()

Local aArea := GetArea()

//PutSx1(cGrupo,cOrdem,cPergunt  ,"","",cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
  PutSx1(cPerg ,"01","Data de Referencia" ,"Data de Referencia","Data de Referencia","mv_ch1","D"  ,8       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR01",""     ,"","",""    ,""           ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"02","Departamento"   ,"Departamento","Departamento","mv_ch2","N"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR02","Todos","","",""    ,"Comercio","","","Industria","","","Administrativo","","","Com.Exterior","","")
  PutSx1(cPerg ,"03","Lista Nat. Diversas","","","mv_ch3","N"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR03","Sim"  ,"","",""    ,"Nao"        ,"","",""         ,"","",""          ,"","","","","")

RestArea(aArea)

Return(.T.)