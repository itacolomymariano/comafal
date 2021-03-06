#INCLUDE "TopConn.Ch"
User Function RGerCam

Private cPerg       := "RGERCO"
Private dDataDe    := CTOD("")
Private dDataAte   := CTOD("")
Private d2AnoMes   := ""
Private d3AnoMes   := ""

   If Pergunte(cPerg,.T.)
   
      //旼컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴?
      //? Grupo de Perguntas - RGERCO                                         ?
      //? MV_PAR01 - Data De                                                  ?
      //? MV_PAR02 - Data At?                                                 ?
      //? MV_PAR03 - Custo - (1.M?dio, 2.Reposi豫o                            ?
      //? MV_PAR04 - Anal?tico - (1.Sim, 2.N?o)                               ?
      //? MV_PAR05 - Depto - (1.Todos,2.Faturamento,3.Ind?stria,4.Financeiro  ?
      //읕컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴?

      dDataDe  := MV_PAR01
      dDataAte := MV_PAR02
      nDeptos  := MV_PAR05

      lComercial := IIF(nDeptos == 1 .Or. nDeptos == 2,.T.,.F.)
      lIndustria := IIF(nDeptos == 1 .Or. nDeptos == 3,.T.,.F.)
      lAdmFin    := IIF(nDeptos == 1 .Or. nDeptos == 4,.T.,.F.)

      titulo       := "Resumo Gerencial COMAFAL Per?odo "+DtoC(dDataDe)+" a "+DtoC(dDataAte)

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
      
      dData02 := CTOD("01/"+Substr(d2AnoMes,5,2)+"/"+Substr(d2AnoMes,1,4))
      dData03 := CTOD("01/"+Substr(d3AnoMes,5,2)+"/"+Substr(d3AnoMes,1,4))
      
      Processa({||RunRE()},"C.A Receber ...")
  
      Alert("Processo Conclu?do com Sucesso !!!")
      
   EndIf

Return
      Static Function RunRE
      ProcRegua(3)
      nCt := 1
      WhIle nCt <= 3
         IncProc("Calc. Saldo A Receber, Proc. "+Alltrim(Str(nCt))+" / 3")
         If nCt == 1
            nSld1MRec := ItaReceb(dDataDe,1,.T.)
         ElseIf nCt == 2
            nSld2MRec := ItaReceb(LastDay(dData02),1,.T.) 
         ElseIf nCt == 3
            nSld3MRec := ItaReceb(LastDay(dData03),1,.T.) 
         EndIf
         nCt ++
      EndDo
      Return


   

//CFOP큦 de Vendas
Private cCFOPVenda  	    := "5101,6101,5102,6102,5103,6103,5104,6104,5105,6105,"
        cCFOPVenda  	    += "5106,6106,6107,6108,5109,6109,5110,6110,5111,6111,5112,6112,"
        cCFOPVenda  	    += "5113,6113,5114,6114,5115,6115,5116,6116,5117,6117,5118,6118,5119,6119,"
        cCFOPVenda  	    += "5120,6120,5122,6122,5123,6123,7101,7102,7105,7106,7127"
        //Ita - Teste em Casa -  cCFOPVenda  	    += "5120,6120,5122,6122,5123,6123,7101,7102,7105,7106,7127,512"

//CFOP큦 de Devolu豫o de Vendas
//Ita - Teste em Casa - Private cCFOPDvVd   := "1201,2201,3201,1202,2202,3202,112,1202" //devolucao de venda
Private cCFOPDvVd   := "1201,2201,3201,1202,2202,3202" //devolucao de venda

//Naturezas - (M.O. + BENEFICIAMENTO + MANUTEN플O + SERVICO TERCEIRIZADO + EPI + INSUMOS + ENERGIA + AGUA + ALUGUEL)
Private cNatProd := "100301,100302,100303,100304,100305,100306,100307,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100409,100410,"
        cNatProd += "100411,100412,100417,200314,600601,200317,200318,200319,200417,200418,200419,200320,200420,200308,200408,200312,200412,200311,200411,200310,200410,"
        cNatProd += "200309,200303,200403"
//Naturezas de Gastos Gerais do Comercial
Private cGGNatComl := "100120,100201,100202,100203,100204,100206,100207,100208,100209,100210,100211,100212,100213,100214,100215,"
        cGGNatComl += "100216,100222,200201,200202,200203,200204,200205,200206,200207,600101,600102,600303,700302"

//Naturezas de Gastos Gerais M?o de Obra
Private cNatGGMO   := "100120,100201,100202,100204,100206,100207,100208,100209,100210,100211,100213,100214,100215,100216,100222"
//Naturezas de Despesas Fixas
Private cDFNatGG   := "100202,100203,100204,100206,100207,100208,100209,100214,100216"
//Naturezas de Despesas Peri?dicas
Private cDPNatGG   := "100120,100201,100212,100213"
//Naturezas de Rescis?es
Private cNResGG    := "100211,100222"
//Naturezas de Premia寤es
Private cNtPremGG  := "100210,100215"
//Naturezas de Marketing
Private cNatMark   := "200204,200207"
//Naturezas co Despesas de Viagens
Private cNtDespV   := "200203"
//Naturezas com Fretes de Vendas
Private cNtFrtVend := "600101"
//Naturezas de Outras Despesas do Comercial
Private cNtOutDesp := "200205"
//Naturezas de Outras Naturezas
Private nOutNatGG  := "100212,200201,200202,200206,600102,600303,700302"
//Naturezas de GASTOS META ORCAMENTO COML
Private cNtGMtOrCml:= "100204,100206,100207,100208,100209,100214,200201,200202,200203,200204,200205,200206,200207,600303"
//NATUREZAS PARA CUSTO DE PRODUCAO
Private cCUSTPROD := "100301,100302,100303,100304,100305,100306,100307,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100409,"
        cCUSTPROD += "100410,100411,100412,100417,200314,600601,200317,200318,200319,200417,200418,200419,200320,200420,200308,200408,200312,200412,200311,200411,200310,"
        cCUSTPROD += "200410,200309,200303,200403"
//GASTOS GERAIS IND.
Private cGASTGERIND := "100301,100302,100303,100304,100305,100306,100307,100308,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100408,100409,100410,100411,100412,"
        cGASTGERIND += "100417,200109,200110,200111,200132,200135,200301,200302,200303,200304,200305,200306,200307,200308,200309,200310,200311,200312,200313,200314,200315,200316,200317,200318,200319,200320,"
        cGASTGERIND += "200321,200322,200360,200361,200362,200363,200364,200365,200366,200367,200368,200369,200370,200372,200373,200374,200375,200376,200377,200401,200402,200403,200404,200405,200406,200407,"
        cGASTGERIND += "200408,200410,200411,200412,200416,200417,200418,200419,200420,200421,600201,600301,600501,600601,700301,700303,700304,700306,700307,700308,700309"
//Natures de Gastos Totais com Ativo Imobilizado
Private cImobTotal := "200109,200110,200111,200306,200406,700309"
//TOTAL DE GASTOS REALIZADO COM M핽UINAS E MOTORES
Private cNatMaqMot := "200111,200306,200406"
//TOTAL DE GASTOS REALIZADO COM IMOBILIZADO IMPORTADO 
Private cNatImpInd := "700309"
//TOTAL DE GASTOS REALIZADO COM IMOBILIZADO DA 핾EA DE INFORM햀ICA 
Private cNatInfInd := "200109"
//TOTAL DE GASTOS REALIZADO COM M?VEIS E UTENS?LIOS 
Private cMovUteInd := "200110"
//TOTAL DE INVESTIMENTOS REALIZADOS NA EMPRESA (SOMAT?RIO DAS NATUREZAS QUE POSSUEM GASTOS COM INVESTIMENTOS)
Private cInvTotInd := "200132,200135,200313,200322,200360,200361,200362,200363,200364,200365,200366,200367,200368,200369,200370,200372,200373,200374,200375,200376,200377"
//INVESTIMENTOS NAS EMPRESAS CABOMAR E COMAFAL-RS
Private cNtEmpresas:= "200135,200360"
//INVESTIMENTOS LAN?ADOS NAS NATUREZAS DE M핽UINAS E EQUIPAMENTOS 
Private cNtMaquinas:= "200361,200362,200363,200364,200369,200370,200375"
//VALOR DE FINAME LAN?ADO NA NATUREZA 
Private cNtFiname := "200313"
//VALORES LAN?ADOS REF. LEASING 
Private cNtLeasing:= "200132"
//SOMAT?RIO DAS NATUREZAS QUE POSSUEM DESPESAS REF. A BENFEITORIAS 
Private cNtBenfeit:= "200322, 200365,200366,200368,200372,200373,200374,200376,200377"
//VALORES LAN?ADOS NA NATUREZA ISO 9001 - 2000  
Private cNtISO92  := "200367"
//SOMAT?RIO DAS NATUREZAS QUE POSSUEM GASTOS RELACIONADOS A MATERIA PRIMA 
Private cNtMPTotal:= "700303,700306,700307,700308,600201"
//MATERIA PRIMA NACIONAL 
Private cNtNacMP := "700303"
//MAT?RIA PRIMA IMPORTADA   
Private cNtImpMP := "700306"
//TRIBUTOS DA IMPORTA플O  
Private cNtTrbImp:= "700307"
//DESPESAS DE IMPORTA플O 
Private cNtDespImp:="700308"
//FRETE IMPORTA플O 
Private cNtFrtImp:= "600201"
//INSUMOS 
Private cNtInsumos := "200308,200408,200312,200412"
//DESPESAS COM MAO DE OBRA
Private cNtMaoObInd:="100301,100302,100303,100304,100305,100306,100307,100309,100310,100311,100312,100313,100317,100401,100402,100403,100404,100405,100406,100407,100409,100410,100411,100412,100417"
//DESPESA FIXA Ind?stria 
Private cNtMODFInd:= "100303,100403,100305,100405,100306,100406,100307,100407,100311,100411,100312,100412"                 
//DESPESA PERI?DICA Ind?stria
Private cNtMODPInd:= "100301,100401,100304,100404,100308,100408,100309,100409"
//RESCIS홒 Ind?stria 
Private cNtRecisInd:= "100310,100410,100317,100417"
//PREMIA합ES Ind?stria 
Private cNtPremInd := "100302,100402,100313"
//BENEFICIAMENTO  Ind?stria 
Private cNtBenefInd:= "200314,600601"
//MANUTEN플O Ind?stria 
Private cNtManuten:= "200317,200318,200319,200417,200418,200419"
//SERVI?OS TERCEIRIZADOS Ind?stria 
Private cNtSerTerInd:= "200303,200403"
//EPI Ind?stria 
Private cNtEPIInd := "200320,200420"
//ENERGIA Ind?stria 
Private cNtEnergInd:= "200311,200411"
//AGUA  Ind?stria 
Private cNtAguaInd := "200310,200410"
//ALUGUEL Ind?stria 
Private cNtAlugInd := "200309"
//CONSERVA플O DE PREDIO Ind?stria 
Private cNtCoPredInd:= "200321,200421"
//DESPESAS DE VIAGENS  Ind?stria 
Private cNtDespViaIn:= "200304,200404"
//OUTRAS DESPESAS DA IND?STRIA  Ind?stria 
Private cNtOutDesInd := "200316,200416"
//OUTRAS NATUREZAS DA IND?STRIA 
Private cNtOutNatInd := "100308,100408,200301,200302,200305,200307,200315,200401,200402,200405,200407,600301,600501,700301,700304"
//TOTAL DE GASTOS REALIZADOS NAS NATUREZAS QUE COMP?E AS METAS DE OR?AMENTO
Private cNtGstMtOrInd := "100303,100304,100305,100306,100307,100311,100312,100403,100404,100405,100406,100407,100411,100412,200301,200302,200303,200304,200305,200308,200309,200310,200311,"
        cNtGstMtOrInd += "200312,200316,200317,200318,200319,200320,200321,200322,200401,200402,200403,200404,200405,200408,200410,200411,200412,200416,200417,200418,200419,200420,200421,"
        cNtGstMtOrInd += "600301"

//Variaveis Privadas de Impress?o
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

//Comiss?o

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

//Metas do Or?amento

nMetasVlr  := 0
nMetasVlr  := 0
nMetasVlr  := 0
      
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

// Variaveis Ind?stria

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

// Gastos Gerais Ind?stria

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
nPerMO1MInd := 0
nPerMO1MInd  := 0
      
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

nMetasVlr  := 0
nMetasVlr  := 0
nMetasVlr  := 0

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

// T?tulos

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
nBx1MDvRE := 0
nBx1MDvRE := 0

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
      
//WW

//Variaveis que guardam Anos para Verificar Titulos Vencidos
Private cAnoAtu    := ""
Private cAno1Menos := ""
Private cAno2Menos := ""
Private cAnoAntes  := ""


//Defini寤es das Metas do Or?amento Comercial
nMetasVlr := 0
If SM0->M0_CODFIL $ "01|06"
	nMetasVlr := (12950)
ElseIf SM0->M0_CODFIL == "02"
	nMetasVlr := (32800)
ElseIf SM0->M0_CODFIL == "03"
	nMetasVlr := (11700)
EndIf	

   //旼컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴?
   //? Processamento para sele豫o dos registros p/ impress?o.              ?
   //읕컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴?
   
   If Pergunte(cPerg,.T.)
   
      //旼컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴?
      //? Grupo de Perguntas - RGERCO                                         ?
      //? MV_PAR01 - Data De                                                  ?
      //? MV_PAR02 - Data At?                                                 ?
      //? MV_PAR03 - Custo - (1.M?dio, 2.Reposi豫o                            ?
      //? MV_PAR04 - Anal?tico - (1.Sim, 2.N?o)                               ?
      //? MV_PAR05 - Depto - (1.Todos,2.Faturamento,3.Ind?stria,4.Financeiro  ?
      //읕컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴?

      dDataDe  := MV_PAR01
      dDataAte := MV_PAR02
      nDeptos  := MV_PAR05

      lComercial := IIF(nDeptos == 1 .Or. nDeptos == 2,.T.,.F.)
      lIndustria := IIF(nDeptos == 1 .Or. nDeptos == 3,.T.,.F.)
      lAdmFin    := IIF(nDeptos == 1 .Or. nDeptos == 4,.T.,.F.)

      titulo       := "Resumo Gerencial COMAFAL Per?odo "+DtoC(dDataDe)+" a "+DtoC(dDataAte)

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
      
      dData02 := CTOD("01/"+Substr(d2AnoMes,5,2)+"/"+Substr(d2AnoMes,1,4))
      dData03 := CTOD("01/"+Substr(d3AnoMes,5,2)+"/"+Substr(d3AnoMes,1,4))
      
      Processa({||RunSelRG()},"Selecionando Informa寤es Gerenciais ...")
      
      Alert("Processo Conclu?do com Sucesso !!!")
      
   EndIf

Return

//旼컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴?
//?                                     F U N ? ? E S   G E N ? R I C A S                                                        ?
//읕컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴?


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
      cNomeMes := "MAR?O"
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
/*
Local aArea     := { Alias() , IndexOrd() , Recno() }
Local aAreaSE2  := { SE2->(IndexOrd()), SE2->(Recno()) }
Local bCondSE2
*/
Local nSaldo    := 0

#IFDEF TOP
	Local cFiltro   := ""
#ENDIF

// 旼컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴커
// ? Testa os parametros vindos do Excel                  ?
// 읕컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴켸
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
/*
dbSelectArea("SE2")
dbSetOrder(3)
bCondSE2  := {|| !Eof() .And. xFilial() == SE2->E2_FILIAL .And.;
	SE2->E2_VENCREA <= dData }
If ( lDtAnterior )
	dbSeek(xFilial(),.T.)
Else
	dbSeek(xFilial()+Dtos(dData))
EndIf
*/
cQryE2 := "SELECT SE2.E2_PREFIXO,SE2.E2_NUM,SE2.E2_PARCELA,SE2.E2_TIPO,SE2.E2_NATUREZ,"
cQryE2 += "SE2.E2_FORNECE,SE2.E2_MOEDA,SE2.E2_LOJA,SE2.E2_VENCREA,SE2.E2_VENCTO,"
cQryE2 += "SE2.E2_SALDO"
cQryE2 += " FROM "+RetSQLName("SE2")+" SE2"
cQryE2 += " WHERE SE2.E2_EMIS1 <= '"+DTOS(dData)+"'"
cQryE2 += " AND SE2.E2_FILIAL = '"+xFilial("SE2")+"'"
cQryE2 += " AND SE2.E2_TIPO NOT IN "+FormatIn(MVPROVIS,"/")
cQryE2 += " AND SE2.E2_TIPO NOT IN "+FormatIn(MVABATIM,"/")
cQryE2 += " AND (( SUBSTRING(SE2.E2_FATURA,1,2) <> '  ' AND SUBSTRING(SE2.E2_FATURA,1,6) = 'NOTFAT') OR "
cQryE2 += "      ( SUBSTRING(SE2.E2_FATURA,1,2) <> '  ' AND SUBSTRING(SE2.E2_FATURA,1,6) <> 'NOTFAT' AND"
cQryE2 += " 	   SE2.E2_DTFATUR > '"+DTOS(dData)+"' ) OR"
cQryE2 += "		   SUBSTRING(SE2.E2_FATURA,1,2) = '  ')"//) )"
cQryE2 += " AND SE2.E2_FLUXO <> 'N'"			
cQryE2 += " AND SE2.D_E_L_E_T_ <> '*'"

TCQuery cQryE2 NEW ALIAS "TSE2"
DbSelectArea("TSE2")
While !Eof()
		If ( !TSE2->E2_TIPO $ MVPAGANT+"/"+MV_CPNEG )
			If ( !lMovSE5 )
				nSaldo += xMoeda(TSE2->E2_SALDO,TSE2->E2_MOEDA,1,dData)
				nSaldo -= SomaAbat(TSE2->E2_PREFIXO,TSE2->E2_NUM,TSE2->E2_PARCELA,"P",1,dData,TSE2->E2_FORNECE)
			Else
				nSaldo += SaldoTit(TSE2->E2_PREFIXO,TSE2->E2_NUM,TSE2->E2_PARCELA,TSE2->E2_TIPO,TSE2->E2_NATUREZ,"P",TSE2->E2_FORNECE,nMoeda,,dData,TSE2->E2_LOJA)
				nSaldo -= SomaAbat(TSE2->E2_PREFIXO,TSE2->E2_NUM,TSE2->E2_PARCELA,"P",1,dData,TSE2->E2_FORNECE)
			EndIf
		Else
			If ( !lMovSE5 .And. !TSE2->E2_TIPO $ MVABATIM )
				nSaldo -= SaldoTit(TSE2->E2_PREFIXO,TSE2->E2_NUM,TSE2->E2_PARCELA,TSE2->E2_TIPO,TSE2->E2_NATUREZ,"P",TSE2->E2_FORNECE,nMoeda,,dData,TSE2->E2_LOJA)
			Else
				nSaldo -= xMoeda(TSE2->E2_SALDO,TSE2->E2_MOEDA,1,dData)
			EndIf
		EndIf
	dbSelectArea("TSE2")
	dbSkip()
EndDo
dbSelectArea("TSE2")
DbCloseArea()
/*
RetIndex("SE2")
dbClearFilter()
dbSetOrder(aAreaSE2[1])
dbGoto(aAreaSE2[2])
dbSelectArea(aArea[1])
dbSetOrder(aArea[2])
dbGoto(aArea[3])
*/
Return(nSaldo)
/*
Static Function RunSldPg
/*
굇쳐컴컴컴컴컵컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴눙?
굇쿛arametros? dData   : Data do Movimento a Receber - Default dDataBase  낢?
굇?          ? nMoeda  : Moeda do Saldo Bancario - Defa 1                 낢?
굇?          ? lDtAnterior: Se .T. Ate a Data,.F. Somente Data - Defa .T. 낢?
굇?          ? lMovSE5 : Se .T. considera o saldo do SE5 - Defa .T.       낢?
굇쳐컴컴컴컴컵컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴눙?
굇? Uso      ? SIGAFIN                                                    낢?
굇읕컴컴컴컴컨컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴袂?
굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇굇?
賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽賽?


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


Static Function ItaReceb(dData,nMoeda,lDtAnterior,lMovSE5)

//Local aArea     := { Alias() , IndexOrd() , Recno() }
//Local aAreaSE1  := { SE1->(IndexOrd()), SE1->(Recno()) }
//Local bCondSE1

Local nSaldo    := 0
#IFDEF TOP
	Local cFiltro   := ""
#ENDIF

// 旼컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴커
// ? Testa os parametros vindos do Excel                  ?
// 읕컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴켸
nMoeda      := If(Empty(nMoeda),1,nMoeda)
dData       := If(Empty(dData),dDataBase,dData)
lDtAnterior := .T. //BoolWindow(lDtAnterior)
lMovSe5     := .T. //BoolWindow(lMovSe5)
If ( ValType(nMoeda) == "C" )
	nMoeda      := Val(nMoeda)
EndIf
//dData       := DataWindow(dData)

// Quando eh chamada do Excel, estas variaveis estao em branco
/*
IF Empty(MVABATIM) .Or.;
	Empty(MV_CRNEG) .Or.;
	Empty(MVRECANT) .Or.;
	Empty(MVPROVIS)
	CriaTipos()
Endif
*/

//dbSelectArea("SE1")
//dbSetOrder(7)
//bCondSE1  := {|| !Eof() .And. xFilial() == SE1->E1_FILIAL .And.;
//	SE1->E1_VENCREA <= dData }
//If ( lDtAnterior )
//	dbSeek(xFilial(),.T.)
//Else
//	dbSeek(xFilial()+Dtos(dData))
//EndIf

cQryE1 := "SELECT SE1.E1_PREFIXO,SE1.E1_NUM,SE1.E1_PARCELA,SE1.E1_TIPO,SE1.E1_NATUREZ,"
cQryE1 += " SE1.E1_CLIENTE,SE1.E1_MOEDA,SE1.E1_VENCREA,SE1.E1_LOJA,SE1.E1_SALDO"
cQryE1 += " FROM "+RetSQLName("SE1")+" SE1"
cQryE1 += " WHERE SE1.E1_FILIAL = '"+xFilial("SE1")+"'"
cQryE1 += " AND SE1.E1_EMISSAO <= '"+DTOS(dData)+"'"
cQryE1 += " AND SE1.E1_TIPO NOT IN "+FormatIn(MVPROVIS,"/")
cQryE1 += " AND SE1.E1_TIPO NOT IN "+FormatIn(MVABATIM,"/")
cQryE1 += " AND ((SUBSTRING(SE1.E1_FATURA,1,2) <> '  ' AND SUBSTRING(SE1.E1_FATURA,1,6) = 'NOTFAT' ) OR"
cQryE1 += " 	 (SUBSTRING(SE1.E1_FATURA,1,2) <> '  ' AND SUBSTRING(SE1.E1_FATURA,1,6) <> 'NOTFAT' AND"
cQryE1 += "		  SE1.E1_DTFATUR > '"+DTOS(dData)+"' ) OR"
cQryE1 += "		  SUBSTRING(SE1.E1_FATURA,1,2) = '  ')"//) )"
cQryE1 += "	AND SE1.E1_FLUXO <> 'N'"
cQryE1 += " AND SE1.D_E_L_E_T_ <> '*'"

TCQuery cQryE1 NEW ALIAS "TSE1"
DbSelectArea("TSE1")
While !Eof()
		If ( !TSE1->E1_TIPO $ MVRECANT+"/"+MV_CRNEG )
			//If ( !lMovSE5 )
			//	nSaldo += xMoeda(TSE1->E1_SALDO,TSE1->E1_MOEDA,1,dData )
			//	nSaldo -= SomaAbat(TSE1->E1_PREFIXO,TSE1->E1_NUM,TSE1->E1_PARCELA,"R",1,dData,TSE1->E1_CLIENTE)
			//Else
				nSaldo += SaldoTit(TSE1->E1_PREFIXO,TSE1->E1_NUM,TSE1->E1_PARCELA,TSE1->E1_TIPO,TSE1->E1_NATUREZ,"R",TSE1->E1_CLIENTE,nMoeda,,dData,TSE1->E1_LOJA)
				nSaldo -= SomaAbat(TSE1->E1_PREFIXO,TSE1->E1_NUM,TSE1->E1_PARCELA,"R",1,dData,TSE1->E1_CLIENTE)
			//EndIf
		Else
			//If ( !lMovSE5 )
			//	nSaldo -= SaldoTit(TSE1->E1_PREFIXO,TSE1->E1_NUM,TSE1->E1_PARCELA,TSE1->E1_TIPO,TSE1->E1_NATUREZ,"R",TSE1->E1_CLIENTE,nMoeda,,dData,TSE1->E1_LOJA)
			//Else
				nSaldo -= xMoeda(TSE1->E1_SALDO,TSE1->E1_MOEDA,1,dData)
			//EndIf
		EndIf
	dbSelectArea("TSE1")
	dbSkip()
EndDo

dbSelectArea("TSE1")
DbCloseArea()


//RetIndex("SE1")
//dbClearFilter()
//dbSetOrder(aAreaSE1[1])
//dbGoto(aAreaSE1[2])
//dbSelectArea(aArea[1])
//dbSetOrder(aArea[2])
//dbGoto(aArea[3])

Return(nSaldo)

Static Function RunSelRG

//旼컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴?
//? 햞ea de Instru寤es SQL - Querys.                                    ?
//읕컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴?

If lComercial .Or. lIndustria

   /**************
   * Calculando Clientes (Comercial)
   * Conceitos COMAFAL
   * Clientes Novos: S?o aqueles cujo a data de cadastro(A1_DTCAD-Customizado) refere-se ao m?s analizado.
   * Clientes Ativos: S?o os clientes novos que nunca compraram, somados aos clientes que compraram nos ?ltimos 
   *                  tr?s(3) meses.
   * Clientes Inativos: S?o aqueles que n?o compraram nos ?ltimos 3 meses.
   * Clientes Atendidos: S?o aqueles que compraram no m?s analizado.
   ***************/
   /*
   ProcRegua(9)
   
   For nMes := 1 To 9   // De 1 a 3(Clientes Novos) - De 4 a 6(Clientes Ativos) - De 7 a 9(Clientes Atendidos)
      
      IncProc("Calculando Clientes, Processo "+Alltrim(Str(nMes))+" / 9 ")
      
      If nMes <= 6
         cQrySD2 := "SELECT COUNT(DISTINCT SA1.A1_COD+SA1.A1_LOJA) AS CLIENTES"
      ElseIf nMes >= 7
         cQrySD2 := "SELECT COUNT(DISTINCT SD2.D2_CLIENTE+SD2.D2_LOJA) AS CLIENTES"
      EndIf
      
      If nMes <= 3
         cQrySD2 += " FROM "+RetSQLName("SA1")+" SA1"
         cQrySD2 += " WHERE SA1.A1_FILIAL = '"+xFilial("SA1")+"'"
      Else
         cQrySD2 += " FROM "+RetSQLName("SD2")+" SD2"
         If nMes >= 4 .And. nMes <= 6
            cQrySD2 += ","+RetSQLName("SA1")+" SA1"
         EndIf
         cQrySD2 += " WHERE SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
         If nMes >= 4 .And. nMes <= 6
            cQrySD2 += "   AND SA1.A1_FILIAL = '"+xFilial("SA1")+"'"
            cQrySD2 += "   AND SD2.D2_CLIENTE = SA1.A1_COD"
            cQrySD2 += "   AND SD2.D2_LOJA = SA1.A1_LOJA"
         EndIf
      EndIf
      
      If nMes <= 3 
         If nMes == 1
            cQrySD2 += "   AND SUBSTRING(SA1.A1_DTCAD,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         ElseIf nMes == 2
            cQrySD2 += "   AND SUBSTRING(SA1.A1_DTCAD,1,6) = '"+d2AnoMes+"'"
         ElseIf nMes == 3
            cQrySD2 += "   AND SUBSTRING(SA1.A1_DTCAD,1,6) = '"+d3AnoMes+"'"
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            cQrySD2 += "   AND SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dDataDe) - 90)+"' AND '"+DTOS(LastDay(dDataDe))+"'"
         ElseIf nMes == 5
            dData02 := CTOD("01/"+Substr(d2AnoMes,5,2)+"/"+Substr(d2AnoMes,1,4))
            cQrySD2 += "   AND SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dData02) - 90)+"' AND '"+DTOS(LastDay(dData02))+"'"
         ElseIf nMes == 6
            dData03 := CTOD("01/"+Substr(d3AnoMes,5,2)+"/"+Substr(d3AnoMes,1,4))
            cQrySD2 += "   AND SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dData03) - 90)+"' AND '"+DTOS(LastDay(dData03))+"'"
         EndIf      
      EndIf
      
      If nMes >= 7
         If nMes == 7
            cQrySD2 += "   AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         ElseIf nMes == 8
            cQrySD2 += "   AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+d2AnoMes+"'"
         ElseIf nMes == 9
            cQrySD2 += "   AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+d3AnoMes+"'"
         EndIf
      EndIf
      
      If nMes <= 3
         cQrySD2 += "   AND SA1.D_E_L_E_T_ <> '*'"
      Else
         //Filtra CFOPs de Vendas - Conforme Solicita豫o de Patr?cia(Auditoria COMAFAL)
         cQrySD2 += "   AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
         cQrySD2 += "   AND SD2.D_E_L_E_T_ <> '*'"
         If nMes >= 4 .And. nMes <= 6
            cQrySD2 += "   AND SA1.D_E_L_E_T_ <> '*'"
         EndIf
      EndIf

      If nMes <= 3
         MemoWrite("C:\TEMP\CliNovos"+Alltrim(Str(nMes))+".SQL",cQrySD2)
      ElseIf nMes >= 4 .And. nMes <= 6
         MemoWrite("C:\TEMP\CliAtivos"+Alltrim(Str(nMes))+".SQL",cQrySD2)
      ElseIf nMes >= 7
         MemoWrite("C:\TEMP\CliAtendidos"+Alltrim(Str(nMes))+".SQL",cQrySD2)
      EndIf
      
      TCQuery cQrySD2 NEW ALIAS "TCLI"
   
      If nMes >= 4 .And. nMes <= 6
         
         cQryCN := "SELECT COUNT(*) AS CLINOV"
         cQryCN += " FROM "+RetSQLName("SA1")+" TSA1"
         If nMes == 4
            cQryCN += " WHERE SUBSTRING(TSA1.A1_DTCAD,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
            cQryCN += " AND (SUBSTRING(TSA1.A1_PRICOM,1,3) = '   ' OR TSA1.A1_PRICOM > '"+DTOS(LastDay(dDataDe))+"')
         ElseIf nMes == 5
            cQryCN += " WHERE SUBSTRING(TSA1.A1_DTCAD,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
            cQryCN += " AND (SUBSTRING(TSA1.A1_PRICOM,1,3) = '   ' OR TSA1.A1_PRICOM > '"+DTOS(LastDay(dData02))+"')
         ElseIf nMes == 6
            cQryCN += " WHERE SUBSTRING(TSA1.A1_DTCAD,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
            cQryCN += " AND (SUBSTRING(TSA1.A1_PRICOM,1,3) = '   ' OR TSA1.A1_PRICOM > '"+DTOS(LastDay(dData03))+"')
         EndIf
         cQryCN += " AND TSA1.D_E_L_E_T_ <> '*'"
      
         TCQuery cQryCN NEW ALIAS "T2CLI"
         
      EndIf
      
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
            DbSelectArea("T2CLI")
            nClNov1Mes := T2CLI->CLINOV
            DbCloseArea()
            DbSelectArea("TCLI")
            nClAtv1Mes := TCLI->CLIENTES + nClNov1Mes
            DbCloseArea()
         EndIf

         If nMes == 5
            DbSelectArea("T2CLI")
            nClNov2Mes := T2CLI->CLINOV
            DbCloseArea()
            DbSelectArea("TCLI")
            nClAtv2Mes := TCLI->CLIENTES + nClNov2Mes
            DbCloseArea()
         EndIf

         If nMes == 6
            DbSelectArea("T2CLI")
            nClNov3Mes := T2CLI->CLINOV
            DbCloseArea()
            DbSelectArea("TCLI")
            nClAtv3Mes := TCLI->CLIENTES + nClNov3Mes
            DbCloseArea()
         EndIf   
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("TCLI")
            nClAtd1Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf

         If nMes == 8
            DbSelectArea("TCLI")
            nClAtd2Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf

         If nMes == 9
            DbSelectArea("TCLI")
            nClAtd3Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf   
      EndIf

   Next nMes
   
   ProcRegua(3)
   
   For nMes := 1 To 3 //Clientes Inativos
      
      IncProc("Calc. Clientes Inativos, Processo "+Alltrim(Str(nMes))+" / 3")
      
      cQrySA1 := "SELECT COUNT(*) AS CLIENTES"
      cQrySA1 += " FROM "+RetSQLName("SA1")+" SA1"
      cQrySA1 += " WHERE SA1.A1_FILIAL = '"+xFilial("SA1")+"'"
      cQrySA1 += " AND (SELECT COUNT(*)"
      cQrySA1 += "      FROM "+RetSQLName("SD2")+" SD2"
      cQrySA1 += "      WHERE SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
      If nMes == 1
         cQrySA1 += "     AND SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dDataDe) - 90)+"' AND '"+DTOS(LastDay(dDataDe))+"'"
      ElseIf nMes == 2
         cQrySA1 += "     AND SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dData02) - 90)+"' AND '"+DTOS(LastDay(dData02))+"'"
      ElseIf nMes == 3
         cQrySA1 += "     AND SD2.D2_EMISSAO BETWEEN '"+DTOS(LastDay(dData03) - 90)+"' AND '"+DTOS(LastDay(dData03))+"'"
      EndIf
      //Filtra CFOPs de Vendas - Conforme Solicita豫o de Patr?cia(Auditoria COMAFAL)
      cQrySA1 += "        AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQrySA1 += "        AND SD2.D2_CLIENTE = SA1.A1_COD"
      cQrySA1 += "        AND SD2.D2_LOJA = SA1.A1_LOJA"
      cQrySA1 += "        AND SD2.D_E_L_E_T_ <> '*') = 0 "
      
      If nMes == 1
         cQrySA1 += " AND SA1.A1_DTCAD <= '"+DTOS(LastDay(dDataDe))+"'"
      ElseIf nMes == 2
         cQrySA1 += " AND SA1.A1_DTCAD <= '"+DTOS(LastDay(dData02))+"'"
      ElseIf nMes == 3
         cQrySA1 += " AND SA1.A1_DTCAD <= '"+DTOS(LastDay(dData03))+"'"
      EndIf
      
      cQrySA1 += " AND SA1.D_E_L_E_T_ <> '*'"
      
      MemoWrite("CliInativos.SQL",cQrySA1)
      
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

   // - **************
   // - * Calculando Pedidos (Quantidade/Valor)
   // - ***************
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. Pedidos, Processo "+Alltrim(Str(nMes))+" / 3")
      
      cQrySC5 := "SELECT COUNT(DISTINCT SC5.C5_NUM) AS PEDIDOS" //Pedidos Emitidos - Quantidade
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
      
      cQrySC6 := "SELECT SUM(C6_VALOR) AS VLRPED" //Pedidos Emitidos - Valor
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


      cQrySD2 := "SELECT COUNT(DISTINCT D2_PEDIDO) AS PEDFAT" //Pedidos Faturados(M?s) - Quantidade
      cQrySD2 += " FROM "+RetSQLName("SD2")+" SD2,"+RetSQLName("SC5")+" SC5,"+RetSQLName("SE4")+" SE4,"+RetSQLName("SF2")+" SF2,"+RetSQLName("SB1")+" SB1"
      cQrySD2 += " WHERE SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
      cQrySD2 += "    AND SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      cQrySD2 += "    AND SE4.E4_FILIAL = '"+xFilial("SE4")+"'"
      cQrySD2 += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"      
      cQrySD2 += "    AND SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
      cQrySD2 += "    AND SD2.D2_PEDIDO = SC5.C5_NUM"
      cQrySD2 += "    AND SF2.F2_COND = SE4.E4_CODIGO"
      If nMes == 1
         cQrySD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         cQrySD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQrySD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
         cQrySD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQrySD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
         cQrySD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      //Filtra CFOPs de Vendas - Conforme Solicita豫o de Patr?cia(Auditoria COMAFAL)
      cQrySD2 += "    AND SF2.F2_DOC = SD2.D2_DOC"
      cQrySD2 += "    AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQrySD2 += "    AND SD2.D2_COD = SB1.B1_COD"
      cQrySD2 += "    AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQrySD2 += "    AND SD2.D_E_L_E_T_ <> '*'"
      cQrySD2 += "    AND SF2.D_E_L_E_T_ <> '*'"
      cQrySD2 += "    AND SE4.D_E_L_E_T_ <> '*'"
      cQrySD2 += "    AND SB1.D_E_L_E_T_ <> '*'"
      cQrySD2 += "    AND SC5.D_E_L_E_T_ <> '*'"
      
      MemoWrite("C:\TEMP\PedFaturados.SQL",cQrySD2)
      
      TCQuery cQrySD2 NEW ALIAS "TSD2"
      
      cQry2SD2 := "SELECT SUM(D2_TOTAL) AS VLRFATPED,SUM(D2_QUANT) AS QTDFATPED"  //Pedidos Faturados(M?s) - Valor
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
         cQry2SD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQry2SD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
         cQry2SD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQry2SD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
         cQry2SD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      //Filtra CFOPs de Vendas - Conforme Solicita豫o de Patr?cia(Auditoria COMAFAL)
      cQry2SD2 += "    AND SF2.F2_DOC = SD2.D2_DOC"
      cQry2SD2 += "    AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQry2SD2 += "    AND SD2.D2_COD = SB1.B1_COD"
      cQry2SD2 += "    AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQry2SD2 += "    AND SD2.D_E_L_E_T_ <> '*'"
      cQry2SD2 += "    AND SF2.D_E_L_E_T_ <> '*'"
      cQry2SD2 += "    AND SE4.D_E_L_E_T_ <> '*'"
      cQry2SD2 += "    AND SB1.D_E_L_E_T_ <> '*'"      
      cQry2SD2 += "    AND SC5.D_E_L_E_T_ <> '*'"      
      
      TCQuery cQry2SD2 NEW ALIAS "T2SD2"

      cQry2SC6 := "SELECT COUNT(DISTINCT SC6.C6_NUM) AS PEDPEND" //Pedidos Pendentes(M?s) - Quantidade
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

      cQry3SC6 := "SELECT SUM(SC6.C6_VALOR) AS VLRPEDPEND" //Pedidos Pendentes(M?s) - Valor
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


      cQryXSD2 := "SELECT COUNT(DISTINCT D2_PEDIDO) AS PEDMAFAT" //Pedidos Faturados - Quantidade(Meses Anteriores)
      cQryXSD2 += " FROM "+RetSQLName("SD2")+" SD2,"+RetSQLName("SC5")+" SC5,"+RetSQLName("SE4")+" SE4,"+RetSQLName("SF2")+" SF2,"+RetSQLName("SB1")+" SB1"
      cQryXSD2 += " WHERE SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
      cQryXSD2 += "    AND SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      cQryXSD2 += "    AND SE4.E4_FILIAL = '"+xFilial("SE4")+"'"
      cQryXSD2 += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"      
      cQryXSD2 += "    AND SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
      cQryXSD2 += "    AND SD2.D2_PEDIDO = SC5.C5_NUM"
      cQryXSD2 += "    AND SF2.F2_COND = SE4.E4_CODIGO"
      If nMes == 1
         cQryXSD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         cQryXSD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQryXSD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
         cQryXSD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQryXSD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
         cQryXSD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      //Filtra CFOPs de Vendas - Conforme Solicita豫o de Patr?cia(Auditoria COMAFAL)
      cQryXSD2 += "    AND SF2.F2_DOC = SD2.D2_DOC"
      cQryXSD2 += "    AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQryXSD2 += "    AND SD2.D2_COD = SB1.B1_COD"
      cQryXSD2 += "    AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQryXSD2 += "    AND SD2.D_E_L_E_T_ <> '*'"
      cQryXSD2 += "    AND SF2.D_E_L_E_T_ <> '*'"
      cQryXSD2 += "    AND SE4.D_E_L_E_T_ <> '*'"
      cQryXSD2 += "    AND SB1.D_E_L_E_T_ <> '*'"      
      
      TCQuery cQryXSD2 NEW ALIAS "TXSD2"
      
      cQry4SD2 := "SELECT SUM(D2_TOTAL) AS VLRMAFATPED,SUM(D2_QUANT) AS QTDMAFTPED"  //Pedidos Faturados - Valor(Meses Anteriores)
      cQry4SD2 += " FROM "+RetSQLName("SD2")+" SD2,"+RetSQLName("SC5")+" SC5,"+RetSQLName("SE4")+" SE4,"+RetSQLName("SF2")+" SF2,"+RetSQLName("SB1")+" SB1"
      cQry4SD2 += " WHERE SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
      cQry4SD2 += "    AND SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      cQry4SD2 += "    AND SE4.E4_FILIAL = '"+xFilial("SE4")+"'"
      cQry4SD2 += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"      
      cQry4SD2 += "    AND SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
      cQry4SD2 += "    AND SD2.D2_PEDIDO = SC5.C5_NUM"
      cQry4SD2 += "    AND SF2.F2_COND = SE4.E4_CODIGO"
      If nMes == 1
         cQry4SD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         cQry4SD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQry4SD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
         cQry4SD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQry4SD2 += " AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
         cQry4SD2 += " AND SUBSTRING(SC5.C5_EMISSAO,1,6) < '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      //Filtra CFOPs de Vendas - Conforme Solicita豫o de Patr?cia(Auditoria COMAFAL)
      cQry4SD2 += "    AND SF2.F2_DOC = SD2.D2_DOC"
      cQry4SD2 += "    AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQry4SD2 += "    AND SD2.D2_COD = SB1.B1_COD"
      cQry4SD2 += "    AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQry4SD2 += "    AND SD2.D_E_L_E_T_ <> '*'"
      cQry4SD2 += "    AND SF2.D_E_L_E_T_ <> '*'"
      cQry4SD2 += "    AND SE4.D_E_L_E_T_ <> '*'"
      cQry4SD2 += "    AND SB1.D_E_L_E_T_ <> '*'"      
      cQry4SD2 += "    AND SC5.D_E_L_E_T_ <> '*'"      
      
      TCQuery cQry4SD2 NEW ALIAS "T4SD2"

      cQry4SC6 := "SELECT COUNT(DISTINCT SC6.C6_NUM) AS PEDMAPEND" //Pedidos Pendentes - Quantidade(Meses Anteriores)
      cQry4SC6 += " FROM "+RetSQLName("SC5")+" SC5,"+RetSQLName("SC6")+" SC6"
      cQry4SC6 += " WHERE SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
      cQry4SC6 += "   AND SC6.C6_FILIAL = '"+xFilial("SC6")+"'"
      cQry4SC6 += "   AND SC5.C5_NUM = SC6.C6_NUM"
      cQry4SC6 += "   AND SC6.C6_CF IN "+FormatIn(cCFOPVenda,",")
      cQry4SC6 += "   AND (SUBSTRING(SC6.C6_NOTA,1,3) = '   ' OR (SUBSTRING(SC6.C6_NOTA,1,3) <> '   ' AND (SELECT SUBSTRING(SF2.F2_EMISSAO,1,6)"
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

      cQry5SC6 := "SELECT SUM(SC6.C6_VALOR) AS VLRPMAPEND" //Pedidos Pendentes - Valor(Meses Anteriores)
      cQry5SC6 += " FROM "+RetSQLName("SC5")+" SC5,"+RetSQLName("SC6")+" SC6"
      cQry5SC6 += " WHERE SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
      cQry5SC6 += "   AND SC6.C6_FILIAL = '"+xFilial("SC6")+"'"
      cQry5SC6 += "   AND SC5.C5_NUM = SC6.C6_NUM"
      cQry5SC6 += "   AND SC6.C6_CF IN "+FormatIn(cCFOPVenda,",")
      cQry5SC6 += "   AND (SUBSTRING(SC6.C6_NOTA,1,3) = '   ' OR (SUBSTRING(SC6.C6_NOTA,1,3) <> '   ' AND (SELECT SUBSTRING(SF2.F2_EMISSAO,1,6)"
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

   ProcRegua(6)
   
   For nMes := 1 To 6  //Faturamento Bruto (Total)
      
      IncProc("Calc. Faturamento Bruto, Processo "+Alltrim(Str(nMes))+" / 6")
      
      If nMes <= 3
         cQry1CFT := " SELECT SUM(SD2.D2_TOTAL) AS TVALBRUTO"
      ElseIf nMes >= 4
         cQry1CFT := " SELECT SUM(SD2.D2_QUANT * SB1.B1_X_CST2) AS CUSFAT"
      EndIf
      cQry1CFT += " FROM "+RetSQLName("SD2")+" SD2,"+RetSQLName("SC5")+" SC5,"+RetSQLName("SE4")+" SE4,"+RetSQLName("SF2")+" SF2,"+RetSQLName("SB1")+" SB1"
      cQry1CFT += "  WHERE SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
      cQry1CFT += "    AND SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      cQry1CFT += "    AND SE4.E4_FILIAL = '"+xFilial("SE4")+"'"
      cQry1CFT += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"      
      cQry1CFT += "    AND SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
      cQry1CFT += "    AND SD2.D2_PEDIDO = SC5.C5_NUM"
      cQry1CFT += "    AND SF2.F2_COND = SE4.E4_CODIGO"
      If nMes == 1 .Or. nMes == 4
         cQry1CFT += "    AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         cQry1CFT += "    AND SUBSTRING(SC5.C5_EMISSAO,1,6) <= '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2 .Or. nMes == 5
         cQry1CFT += "    AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
         cQry1CFT += "    AND SUBSTRING(SC5.C5_EMISSAO,1,6) <= '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3 .Or. nMes == 6
         cQry1CFT += "    AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
         cQry1CFT += "    AND SUBSTRING(SC5.C5_EMISSAO,1,6) <= '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      //Filtra CFOPs de Vendas - Conforme Solicita豫o de Patr?cia(Auditoria COMAFAL)
      cQry1CFT += "    AND SF2.F2_DOC = SD2.D2_DOC"
      cQry1CFT += "    AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQry1CFT += "    AND SD2.D2_COD = SB1.B1_COD"
      cQry1CFT += "    AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQry1CFT += "    AND SD2.D_E_L_E_T_ <> '*'"
      cQry1CFT += "    AND SF2.D_E_L_E_T_ <> '*'"
      cQry1CFT += "    AND SE4.D_E_L_E_T_ <> '*'"
      cQry1CFT += "    AND SB1.D_E_L_E_T_ <> '*'"
      cQry1CFT += "    AND SC5.D_E_L_E_T_ <> '*'"
      
      MemoWrite("FatBrutTot.SQL",cQry1CFT)
      
      TCQuery cQry1CFT NEW ALIAS "CFT1"
      
      If nMes == 1
         DbSelectArea("CFT1")
            nFat1MesBru := CFT1->TVALBRUTO
         DbCloseArea()
      ElseIf nMes == 2
         DbSelectArea("CFT1")
            nFat2MesBru := CFT1->TVALBRUTO
         DbCloseArea()
      ElseIf nMes == 3
         DbSelectArea("CFT1")
            nFat3MesBru := CFT1->TVALBRUTO
         DbCloseArea()
      ElseIf nMes == 4
         DbSelectArea("CFT1")
            nCus1MesFat := CFT1->CUSFAT
         DbCloseArea()
      ElseIf nMes == 5
         DbSelectArea("CFT1")
            nCus2MesFat := CFT1->CUSFAT
         DbCloseArea()
      ElseIf nMes == 6
         DbSelectArea("CFT1")
            nCus3MesFat := CFT1->CUSFAT
         DbCloseArea()
      EndIf
   
   Next nMes
   
   aCFTList := {}
   
   ProcRegua(3)
   
   For nMes := 1 To 3  //Faturamento Bruto por Condi豫o de Faturamento(Pagamento)
      
      IncProc("Calc. Fat.Bruto p/Cond.Pg, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryCFT := " SELECT SF2.F2_COND,SE4.E4_DESCRI,SUM(SD2.D2_TOTAL) AS VALBRUTO"
      cQryCFT += " FROM "+RetSQLName("SD2")+" SD2,"+RetSQLName("SC5")+" SC5,"+RetSQLName("SE4")+" SE4,"+RetSQLName("SF2")+" SF2,"+RetSQLName("SB1")+" SB1"
      cQryCFT += "  WHERE SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
      cQryCFT += "    AND SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      cQryCFT += "    AND SE4.E4_FILIAL = '"+xFilial("SE4")+"'"
      cQryCFT += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"      
      cQryCFT += "    AND SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
      cQryCFT += "    AND SD2.D2_PEDIDO = SC5.C5_NUM"
      cQryCFT += "    AND SF2.F2_COND = SE4.E4_CODIGO"
      If nMes == 1
         cQryCFT += "    AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         cQryCFT += "    AND SUBSTRING(SC5.C5_EMISSAO,1,6) <= '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQryCFT += "    AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
         cQryCFT += "    AND SUBSTRING(SC5.C5_EMISSAO,1,6) <= '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQryCFT += "    AND SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
         cQryCFT += "    AND SUBSTRING(SC5.C5_EMISSAO,1,6) <= '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQryCFT += "    AND SF2.F2_DOC = SD2.D2_DOC"
      cQryCFT += "    AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQryCFT += "    AND SD2.D2_COD = SB1.B1_COD"
      cQryCFT += "    AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQryCFT += "    AND SD2.D_E_L_E_T_ <> '*'"
      cQryCFT += "    AND SF2.D_E_L_E_T_ <> '*'"
      cQryCFT += "    AND SE4.D_E_L_E_T_ <> '*'"
      cQryCFT += "    AND SB1.D_E_L_E_T_ <> '*'"
      cQryCFT += "    AND SC5.D_E_L_E_T_ <> '*'"
      cQryCFT += "GROUP BY SF2.F2_COND,SE4.E4_DESCRI"
      
      MemoWrite("FatBrutCndPg.SQL",cQryCFT)
      
      TCQuery cQryCFT NEW ALIAS "CFT"
      
      If nMes == 1
         TCSetField("CFT","VALBRUTO","N",14,02)
      EndIf
      
      DbSelectArea("CFT")
      While !Eof()
         nPos := aScan(aCFTList, {|x| x[1] == CFT->F2_COND})
         If nPos == 0
            Aadd(aCFTList, {CFT->F2_COND,CFT->E4_DESCRI,CFT->VALBRUTO,0,0})
         Else
            aCFTList[nPos,(nMes+2)] := CFT->VALBRUTO
         EndIf
         DbSelectArea("CFT")
         DbSkip()
      EndDo
      DbSelectArea("CFT")
      DbCloseArea()
      
   Next nMes
   
   // - **************************
   // - Devolu豫o de Clientes
   // -
   // - **************************

   ProcRegua(6)
   
   For nMes := 1 To 6
      
      IncProc("Calc. Devolu寤es de Clientes, Proc. "+Alltrim(Str(nMes))+" / 6")
      
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
   
   // **************************
   // *
   // * Produ豫o
   // *
   // ***************************
   
   //Produ豫o por M?s(Quantidade)
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. Produ豫o, Processo "+Alltrim(Str(nMes))+" / 3")
      
      cQrySD3 := " SELECT SUBSTRING(D3_EMISSAO,1,6) AS EMISSAO, SUM(D3_QUANT) AS PRODUZIDO"
      cQrySD3 += " FROM "+RetSqlName('SD3')
      cQrySD3 += " WHERE D3_FILIAL = '"+xFilial("SD3")+"'  "
      If nMes == 1
         cQrySD3 += " AND SUBSTRING(D3_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQrySD3 += " AND SUBSTRING(D3_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQrySD3 += " AND SUBSTRING(D3_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQrySD3 += " AND D3_CF = 'PR0'"
      cQrySD3 += " AND D3_TM = '300'"
      cQrySD3 += " AND D3_LOCAL = '01'"
      cQrySD3 += " AND D3_GRUPO IN ('30','40','50','52','60','70') "
      cQrySD3 += " AND D_E_L_E_T_ = ' '  "
      cQrySD3 += " GROUP BY SUBSTRING(D3_EMISSAO,1,6) "
      cQrySD3 += " ORDER BY SUBSTRING(D3_EMISSAO,1,6) ASC "
      
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

   // - *************************
   // *
   // * Baixas
   // *
   // *************************
   
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

   // ************************
   // *
   // * COMISS홒
   // *
   // ************************
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Comiss?es, Processo "+Alltrim(Str(nMes))+" / 3")
      
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
   
      cQuery += "  AND SUBSTRING(SE5.E5_NATUREZ,1,6) = '100203'"
      cQuery += "  AND SE5.D_E_L_E_T_ = ' '"
      cQuery += "  AND SE5.E5_TIPODOC IN ('VL','BA')"
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

   // ************************
   // *
   // * GASTOS GERAIS
   // *
   // ************************
   
   ProcRegua(33)
   
   For nMes := 1 To 33 //De 1 a 3 Gastos Gerais(cGGNatComl), 
                       //4 a 6 Gastos Gerais M?o de Obra.(cNatGGMO)
                       //7 a 9 Despesas Fixas
                      //10 a 12 Despesas Peri?dicas
                      //13 a 15 Rescis?es
                      //16 a 18 Premia寤es
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

   ProcRegua(3)
   
   For nMes := 1 To 3 
                      
      IncProc("Calc.Gastos Metas/Orc. Proc. "+Alltrim(Str(nMes))+" / 3")
      
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

   // ****************************
   // *
   // * FATURAMENTO SUCATA
   // *
   // * Obs: Conforme orienta豫o do Sr. Arimat?ia, selecionamentos os produtos Sucata pelo
   // *      Grupo 80.
   // *
   // ****************************

   aTSUCList := {}
   
   ProcRegua(12)
   
   For nMes := 1 To 12
   
      IncProc("Calc. Faturamento Sucata, Proc. "+Alltrim(Str(nMes))+" / 12")
      
      If nMes <= 3
         cQrySUC := " SELECT SUM(SD2.D2_TOTAL) AS VLRSUCAT"
      ElseIf nMes >= 4 .And. nMes <= 6
         cQrySUC := "SELECT SUM(SD2.D2_QUANT * SB1.B1_X_CST2) AS VLRCUSSUC"
      ElseIf nMes >= 7 .And. nMes <= 9
         cQrySUC := "SELECT SUM(SD2.D2_QUANT) AS QTDSUCVEN"
      ElseIf nMes >= 10 .And. nMes <= 12
         cQrySUC := " SELECT SF2.F2_COND,SE4.E4_DESCRI,SUM(SD2.D2_TOTAL) AS VLRSUCAT"
      EndIf
      
      cQrySUC += " FROM "+RetSqlName('SF2') + " SF2,"+RetSQLName("SD2")+" SD2,"+RetSQLName("SB1")+" SB1"
      
      If nMes >= 10 .And. nMes <= 12
         cQrySUC += ","+RetSQLName("SE4")+" SE4"
      EndIf
      
      cQrySUC += " WHERE SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      cQrySUC += " AND SD2.D2_FILIAL = '"+xFilial("SD2")+"'"
      cQrySUC += " AND SB1.B1_FILIAL = '"+xFilial("SD2")+"'"
      
      If nMes >= 10 .And. nMes <= 12
         cQrySUC += " AND SE4.E4_FILIAL = '"+xFilial("SE4")+"'"
      EndIf
      cQrySUC += " AND SF2.F2_DOC = SD2.D2_DOC"
      cQrySUC += " AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQrySUC += " AND SD2.D2_COD = SB1.B1_COD"
      cQrySUC += " AND SUBSTRING(SB1.B1_GRUPO,1,2) = '80'"
      
      If nMes == 1 .Or. nMes == 4 .Or. nMes == 7 .Or. nMes == 10
         cQrySUC += " AND SUBSTRING(SF2.F2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'" 
      ElseIf nMes == 2 .Or. nMes == 5 .Or. nMes == 8 .Or. nMes == 11
         cQrySUC += " AND SUBSTRING(SF2.F2_EMISSAO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'" 
      ElseIf nMes == 3 .Or. nMes == 6 .Or. nMes == 9 .Or. nMes == 12
         cQrySUC += " AND SUBSTRING(SF2.F2_EMISSAO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'" 
      EndIf
      
      cQrySUC += " AND SF2.F2_TIPO = 'N'"
      cQrySUC += " AND SD2.D2_LOCAL IN ('03','04')"
      cQrySUC += " AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQrySUC += " AND SD2.D_E_L_E_T_ <> '*'"
      cQrySUC += " AND SB1.D_E_L_E_T_ <> '*'"
      
      If nMes >= 10 .And. nMes <= 12
         cQrySUC += " AND SE4.D_E_L_E_T_ <> '*'"
         cQrySUC += " GROUP BY SF2.F2_COND,SE4.E4_DESCRI "
         cQrySUC += " ORDER BY SF2.F2_COND,SE4.E4_DESCRI ASC "   
      EndIf
      
      TCQuery cQrySUC NEW ALIAS "TSUC"
      
      If nMes <= 3
         If nMes == 1
            DbSelectArea("TSUC")
            nVl1SUCMFat := TSUC->VLRSUCAT
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("TSUC")
            nVl2SUCMFat := TSUC->VLRSUCAT
            DbCloseArea()
         ElseIf nMes == 3
            DbSelectArea("TSUC")
            nVl3SUCMFat := TSUC->VLRSUCAT
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("TSUC")
            nVlM1CusSUC := TSUC->VLRCUSSUC
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("TSUC")
            nVlM2CusSUC := TSUC->VLRCUSSUC
            DbCloseArea()
         ElseIf nMes == 6
            DbSelectArea("TSUC")
            nVlM3CusSUC := TSUC->VLRCUSSUC
            DbCloseArea()
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("TSUC")
            nQt1MSUCVen := TSUC->QTDSUCVEN
            DbCloseArea()      
         ElseIf nMes == 8
            DbSelectArea("TSUC")
            nQt2MSUCVen := TSUC->QTDSUCVEN
            DbCloseArea()
         ElseIf nMes == 9
            DbSelectArea("TSUC")
            nQt3MSUCVen := TSUC->QTDSUCVEN
            DbCloseArea()
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         DbSelectArea("TSUC")
         While !Eof()
            nPos := aScan(aTSUCList, {|x| x[1] == TSUC->F2_COND})
            If nPos == 0
               Aadd(aTSUCList, {TSUC->F2_COND,TSUC->E4_DESCRI,TSUC->VLRSUCAT,0,0})
            Else
               aTSUCList[nPos,((nMes-9)+2)] := TSUC->VLRSUCAT
            EndIf
            DbSelectArea("TSUC")
            DbSkip()
         EndDo
         DbSelectArea("TSUC")
         DbCloseArea()
      EndIf
      
   Next nMes

EndIf

If lIndustria
   
   // ************************
   // *
   // * Produ豫o por M?quina
   // *
   // ************************
   
   aMaqPrdM := {}
   
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Produ豫o Maq.Ind. Proc. "+Alltrim(Str(nMes))+" / 3")
      
      //cQryPMQ := " SELECT SUBSTRING(SH6.H6_DTAPONT,1,6) AS EMISSAO,SH1.H1_DESCRI AS MAQUINA, SUM(SH6.H6_QTDPROD) AS PRODUZIDO "
      cQryPMQ := " SELECT SH1.H1_DESCRI AS MAQUINA, SUM(SH6.H6_QTDPROD) AS MQPRODUZIDO "
      cQryPMQ += " FROM "
      cQryPMQ += RetSqlName('SH6') + " SH6, "   
      cQryPMQ += RetSqlName('SH1') + " SH1 " 
      cQryPMQ += " WHERE "
      cQryPMQ += " SH6.H6_FILIAL = '"+xFilial("SH6")+"'  "
      
      If nMes == 1
         cQryPMQ += " AND SUBSTRING(SH6.H6_DTAPONT,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"' " 
      ElseIf nMes == 2
         cQryPMQ += " AND SUBSTRING(SH6.H6_DTAPONT,1,6) = '"+Substr(DTOS(dData02),1,6)+"' " 
      ElseIf nMes == 3
         cQryPMQ += " AND SUBSTRING(SH6.H6_DTAPONT,1,6) = '"+Substr(DTOS(dData03),1,6)+"' "
      EndIf
      
      cQryPMQ += " AND SH6.D_E_L_E_T_ <> '*'  "  
      cQryPMQ += " AND SH6.H6_QTDPROD > 0 "
      cQryPMQ += " AND SH6.H6_RECURSO = SH1.H1_CODIGO " 
      cQryPMQ += " AND SH1.H1_FILIAL = '"+xFilial("SH1")+"'  " 
      cQryPMQ += " AND SH1.D_E_L_E_T_ <> '*'"  
      cQryPMQ += " GROUP BY SH1.H1_DESCRI "
      cQryPMQ += " ORDER BY SH1.H1_DESCRI ASC "
      
      TCQuery cQryPMQ NEW ALIAS "TPMQ"
      
      DbSelectArea("TPMQ")
      While !Eof()
         nPosMQ := aScan(aMaqPrdM, { |x| x[1] == TPMQ->H1_DESCRI } )

         If nPosMQ == 0
            Aadd(aMaqPrdM, {TPMQ->H1_DESCRI,TPMQ->MQPRODUZIDO,0,0})
         Else
            aMaqPrdM[nPosMQ,nMes+1]
         EndIf
         
         DbSelectArea("TPMQ")
         DbSkip()
      EndDo
      DbSelectArea("TPMQ")
      DbCloseArea()      
     
   
   Next nMes
   
   // **************************
   // *
   // * Calculando Mat?ria Prima
   // *
   // ***************************
   
   ProcRegua(3)
   
   For nMes := 1 To 3

      IncProc("Calc. Mat.Prima Ind. Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryMP := " SELECT SUM(SB9.B9_QINI) as QUANT, SUM(SB9.B9_QINI*SB9.B9_CUSTD) AS CUSTO "
      cQryMP += " FROM "
      cQryMP += RetSqlName('SB9') + " SB9, " 
      cQryMP += RetSqlName('SB1') + " SB1 "
      cQryMP += " WHERE SB9.D_E_L_E_T_ <> '*' AND "
      cQryMP += " SB9.B9_FILIAL = '"+xFilial("SB9")+"' AND "               
      
      If nMes == 1
         cQryMP += " SUBSTRING(SB9.B9_DATA,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"' AND " 
      ElseIf nMes == 2
         cQryMP += " SUBSTRING(SB9.B9_DATA,1,6) = '"+Substr(DTOS(dData02),1,6)+"' AND " 
      ElseIf nMes == 3
         cQryMP += " SUBSTRING(SB9.B9_DATA,1,6) = '"+Substr(DTOS(dData03),1,6)+"' AND " 
      EndIf

      cQryMP += " SUBSTRING(SB9.B9_COD,1,1) <> 'R' AND "
      cQryMP += " SB9.B9_LOCAL IN ('01') AND "
      cQryMP += " SB1.B1_FILIAL = '"+xFilial("SB1")+"' AND "               
      cQryMP += " SB1.D_E_L_E_T_ <> '*' AND "
      cQryMP += " SB1.B1_COD = SB9.B9_COD AND " 
      cQryMP += " SB1.B1_TIPO = 'MP' "
      //cQryMP += " GROUP BY SUBSTRING(SB9.B9_DATA,1,6) "
      
      TCQuery cQryMP NEW ALIAS "RMP"
      
      If nMes == 1
         DbSelectArea("RMP")
         nQtdFM1MEst := RMP->QUANT
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
   
   // ****************************
   // *
   // * Calculando Produto Acabado
   // *
   // ****************************
   
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Prod.Acab. Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryPA := " SELECT SUM(SB9.B9_QINI) as QUANTPA, SUM(SB9.B9_QINI*SB9.B9_CUSTD) AS CUSTOPA "
      cQryPA += " FROM "
      cQryPA += RetSqlName('SB9') + " SB9,"+RetSQLName("SB1")+" SB1"
      cQryPA += " WHERE SB9.D_E_L_E_T_ <> '*'  "
      cQryPA += " AND SB1.D_E_L_E_T_ <> '*'  "
      cQryPA += " AND SB9.B9_FILIAL = '"+xFilial("SB9")+"'  "               
      cQryPA += " AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'  "
      cQryPA += " AND SB9.B9_COD = SB1.B1_COD"
      If nMes == 1
         cQryPA += " AND SUBSTRING(SB9.B9_DATA,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"' " 
      ElseIf nMes == 2
         cQryPA += " AND SUBSTRING(SB9.B9_DATA,1,6) = '"+Substr(DTOS(dData02),1,6)+"' " 
      ElseIf nMes == 3
         cQryPA += " AND SUBSTRING(SB9.B9_DATA,1,6) = '"+Substr(DTOS(dData03),1,6)+"' " 
      EndIf
      cQryPA += " AND SB9.B9_LOCAL IN ('03','04') "
      cQryPA += " AND SB1.B1_TIPO = 'PA'"
      //cQryPA += " AND SB1.B1_GRUPO = ''"
      
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
   
   // ***************************
   // *
   // * Custo de Produ豫o
   // *
   // ***************************
   
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Custo Produ豫o, Proc. "+Alltrim(Str(nMes))+" / 3")
      
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

   // ***************************
   // *
   // * Gastos Gerais Ind?stria
   // *
   // ***************************
   
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
    
   // /***************************
   // *
   // * Imobilizado Total      
   // *
   // ***************************
   
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

   // ***************************
   // *
   // * M?quinas e Motores     
   // *
   // ***************************
   
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

   // ***************************
   // *
   // * Importado              
   // *
   // ***************************
   
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

   // ***************************
   // *
   // * Informatica            
   // *
   // ***************************
   
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

   // ***************************
   // *
   // * Moveis e Utens?lios
   // *
   // ***************************
   
   
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Mov. Utens?lios, Proc. "+Alltrim(Str(nMes))+" / 3")
      
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

   // ***************************
   // *
   // * INVESTIMENTOS (Total)
   // *
   // ***************************
   
   ProcRegua(21)
   
   For nMes := 1 To 21 //  1 a  3 - Investimentos Total
                      //  4 a  6 - Empresas
                      //  7 a  9 - M?quinas
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

   // ***************************
   // *
   // * MATERIA PRIMA (Total) - Ind?stria
   // *
   // ***************************
   
   ProcRegua(18)
   
   For nMes := 1 To 18 // 1 a  3 - Materia Prima (Total)
                       // 4 a  6 - M. P. Nacional
                      //  7 a  9 - M. P. Importada
                      // 10 a 12 - Tributos Importa豫o
                      // 13 a 15 - Despesas Importa豫o
                      // 16 a 18 - Frete Importa豫o

   
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

   // ***************************
   // *
   // * OUTRAS NATUREZAS - Ind?stria
   // *
   // ***************************
   
   ProcRegua(51)
   
   For nMes := 1 To 51 // 1 a  3 - INSUMOS
                       // 4 a  6 - M홒 DE OBRA
                       // 7 a  9 -  * Desp. Fixa
                       //10 a 12 -  * Desp. Period.
                       //13 a 15 -  * Rescis?es
                       //16 a 18 -  * Premia寤es
                      // 19 a 21 - BENEFICIAMENTO
                      // 22 a 24 - MANUTEN플O
                      // 25 a 27 - SERVI?OS TERCEIRIZADOS
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
   
   // ****************************************
   // *
   // * Gastos Meta Or?amento Ind.
   // *
   // *****************************************
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. Gastos Metas/Orc., Proc. "+Alltrim(Str(nMes))+" / 3")
      
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
   
   // ***********************
   // *
   // * Pedidos na Trava (Total)
   // *
   // ***********************
   
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

   // ***********************
   // *
   // * T?tulos na Trava (Total)
   // *
   // ***********************
   
   ProcRegua(6)
   
   For nMes := 1 To 6
      
      IncProc("Calc. T?tulos na Trava, Proc. "+Alltrim(Str(nMes))+" / 6")
      
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

   // ***********************
   // *
   // * T?tulos a Pagar (M?s)   
   // *
   // ***********************
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. T?tulos A Pagar, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryAP := "SELECT SUM(SE2.E2_VALOR) AS TITAPAG"
      cQryAP += " FROM "+RetSQLName("SE2")+" SE2"
      cQryAP += " WHERE SE2.E2_FILIAL = '"+xFilial("SE2")+"'"

      cQryAP += "  AND SE2.E2_TIPO NOT IN ('NCF','PA ','NDF')"
      cQryAP += "  AND SUBSTRING(SE2.E2_TIPO,3,1) <> '-'"

      If nMes == 1
         cQryAP += " AND SUBSTRING(SE2.E2_VENCTO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQryAP += " AND SUBSTRING(SE2.E2_VENCTO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQryAP += " AND SUBSTRING(SE2.E2_VENCTO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQryAP += " AND SE2.D_E_L_E_T_ <> '*'"
      
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
   
   // ************************
   // *
   // * Pagamentos Realizados
   // *
   // *************************

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

   // ************************
   // *
   // * COMPESAN플O (Baixas por Compensa豫o) - A Pagar
   // *
   // *************************
   
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Bx.p/Compensa豫o, Proc. "+Alltrim(Str(nMes))+" / 3")
      
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

   // ************************
   // *
   // * DEBITO CC 
   // *
   // *************************

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
      
   // ************************
   // *
   // * NORMAL - PAGAR
   // *
   // *************************

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

   // ************************
   // *
   // * DEVOLU플O A PAGAR
   // *
   // *************************
   
   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Bx.p/Devolu豫o, Proc. "+Alltrim(Str(nMes))+" / 3")
      
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

   // ************************
   // *
   // * PAGAMENTOS EM ABERTO
   // *
   // *************************
   */
   //"Calculando Saldo a Pagar ..."
   //RunSldPg()
   
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
   */
   
   /*

   // ***********************
   // *
   // * T?tulos a Pagar (Total)   
   // *
   // ***********************
   
   ProcRegua(12)
   
   For nMes := 1 To 12  // 1 a 3 - A pagar de 1 a 30 dias
                        // 4 a 6 - A pagar de 31 a 60 dias
                        // 7 a 9 - A pagar de 61 a 90 dias
                        //10 a 12 - A pagar a mais de 90 dias
      
      IncProc("Calc. Tit.A Pagar(Fluxo), Proc. "+Alltrim(Str(nMes))+" / 12")
      
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
      
      If nMes <= 3
         If nMes == 1
            DbSelectArea("TTPG")
            nAPg1M1a30 := TTPG->TITAPAG
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("TTPG")
            nAPg2M1a30 := TTPG->TITAPAG
            DbCloseArea()
         ElseIf nMes == 3
            DbSelectArea("TTPG")
            nAPg3M1a30 := TTPG->TITAPAG
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("TTPG")
            nAPg1M31a60 := TTPG->TITAPAG
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("TTPG")
            nAPg2M31a60 := TTPG->TITAPAG
            DbCloseArea()
         ElseIf nMes == 6
            DbSelectArea("TTPG")
            nAPg3M31a60 := TTPG->TITAPAG
            DbCloseArea()
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("TTPG")
            nAPg1M61a90 := TTPG->TITAPAG
            DbCloseArea()
         ElseIf nMes == 8
            DbSelectArea("TTPG")
            nAPg2M61a90 := TTPG->TITAPAG
            DbCloseArea()
         ElseIf nMes == 9
            DbSelectArea("TTPG")
            nAPg3M61a90 := TTPG->TITAPAG
            DbCloseArea()
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10
            DbSelectArea("TTPG")
            nAPg1M90M := TTPG->TITAPAG
            DbCloseArea()
         ElseIf nMes == 11
            DbSelectArea("TTPG")
            nAPg2M90M := TTPG->TITAPAG
            DbCloseArea()
         ElseIf nMes == 12
            DbSelectArea("TTPG")
            nAPg3M90M := TTPG->TITAPAG
            DbCloseArea()
         EndIf
      EndIf
   Next nMes
   
   // ***********************
   // *
   // * T?tulos a Receber (M?s) 
   // *
   // ***********************
   
   ProcRegua(3)
   
   For nMes := 1 To 3
      
      IncProc("Calc. Tit. A Receber, Proc. "+Alltrim(Str(nMes))+" / 3")
      
      cQryAP := "SELECT SUM(SE1.E1_VALOR) AS TITAREC"
      cQryAP += " FROM "+RetSQLName("SE1")+" SE1"
      cQryAP += " WHERE SE1.E1_FILIAL = '"+xFilial("SE1")+"'"
      If nMes == 1
         cQryAP += " AND SUBSTRING(SE1.E1_VENCTO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQryAP += " AND SUBSTRING(SE1.E1_VENCTO,1,6) = '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQryAP += " AND SUBSTRING(SE1.E1_VENCTO,1,6) = '"+Substr(DTOS(dData03),1,6)+"'"
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

   // ************************
   // *
   // * T?tulos Recebidos      
   // *
   // *************************
   
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

   // ************************
   // *
   // * COMPESAN플O (Baixas por Compensa豫o) - A Receber
   // *
   // *************************

   ProcRegua(3)
   
   For nMes := 1 To 3
   
      IncProc("Calc. Bx.p/Compensa豫o(REC), Proc. "+Alltrim(Str(nMes))+" / 3")
      
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

   // ************************
   // *
   // * DEVOLU플O A RECEBER
   // *
   // *************************
   
   ProcRegua(3)

   For nMes := 1 To 3
   
      IncProc("Calc.Bx.p/Devolu豫o(REC), Proc. "+Alltrim(Str(nMes))+" / 3")
      
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

   // ************************
   // *
   // * NORMAL - RECEBER
   // *
   // *************************

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

   // ************************
   // *
   // * TITULOS EM ABERTO (M?s)
   // *
   // *************************

   //"Calculando Saldo a Receber ..."
   //RunSldRe()
   */
   ProcRegua(3)
   nCt := 1
   WhIle nCt <= 2 //3
      IncProc("Calc. Saldo A Receber, Proc. "+Alltrim(Str(nCt))+" / 3")
      If nCt == 1
         nSld1MRec := SldReceber(dDataDe,1,.T.,.T.)
      ElseIf nCt == 2
         nSld2MRec := SldReceber(LastDay(dData02),1,.T.,.T.) 
      ElseIf nCt == 3
        nSld3MRec := SldReceber(LastDay(dData03),1,.T.,.T.) 
      EndIf
      nCt ++
   EndDo
   
   /*
   
   // ******************
   // *
   // * TITULOS VENCIDOS (A RECEBER)
   // *
   // ******************
   
   cAnoAtu    := DTOC(YEAR(dDataDe))
   cAno1Menos := DTOC(YEAR(dDataDe)-1)
   cAno2Menos := DTOC(YEAR(dDataDe)-2)
   cAnoAntes  := DTOC(YEAR(dDataDe)-3)

   ProcRegua(15)
   
   For nMes := 1 To 15
      
      IncProc("Calc. T?t. Vencidos, Proc. "+Alltrim(Str(nMes))+" / 15")
      
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
      
      TCQuery cQryTV NEW ALIAS "TTVENC"
      
      If nMes <= 3
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
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("TTVENC")
            nAAtTt1MVc := TTVENC->SLDVENCID
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("TTVENC")
            nAAtTt2MVc := TTVENC->SLDVENCID
            DbCloseArea()
         ElseIf nMes == 6
            DbSelectArea("TTVENC")
            nAAtTt3MVc := TTVENC->SLDVENCID
            DbCloseArea()
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("TTVENC")
            nA1ATt1MVc := TTVENC->SLDVENCID
            DbCloseArea()
         ElseIf nMes == 8
            DbSelectArea("TTVENC")
            nA1ATt2MVc := TTVENC->SLDVENCID
            DbCloseArea()
         ElseIf nMes == 9
            DbSelectArea("TTVENC")
            nA1ATt3MVc := TTVENC->SLDVENCID
            DbCloseArea()
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10
            DbSelectArea("TTVENC")
            nA2ATt1MVc := TTVENC->SLDVENCID
            DbCloseArea()
         ElseIf nMes == 11
            DbSelectArea("TTVENC")
            nA2ATt2MVc := TTVENC->SLDVENCID
            DbCloseArea()
         ElseIf nMes == 12
            DbSelectArea("TTVENC")
            nA2ATt3MVc := TTVENC->SLDVENCID
            DbCloseArea()
         EndIf
      ElseIf nMes >= 13 .And. nMes <= 15
         If nMes == 13
            DbSelectArea("TTVENC")
            nAAntTt1MVc := TTVENC->SLDVENCID
            DbCloseArea()
         ElseIf nMes == 14
            DbSelectArea("TTVENC")
            nAAntTt2MVc := TTVENC->SLDVENCID
            DbCloseArea()
         ElseIf nMes == 15
            DbSelectArea("TTVENC")
            nAAntTt3MVc := TTVENC->SLDVENCID
            DbCloseArea()
         EndIf
      EndIf
       
   Next nMes 
   
   // ***********************
   // *
   // * T?tulos a Receber (Fluxo)
   // *
   // ***********************
   
   ProcRegua(12)
   
   For nMes := 1 To 12  // 1 a 3 - A Receber de 1 a 30 dias
                        // 4 a 6 - A Receber de 31 a 60 dias
                        // 7 a 9 - A Receber de 61 a 90 dias
                        //10 a 12 - A Receber a mais de 90 dias
      
      IncProc("Calc. Tit. A Receber(Fluxo), Proc. "+Alltrim(Str(nMes))+" / 12")
      
      cQryAP := "SELECT SUM(SE1.E1_VALOR) AS TFLXAREC"
      cQryAP += " FROM "+RetSQLName("SE1")+" SE1"
      cQryAP += " WHERE SE1.E1_FILIAL = '"+xFilial("SE1")+"'"

      cQryAP += "  AND SE1.E1_TIPO NOT IN ('NCC','RA ')"
      cQryAP += "  AND SUBSTRING(SE1.E1_TIPO,3,1) <> '-'"
            
      If nMes <= 3
         If nMes == 1                                                     
            cQryAP += " AND SE1.E1_VENCTO BETWEEN '"+DTOS((LastDay(dDataDe)+1))+"' AND '"+DTOS((LastDay(dDataDe)+30))+"'"
         ElseIf nMes == 2
            cQryAP += " AND SE1.E1_VENCTO BETWEEN '"+DTOS((LastDay(dData02)+1))+"' AND '"+DTOS((LastDay(dData02)+30))+"'"
         ElseIf nMes == 3
            cQryAP += " AND SE1.E1_VENCTO BETWEEN '"+DTOS((LastDay(dData03)+1))+"' AND '"+DTOS((LastDay(dData03)+30))+"'"
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4                                                     
            cQryAP += " AND SE1.E1_VENCTO BETWEEN '"+DTOS((LastDay(dDataDe)+31))+"' AND '"+DTOS((LastDay(dDataDe)+60))+"'"
         ElseIf nMes == 5
            cQryAP += " AND SE1.E1_VENCTO BETWEEN '"+DTOS((LastDay(dData02)+31))+"' AND '"+DTOS((LastDay(dData02)+60))+"'"
         ElseIf nMes == 6
            cQryAP += " AND SE1.E1_VENCTO BETWEEN '"+DTOS((LastDay(dData03)+31))+"' AND '"+DTOS((LastDay(dData03)+60))+"'"
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7                                                     
            cQryAP += " AND SE1.E1_VENCTO BETWEEN '"+DTOS((LastDay(dDataDe)+61))+"' AND '"+DTOS((LastDay(dDataDe)+90))+"'"
         ElseIf nMes == 8
            cQryAP += " AND SE1.E1_VENCTO BETWEEN '"+DTOS((LastDay(dData02)+61))+"' AND '"+DTOS((LastDay(dData02)+90))+"'"
         ElseIf nMes == 9
            cQryAP += " AND SE1.E1_VENCTO BETWEEN '"+DTOS((LastDay(dData03)+61))+"' AND '"+DTOS((LastDay(dData03)+90))+"'"
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10                                                     
            cQryAP += " AND SE1.E1_VENCTO > '"+DTOS((LastDay(dDataDe)+90))+"'"
         ElseIf nMes == 11
            cQryAP += " AND SE1.E1_VENCTO > '"+DTOS((LastDay(dData02)+90))+"'"
         ElseIf nMes == 12
            cQryAP += " AND SE1.E1_VENCTO > '"+DTOS((LastDay(dData03)+90))+"'"
         EndIf
      EndIf
      
      cQryAP += " AND SE1.D_E_L_E_T_ <> '*'"
      
      TCQuery cQryAP NEW ALIAS "REFLX"
      
      If nMes <= 3
         If nMes == 1
            DbSelectArea("REFLX")
            nARe1M1a30 := REFLX->TFLXAREC
            DbCloseArea()
         ElseIf nMes == 2
            DbSelectArea("REFLX")
            nARe2M1a30 := REFLX->TFLXAREC
            DbCloseArea()
         ElseIf nMes == 3
            DbSelectArea("REFLX")
            nARe3M1a30 := REFLX->TFLXAREC
            DbCloseArea()
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("REFLX")
            nARe1M31a60 := REFLX->TFLXAREC
            DbCloseArea()
         ElseIf nMes == 5
            DbSelectArea("REFLX")
            nARe2M31a60 := REFLX->TFLXAREC
            DbCloseArea()
         ElseIf nMes == 6
            DbSelectArea("REFLX")
            nARe3M31a60 := REFLX->TFLXAREC
            DbCloseArea()
         EndIf
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("REFLX")
            nARe1M61a90 := REFLX->TFLXAREC
            DbCloseArea()
         ElseIf nMes == 8
            DbSelectArea("REFLX")
            nARe2M61a90 := REFLX->TFLXAREC
            DbCloseArea()
         ElseIf nMes == 9
            DbSelectArea("REFLX")
            nARe3M61a90 := REFLX->TFLXAREC
            DbCloseArea()
         EndIf
      ElseIf nMes >= 10 .And. nMes <= 12
         If nMes == 10
            DbSelectArea("REFLX")
            nARe1M90M := REFLX->TFLXAREC
            DbCloseArea()
         ElseIf nMes == 11
            DbSelectArea("REFLX")
            nARe2M90M := REFLX->TFLXAREC
            DbCloseArea()
         ElseIf nMes == 12
            DbSelectArea("REFLX")
            nARe3M90M := REFLX->TFLXAREC
            DbCloseArea()
         EndIf
      EndIf
   Next nMes
   */
EndIf

Return
