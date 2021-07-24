#include "TopConn.Ch"
#include "Protheus.Ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³CRVAABC   ºAutor  ³Five Solutions      º Data ³  14/07/2010 º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Relação de Produtos por Percentual de Participação por     º±±
±±º          ³ Curva ABC ou Ranking de vendas.                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ COMAFAL - PE,RS e SP.                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function CrvaABC
   
   Private cPerg       := "RKGABC"
   Private nFatTotal   := 0
   Private nTonTotal   := 0
   
   Private cCFOPVenda  := "5101,6101,5102,6102,5103,6103,5104,6104,5105,6105,"
           cCFOPVenda  += "5106,6106,6107,6108,5109,6109,5110,6110,5111,6111,5112,6112,"
           cCFOPVenda  += "5113,6113,5114,6114,5115,6115,5116,6116,5117,6117,5118,6118,5119,6119,"
           cCFOPVenda  += "5120,6120,5122,6122,5123,6123"//,7101,7102,7105,7106,7127"
           
   MontPerg()
    
   If Pergunte(cPerg,.T.)
   
      //Cria arquivo temporario para listagem da Curva ou Ranking
      Private aCampos := {}
      Private dDatIni := MV_PAR01   //Data Inicial
      Private dDatFin := MV_PAR02   //Data Final
      Private nModRel := MV_PAR03   //Modelo do Relatorio - 1 = Curva ABC, 2=Ranking
      Private nAPPart := MV_PAR04   //Percentual de Participação da Curva A
      Private nBPPart := MV_PAR05   //Percentual de Participação da Curva B
      Private nCPPart := MV_PAR06   //Percentual de Participação da Curva C
      Private cGrpIni := MV_PAR07   //Grupo de Produtos Inicial
      Private cGrpFin := MV_PAR08   //Grupo de Produtos Final
      Private nQtdLst := MV_PAR09   //Quantidade de Produtos a Listar
      Private nOrdRkg := MV_PAR10   //Tipo de Ordem do Ranking 1=Quantidade, 2=Valor, 3=Margem.
      Private nTpOrde := MV_PAR11   //Tipo de Ordem do Indice de Apresentação - 1=Crescente,2=Decrescente
      Private nTpCusto:= MV_PAR12   //Tipo de Custo do Material - 1=Standar;2=Médio
      //Private nQtdMes := (Month(dDatFin) - Month(dDatIni)) //Quantidade de Meses do Periodo
      
      If nQtdLst <= 0
         Alert("Favor informar a quantidade de produtos a listar")
         Return
      EndIf
      
      If nModRel == 1
         If (nAPPart + nBPPart + nCPPart) <> 100
            Alert("Os percentuais da curva A,B,C não compõe 100%, favor redefinir percentuais")
            Return
         EndIf
      EndIf
      
      aAdd(aCampos, {"CODPROD","C",15,00})
      aAdd(aCampos, {"DESCPRO","C",60,00})
      aAdd(aCampos, {"QUANCLI","N",10,00})
      aAdd(aCampos, {"QTDTONE","N",17,04})
      aAdd(aCampos, {"CUSTOPR","N",17,02})
      aAdd(aCampos, {"VLRVEND","N",17,02})
      aAdd(aCampos, {"NMARGEM","N",17,02})
      aAdd(aCampos, {"TPCURVA","C",01,00})
      aAdd(aCampos, {"PPARTIC","N",05,02})
      //aAdd(aCampos, {"MEDIAVE","N",17,02})
      
      Processa({||RSelProd()},"Calculando Vendas, Aguarde ...")
      
   EndIf
   
Return

Static Function RSelProd
   
   //Execuções da Query
   // 1- Seleciona Quantidade de Registros a serem listados
   // 2- Seleciona o Faturamento Total
   // 3- Seleciona o Faturamento por Produto
   
   For nSel := 1 To 3
      
      If nSel == 1
         cQrySEL := " SELECT COUNT(*) nTOTREG "
      Else
         
         cQrySEL := " SELECT "
         
         If nSel == 3
            //cQrySEL += "    TOP "+Alltrim(Str(nQtdLst))
            cQrySEL += "    SD2.D2_COD,SB1.B1_DESC,"
            cQrySEL += "    ( SELECT COUNT(DISTINCT TMD2.D2_CLIENTE+D2_LOJA) "
            cQrySEL += "        FROM "+RetSQLName("SD2")+" TMD2 "
            cQrySEL += "       WHERE TMD2.D2_FILIAL = '"+xFilial("SD2")+"'"
            cQrySEL += "         AND TMD2.D2_COD = SD2.D2_COD"
            cQrySEL += "         AND TMD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
            cQrySEL += "         AND TMD2.D2_TIPO <> 'I'"
            cQrySEL += "         AND TMD2.D2_EMISSAO BETWEEN '"+DTOS(dDatIni)+"' AND '"+DTOS(dDatFin)+"'" 
            cQrySEL += "         AND TMD2.D2_GRUPO BETWEEN '"+cGrpIni+"' AND '"+cGrpFin+"'"
            cQrySEL += "         AND TMD2.D2_CLIENTE <> '006629'"
            cQrySEL += "         AND TMD2.D_E_L_E_T_ <> '*' ) nQTDCLI," 
         EndIf
         
         cQrySEL += "        SUM(D2_TOTAL-D2_VALICM-D2_VALIPI-D2_VALIMP5-D2_VALIMP6-D2_ICMSRET-D2_VALFRE-D2_SEGURO-D2_DESPESA) AS TOTLIQFAT,"
         If nTpCusto == 2  //Custo Médio
            cQrySEL += "        SUM(D2_CUSTO1) AS CUSTOMP,"
         EndIf
         cQrySEL += "        SUM(D2_QUANT)  AS TONVEND"
         //cQrySEL += "        (SUM(D2_QUANT)/"+Alltrim(Str(nQtdMes))+")  AS MEDVEND,"
         
         If nTpCusto == 2  //Custo Médio
            cQrySEL += "        ,((SUM(D2_TOTAL-D2_VALICM-D2_VALIPI-D2_VALIMP5-D2_VALIMP6-D2_ICMSRET-D2_VALFRE-D2_SEGURO-D2_DESPESA) - SUM(D2_CUSTO1)) * 100 )"
            cQrySEL += "         / SUM(D2_CUSTO1) nMARGEM"
         EndIf
      
      EndIf
      
      cQrySEL += " FROM "+RetSqlName('SD2') + " SD2,"+RetSQLName("SB1")+" SB1,"+RetSQLName("SF2")+" SF2 "
      cQrySEL += " WHERE SD2.D2_FILIAL  = '"+xFilial("SD2")+"' AND  SD2.D_E_L_E_T_ = ' ' "
      cQrySEL += " AND SB1.B1_FILIAL  = '"+xFilial("SB1")+"' AND  SB1.D_E_L_E_T_ = ' ' "
      cQrySEL += " AND SF2.F2_FILIAL  = '"+xFilial("SF2")+"' AND  SF2.D_E_L_E_T_ = ' ' "
      cQrySEL += " AND SF2.F2_CLIENTE <> '006629'" //Five Solutions Consultoria - 04/11/2009
      cQrySEL += " AND SD2.D2_COD = SB1.B1_COD" 
      cQrySEL += " AND SUBSTRING(SB1.B1_GRUPO,1,2) <> '40'" //Faturamento sem Sucatas
      cQrySEL += " AND SF2.F2_DOC = SD2.D2_DOC"
      cQrySEL += " AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQrySEL += " AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQrySEL += " AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      cQrySEL += " AND SD2.D2_EMISSAO BETWEEN '"+DTOS(dDatIni)+"' AND '"+DTOS(dDatFin)+"'" 
      cQrySEL += " AND SB1.B1_GRUPO BETWEEN '"+cGrpIni+"' AND '"+cGrpFin+"'"

      cQrySEL += " GROUP BY SD2.D2_COD,SB1.B1_DESC" 
      
      cTipoOrd := If(nTpOrde == 2,"D","")         
      If nSel == 3
         If nModRel == 1
            cQrySEL += " ORDER BY TONVEND DESC" //+If(nTpOrde == 2,"DESC","")         
         Else
            If nOrdRkg == 1 //Quantidade
               cQrySEL += " ORDER BY TONVEND DESC" //+If(nTpOrde == 2,"DESC","")
            ElseIf nOrdRkg == 2 //Valor
               cQrySEL += " ORDER BY TOTLIQFAT DESC" //+If(nTpOrde == 2,"DESC","")
            ElseIf nOrdRkg == 3 //Margem
               cQrySEL += " ORDER BY nMARGEM DESC" //+If(nTpOrde == 2,"DESC","")
            EndIf      
         EndIf
      EndIf
       
      MemoWrite("C:\TEMP\CrvaABC_"+Alltrim(Str(nSel))+".SQL",cQrySEL)
      
      TCQuery cQrySEL NEW ALIAS "CABC"
      
      If nSel == 1
         DbSelectArea("CABC")
         nReg := CABC->nTOTREG
         DbCloseArea()
      EndIf
      
      If nSel == 2
         TCSetField("CABC","N","TONVEND",17,3)
         TCSetField("CABC","N","TOTLIQFAT",17,2)
         DbSelectArea("CABC")
         nFatTotal   := CABC->TOTLIQFAT
         nTonTotal   := CABC->TONVEND
         DbCloseArea()
      EndIf

   Next nSel
   
   TCSetField("CABC","N","TONVEND",17,3)
   TCSetField("CABC","N","NMARGEM",17,2)
   TCSetField("CABC","N","nQTDCLI",10,0)
   TCSetField("CABC","N","TOTLIQFAT",17,2)
   TCSetField("CABC","N","CUSTOMP",17,2)

   cArqTrb := CriaTrab(aCampos)
   DbUseArea(.T.,"DBFCDX",cArqTrb,"TRLST")
   _cIndice := ""
   If nModRel == 1  //Curva ABC
      _cIndice := "QTDTONE"
   ElseIf nModRel == 2  //Ranking de Produtos
      If nOrdRkg == 1 //Quantidade
         _cIndice := "QTDTONE"
      ElseIf nOrdRkg == 2 //Valor
         _cIndice := "VLRVEND"
      ElseIf nOrdRkg == 3 //Margem
         _cIndice := "NMARGEM"
      EndIf
   EndIf

   //IndRegua("TRLST",cArqTrb,_cIndice,cTipoOrd,,"Selecionando Registros...")
   IndRegua("TRLST",cArqTrb,_cIndice,"D",,"Selecionando Registros...")
   
   ProcRegua(nQtdLst) //nReg)
   nCont := 1

   nFatTotal   := 0
   nTonTotal   := 0
   
   DbSelectArea("CABC")
   While CABC->(!Eof()) .And. nCont <= nQtdLst
      
      IncProc("Selecionando Vendas ... "+Alltrim(Str(nCont))+" / "+Alltrim(Str(nReg)))
      
      DbSelectArea("TRLST")
      RecLock("TRLST",.T.)
         TRLST->CODPROD := CABC->D2_COD
         TRLST->DESCPRO := CABC->B1_DESC
         TRLST->QUANCLI := CABC->nQTDCLI
         TRLST->QTDTONE := CABC->TONVEND
/*
            cQrySEL += "        ((SUM(D2_TOTAL-D2_VALICM-D2_VALIPI-D2_VALIMP5-D2_VALIMP6-D2_ICMSRET-D2_VALFRE-D2_SEGURO-D2_DESPESA) - SUM(D2_CUSTO1)) * 100 )"
            cQrySEL += "         / SUM(D2_CUSTO1) nMARGEM"
*/
         nStandBZ := Posicione("SBZ",1,xFilial("SBZ")+CABC->D2_COD,"BZ_CUSTD")
         nCustPrd := CABC->TONVEND * nStandBZ
         //MsgInfo("CABC->D2_COD "+CABC->D2_COD+" nCustPrd "+Transform(nCustPrd,"@E 999,999,999.99")+" CABC->TONVEND "+Transform(CABC->TONVEND,"@E 999,999,999.9999")+" nStandBZ "+tRANSFORM(nStandBZ,"@e 999,999,999.99"))
         If nTpCusto == 1 //Custo Standard
            TRLST->CUSTOPR := nCustPrd
            //Margem de Lucro = (Lucro(Venda-Custo) * 100 / Custo)
            If CABC->TOTLIQFAT > nCustPrd
               TRLST->NMARGEM := ((CABC->TOTLIQFAT - nCustPrd) * 100) / 100
            Else
               TRLST->NMARGEM := (CABC->TOTLIQFAT - nCustPrd) / 100
            EndIf
         Else
            TRLST->CUSTOPR := CABC->CUSTOMP
            TRLST->NMARGEM := CABC->nMARGEM
         EndIf
         TRLST->VLRVEND := CABC->TOTLIQFAT


         
         //TRLST->TPCURVA := CABC->
         //TRLST->PPARTIC := (CABC->TONVEND / nFatTotal)
      MsUnLock()
      //Calcula os Totais do Range de Produtos selecionados.
      nFatTotal   += CABC->TOTLIQFAT
      nTonTotal   += CABC->TONVEND

      DbSelectArea("CABC")
      DbSkip()
      nCont ++
      
   EndDo
   DbSelectArea("CABC")
   DbCloseArea()
   
   If nModRel == 1 //Curva ABC

      nCAQtd := (nTonTotal * (nAPPart/100))
      nCBQtd := (nTonTotal * (nBPPart/100))
      nCCQtd := (nTonTotal * (nCPPart/100))

      //MsgInfo("nCAQtd "+Alltrim(Str(nCAQtd))+" nCBQtd "+Alltrim(Str(nCBQtd))+" nCCQtd "+Alltrim(Str(nCCQtd)))      

      aCurvaA := {}
      aCurvaB := {}
      aCurvaC := {}
      nAcuCrv := 0
      lFazCA  := .T.
      lFazCB  := .F. 
      lFazCC  := .T.
      DbSelectArea("TRLST")
      nReg  := RecCount()
      nCont := 1
      DbGoTop()
      DbSelectArea("TRLST")
      INDEX ON DESCEND(TRLST->QTDTONE) TO TRLSTIND
      While TRLST->(!Eof())
         
         IncProc("Classificando Curvas ... "+Alltrim(Str(nCont))+" / "+Alltrim(Str(nReg)))
         
         nAcuCrv += TRLST->QTDTONE
                
         //MsgInfo("nAcuCrv "+Transform(nAcuCrv,"@E 999,999,999.999")+" nCAQtd "+Transform(nCAQtd,"@E 999,999,999.999")+" nCBQtd "+Transform(nCBQtd,"@E 999,999,999.999")+" nCCQtd "+Transform(nCCQtd,"@E 999,999,999.999")+" nTonTotal "+Transform(nTonTotal,"@E 999,999,999.999"))
         
         If nAcuCrv <= nCAQtd .And. lFazCA
            //MsgInfo("Alimentei Curva A")
            aAdd(aCurvaA,{TRLST->CODPROD,TRLST->DESCPRO,TRLST->QUANCLI,TRLST->QTDTONE,TRLST->CUSTOPR,TRLST->VLRVEND,TRLST->NMARGEM})
         Else
            
            If lFazCA
               //Alert("Saindo da Curva A")
               nAcuCrv := 0
            EndIf
            lFazCA  := .F.
            lFazCB  := .T.
            
         EndIf

         If nAcuCrv <= nCBQtd .And. lFazCB
            //MsgInfo("Alimentei Curva B")
            aAdd(aCurvaB,{TRLST->CODPROD,TRLST->DESCPRO,TRLST->QUANCLI,TRLST->QTDTONE,TRLST->CUSTOPR,TRLST->VLRVEND,TRLST->NMARGEM})
         Else
           
            If lFazCB
               //Alert("Saindo da Curva B")
               nAcuCrv := 0
            EndIf
            lFazCB  := .F.
         EndIf

         If nAcuCrv <= nCCQtd .And. lFazCC .And. (!lFazCA .And. !lFazCB)
            //MsgInfo("Alimentei Curva C")
            aAdd(aCurvaC,{TRLST->CODPROD,TRLST->DESCPRO,TRLST->QUANCLI,TRLST->QTDTONE,TRLST->CUSTOPR,TRLST->VLRVEND,TRLST->NMARGEM})
         Else
            If (!lFazCA .And. !lFazCB)
               //Alert("Saindo da Curva C")
               lFazCC := .F.
               //nAcuCrv := 0
            EndIf
         EndIf
                           
         DbSelectArea("TRLST")
         DbSkip()
         nCont ++
         
      EndDo
      DbSelectArea("TRLST")
      DbCloseArea()
      
      If nTpOrde == 1
        aSort(aCurvaA,,, {|x,y| x[4] < y[4]}) //Ordena Ascendente por Qtd. Tonelada Vendida
        aSort(aCurvaB,,, {|x,y| x[4] < y[4]}) 
        aSort(aCurvaC,,, {|x,y| x[4] < y[4]}) 
      Else
        aSort(aCurvaA,,, {|x,y| x[4] > y[4]}) //Ordena Descendente por Qtd. Tonelada Vendida
        aSort(aCurvaB,,, {|x,y| x[4] > y[4]}) 
        aSort(aCurvaC,,, {|x,y| x[4] > y[4]}) 
      EndIf
      
      RABCImp(1) //Imprime Relatório - Curva ABC
      
   Else
   
      RABCImp(2) //Imprime Relatório - Ranking de Produtos
      
   EndIf
   
Return

Static Function MontPerg

   Local aArea := GetArea()

   //PutSx1(cGrupo,cOrdem,cPergunt  ,"","",cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
 
   Local aHelpPor := {}
   Local aHelpEng := {}
   Local aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Data Inicial do Periodo do')
   Aadd( aHelpPor, 'Relatorio                           ')
                                                                                                           //F3
   PutSx1(cPerg ,"01","Da Data"  ,"Da Data","Da Data","mv_ch1","D"  ,08       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR01","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Data Final do Periodo do')
   Aadd( aHelpPor, 'Relatorio                           ')
                                                                                                           //F3
   PutSx1(cPerg ,"02","Ate Data"  ,"Ate Data","Ate Data","mv_ch2","D"  ,08       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR02","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe qual o Modelo do Relatorio')

                                                                                                           //F3
 //PutSx1(cPerg ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPerg ,"03","Modelo"  ,"Modelo","Modelo","mv_ch3","N"  ,01       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR03","Curva ABC","","",""    ,"Ranking"  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe qual o percentual de ')
   Aadd( aHelpPor, 'participação da Curva A')

                                                                                                           //F3
 //PutSx1(cPerg ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPerg ,"04","(%) Curva A"  ,"(%) Curva A","(%) Curva A","mv_ch4","N"  ,5       ,2       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR04","","","",""    ,""  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   Aadd( aHelpPor, 'Informe qual o percentual de ')
   Aadd( aHelpPor, 'participação da Curva B')

                                                                                                           //F3
 //PutSx1(cPerg ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPerg ,"05","(%) Curva B"  ,"(%) Curva B","(%) Curva B","mv_ch5","N"  ,5       ,2       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR05","","","",""    ,""  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   Aadd( aHelpPor, 'Informe qual o percentual de ')
   Aadd( aHelpPor, 'participação da Curva C')

                                                                                                           //F3
 //PutSx1(cPerg ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPerg ,"06","(%) Curva C"  ,"(%) Curva C","(%) Curva C","mv_ch6","N"  ,5       ,2       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR06","","","",""    ,""  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
      

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe o Grupo de Produtos Inicial')
   Aadd( aHelpPor, 'da Faixa de Grupos do Relatorio.   ')
 //PutSx1(cGrupo,cOrdem,cPergunt       ,""           ,""           ,cVar    ,cTipo,nTamanho ,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
   PutSx1(cPerg ,"07"  ,"Do Grupo"  ,"Do Grupo","Do Grupo","mv_ch7","C"  ,4       ,0       ,0      ,"G" ,""    ,"SBM" ,""     ,"","MV_PAR07","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe o Grupo de Produtos Final')
   Aadd( aHelpPor, 'da Faixa de Grupos do Relatorio.   ')
 //PutSx1(cGrupo,cOrdem,cPergunt       ,""           ,""           ,cVar    ,cTipo,nTamanho ,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
   PutSx1(cPerg ,"08"  ,"Até Grupo"  ,"Até Grupo","Até Grupo","mv_ch8","C"  ,4       ,0       ,0      ,"G" ,""    ,"SBM" ,""     ,"","MV_PAR08","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   Aadd( aHelpPor, 'Informe a Qtd. de Produtos a serem')
   Aadd( aHelpPor, 'listados.   ')
 //PutSx1(cGrupo,cOrdem,cPergunt       ,""           ,""           ,cVar    ,cTipo,nTamanho ,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
   PutSx1(cPerg ,"09"  ,"Qtd.Produtos"  ,"Qtd.Produtos","Qtd.Produtos","mv_ch9","N"  ,5       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR09","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)


   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Ordem de impressão do   ')
   Aadd( aHelpPor, 'Ranking de Produtos.              ')

                                                                                                           //F3
 //PutSx1(cPerg ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPerg ,"10","Ordem Ranking"  ,"Ordem Ranking","Ordem Ranking","mv_cha","N"  ,01       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR10","Quantidade","","",""    ,"Valor"  ,"","","Margem","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe o Tipo de Ordem de impressão')
   Aadd( aHelpPor, 'se Crescente ou Decrescente.        ')

                                                                                                           //F3
 //PutSx1(cPerg ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPerg ,"11","Tipo de Ordem"  ,"Tipo de Ordem","Tipo de Ordem","mv_chb","N"  ,01       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR11","Crescente","","",""    ,"Decrescente"  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   Aadd( aHelpPor, 'Informe o Tipo de Custo do Material ')
   Aadd( aHelpPor, 'a ser apresentado no relatorio.     ')

                                                                                                           //F3
 //PutSx1(cPerg ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPerg ,"12","Tipo de Custo"  ,"Tipo de Custo","Tipo de Custo","mv_chc","N"  ,01       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR12","Standard","","",""    ,"Médio"  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   

   /*
   Aadd( aHelpPor, 'Informe a Filial Inicial da Faixa ')
   Aadd( aHelpPor, 'de Filiais do Relatorio.            ')
                                                                                                           //F3
   PutSx1(cPerg ,"04","Da Filial"  ,"Da Filial","Da Filial","mv_ch4","C"  ,02       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR04","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Filial Final da Faixa ')
   Aadd( aHelpPor, 'de Filiais do Relatorio.            ')
                                                                                                           //F3
   PutSx1(cPerg ,"05","Ate Filial"  ,"Ate Filial","Ate Filial","mv_ch5","C"  ,02       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR05","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   */

   RestArea(aArea)

Return(.T.)

Static Function RABCImp(nOpImp)


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Declaracao de Variaveis                                             ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

Local cDesc1         := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2         := "de acordo com os parametros informados pelo usuario."
Local cDesc3         := "Curva ABC por Produto X Tonelada Vendida"
Local cPict          := ""
Local titulo       := "Curva ABC por Produto X Tonelada Vendida"
Local nLin         := 80

Local Cabec1       := "Codigo           Descrição                                                       Qtd. Clientes     Qtd. Toneladas             Custo              Venda          Margem  Curva"
Local Cabec2       := ""
Local imprime      := .T.
Local aOrd := {}
If nOpImp == 1
   titulo       := "Curva ABC por Produto X Tonelada Vendida"
   Cabec1       := "Codigo           Descrição                                                       Qtd. Clientes     Qtd. Toneladas             Custo              Venda          Margem  Curva"
Else
   titulo       := "Ranking de Produtos por Tonelada Vendida"
   Cabec1       := "Codigo           Descrição                                                       Qtd. Clientes     Qtd. Toneladas             Custo              Venda          Margem"
EndIf

Private lEnd         := .F.
Private lAbortPrint  := .F.
Private CbTxt        := ""
Private limite           := 220
Private tamanho          := "G"
Private nomeprog         := "CRVAABC" // Coloque aqui o nome do programa para impressao no cabecalho
Private nTipo            := 15
Private aReturn          := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey        := 0
Private cbtxt      := Space(10)
Private cbcont     := 00
Private CONTFL     := 01
Private m_pag      := 01
Private wnrel      := "CRVAABC" // Coloque aqui o nome do arquivo usado para impressao em disco

Private nPrintOpc := nOpImp

Private cString := "SD2"

dbSelectArea("SD2")
dbSetOrder(1)


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Monta a interface padrao com o usuario...                           ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

wnrel := SetPrint(cString,NomeProg,"",@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,.F.,Tamanho,,.F.)

If nLastKey == 27
	Return
Endif

SetDefault(aReturn,cString)

If nLastKey == 27
   Return
Endif

nTipo := If(aReturn[4]==1,15,18)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Processamento. RPTSTATUS monta janela com a regua de processamento. ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

RptStatus({|| RunReport(Cabec1,Cabec2,Titulo,nLin) },Titulo)
Return

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFun‡„o    ³RUNREPORT º Autor ³ AP6 IDE            º Data ³  16/07/10   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescri‡„o ³ Funcao auxiliar chamada pela RPTSTATUS. A funcao RPTSTATUS º±±
±±º          ³ monta a janela com a regua de processamento.               º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ Programa principal                                         º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/

Static Function RunReport(Cabec1,Cabec2,Titulo,nLin)

Local nOrdem

dbSelectArea(cString)
dbSetOrder(1)

If nPrintOpc == 1

   //ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
   //³ SETREGUA -> Indica quantos registros serao processados para a regua ³
   //ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
   nRCA := Len(aCurvaA)
   nRCB := Len(aCurvaB)
   nRCC := Len(aCurvaC)

   lRoda := .T.

   //Totalizadores Geral
   nGrCliCva := 0 //Qtd. Clientes que comprarameste produto
   nGrTonCva := 0 //Qtd. Toneladas Vendidas no Período
   nGrCstCva := 0 //Custo s/Impostos      
   nGrVdaCva := 0 //Vendas s/Impostos
   nGrMrgCva := 0 //Margem

   While lRoda
   
      SetRegua(nRCA)

      //Totalizadores por Curva
      nCliCva := 0 //Qtd. Clientes que comprarameste produto
      nTonCva := 0 //Qtd. Toneladas Vendidas no Período
      nCstCva := 0 //Custo s/Impostos      
      nVdaCva := 0 //Vendas s/Impostos
      nMrgCva := 0 //Margem
      //MsgInfo("Len(aCurvaA) "+Alltrim(Str(Len(aCurvaA)))+" Len(aCurvaB) "+Alltrim(Str(Len(aCurvaB)))+" Len(aCurvaC) "+Alltrim(Str(Len(aCurvaC))))
      For nImp := 1 To Len(aCurvaA)
   
         IncRegua()
   
         //ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
         //³ Verifica o cancelamento pelo usuario...                             ³
         //ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

         If lAbortPrint
            @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
            Exit
         Endif


         //ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
         //³ Impressao do cabecalho do relatorio. . .                            ³
         //ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

         If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         //aAdd(aCurvaA,{TRLST->CODPROD,TRLST->DESCPRO,TRLST->QUANCLI,TRLST->QTDTONE,TRLST->CUSTOPR,TRLST->VLRVEND,TRLST->NMARGEM})
        //          10        20        30        40        50        60        70        80        90       100       110       120       130       140       150       160       170       180       190       200       210       220
        // 01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
        //"Codigo           Descrição                                                       Qtd. Clientes     Qtd. Toneladas             Custo              Venda          Margem  Curva"
        // xxxxxxxxxXxxxxx  xxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxX  999,999,999,999   999,999,999.9999    999,999,999,99     999,999,999.99  999,999,999.99    A
     
         If nImp == 1
            @nLin,000 PSAY "Apresentando Curva A ..."
            nLin := nLin + 2 // Avanca a linha de impressao
         EndIf

         If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         
         @nLin,000 PSAY Substr(aCurvaA[nImp,1],1,15)                     //Codigo do Produto
         @nLin,017 PSAY Substr(aCurvaA[nImp,2],1,60)                     //Descrição do Produto
         @nLin,079 PSAY Transform(aCurvaA[nImp,3],"@E 999,999,999,999")  //Qtd. Clientes que comprarameste produto
         @nLin,097 PSAY Transform(aCurvaA[nImp,4],"@E 999,999,999.9999") //Qtd. Toneladas Vendidas no Período
         @nLin,117 PSAY Transform(aCurvaA[nImp,5],"@E 999,999,999.99")   //Custo s/Impostos      
         @nLin,136 PSAY Transform(aCurvaA[nImp,6],"@E 999,999,999.99")   //Vendas s/Impostos
         @nLin,152 PSAY Transform(aCurvaA[nImp,7],"@E 999,999,999.99")+" %"   //Margem
         @nLin,170 PSAY "A"                                              //Curva

         nLin := nLin + 1 // Avanca a linha de impressao

         //Totalizadores por Curva
         nCliCva += aCurvaA[nImp,3] //Qtd. Clientes que comprarameste produto
         nTonCva += aCurvaA[nImp,4] //Qtd. Toneladas Vendidas no Período
         nCstCva += aCurvaA[nImp,5] //Custo s/Impostos      
         nVdaCva += aCurvaA[nImp,6] //Vendas s/Impostos
         nMrgCva += aCurvaA[nImp,7] //Margem

         //Totalizadores Geral
         nGrCliCva += aCurvaA[nImp,3] //Qtd. Clientes que comprarameste produto
         nGrTonCva += aCurvaA[nImp,4] //Qtd. Toneladas Vendidas no Período
         nGrCstCva += aCurvaA[nImp,5] //Custo s/Impostos      
         nGrVdaCva += aCurvaA[nImp,6] //Vendas s/Impostos
         nGrMrgCva += aCurvaA[nImp,7] //Margem
   
      Next nImp
   
      nLin := nLin + 1 // Avanca a linha de impressao 
   
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif   
   
      @nLin,000 PSAY "Totais Curva A ..."
      @nLin,079 PSAY Transform(nCliCva,"@E 999,999,999,999")  //Qtd. Clientes que comprarameste produto
      @nLin,097 PSAY Transform(nTonCva,"@E 999,999,999.9999") //Qtd. Toneladas Vendidas no Período
      @nLin,117 PSAY Transform(nCstCva,"@E 999,999,999.99")   //Custo s/Impostos      
      @nLin,136 PSAY Transform(nVdaCva,"@E 999,999,999.99")   //Vendas s/Impostos
      @nLin,152 PSAY Transform(nMrgCva,"@E 999,999,999.99")+" %"   //Margem

      nLin := nLin + 2 // Avanca a linha de impressao 
   
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif   

      SetRegua(nRCB)

      //Totalizadores por Curva
      nCliCva := 0 //Qtd. Clientes que comprarameste produto
      nTonCva := 0 //Qtd. Toneladas Vendidas no Período
      nCstCva := 0 //Custo s/Impostos      
      nVdaCva := 0 //Vendas s/Impostos
      nMrgCva := 0 //Margem
   
      For nImp := 1 To Len(aCurvaB)
   
         IncRegua()
   
         //ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
         //³ Verifica o cancelamento pelo usuario...                             ³
         //ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

         If lAbortPrint
            @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
            Exit
         Endif

         //ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
         //³ Impressao do cabecalho do relatorio. . .                            ³
         //ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

         If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         //aAdd(aCurvaA,{TRLST->CODPROD,TRLST->DESCPRO,TRLST->QUANCLI,TRLST->,TRLST->CUSTOPR,TRLST->VLRVEND,TRLST->NMARGEM})
        //          10        20        30        40        50        60        70        80        90       100       110       120       130       140       150       160       170       180       190       200       210       220
        // 01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
        //"Codigo           Descrição                                                       Qtd. Clientes     Qtd. Toneladas  Custo s/Impostos  Vendas s/Impostos          Margem  Curva"
        // xxxxxxxxxXxxxxx  xxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxX  999,999,999,999   999,999,999.9999    999,999,999,99     999,999,999.99  999,999,999.99
     
         If nImp == 1
            @nLin,000 PSAY "Apresentando Curva B ..."
            nLin := nLin + 2 // Avanca a linha de impressao
         EndIf

         If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         
         @nLin,000 PSAY Substr(aCurvaB[nImp,1],1,15)                     //Codigo do Produto
         @nLin,017 PSAY Substr(aCurvaB[nImp,2],1,60)                     //Descrição do Produto
         @nLin,079 PSAY Transform(aCurvaB[nImp,3],"@E 999,999,999,999")  //Qtd. Clientes que comprarameste produto
         @nLin,097 PSAY Transform(aCurvaB[nImp,4],"@E 999,999,999.9999") //Qtd. Toneladas Vendidas no Período
         @nLin,117 PSAY Transform(aCurvaB[nImp,5],"@E 999,999,999.99")   //Custo s/Impostos      
         @nLin,136 PSAY Transform(aCurvaB[nImp,6],"@E 999,999,999.99")   //Vendas s/Impostos
         @nLin,152 PSAY Transform(aCurvaB[nImp,7],"@E 999,999,999.99")+" %"   //Margem
         @nLin,170 PSAY "B"                                              //Curva

         nLin := nLin + 1 // Avanca a linha de impressao

         //Totalizadores por Curva
         nCliCva += aCurvaB[nImp,3] //Qtd. Clientes que comprarameste produto
         nTonCva += aCurvaB[nImp,4] //Qtd. Toneladas Vendidas no Período
         nCstCva += aCurvaB[nImp,5] //Custo s/Impostos      
         nVdaCva += aCurvaB[nImp,6] //Vendas s/Impostos
         nMrgCva += aCurvaB[nImp,7] //Margem

         //Totalizadores Geral
         nGrCliCva += aCurvaB[nImp,3] //Qtd. Clientes que compraram este produto
         nGrTonCva += aCurvaB[nImp,4] //Qtd. Toneladas Vendidas no Período
         nGrCstCva += aCurvaB[nImp,5] //Custo s/Impostos      
         nGrVdaCva += aCurvaB[nImp,6] //Vendas s/Impostos
         nGrMrgCva += aCurvaB[nImp,7] //Margem
   
      Next nImp
   
      nLin := nLin + 1 // Avanca a linha de impressao 
   
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif   
   
      @nLin,000 PSAY "Totais Curva B ..."
      @nLin,079 PSAY Transform(nCliCva,"@E 999,999,999,999")  //Qtd. Clientes que compraram este produto
      @nLin,097 PSAY Transform(nTonCva,"@E 999,999,999.9999") //Qtd. Toneladas Vendidas no Período
      @nLin,117 PSAY Transform(nCstCva,"@E 999,999,999.99")   //Custo s/Impostos      
      @nLin,136 PSAY Transform(nVdaCva,"@E 999,999,999.99")   //Vendas s/Impostos
      @nLin,152 PSAY Transform(nMrgCva,"@E 999,999,999.99")+" %"   //Margem

      nLin := nLin + 2 // Avanca a linha de impressao 
   
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif   

      SetRegua(nRCC)

      //Totalizadores por Curva
      nCliCva := 0 //Qtd. Clientes que comprarameste produto
      nTonCva := 0 //Qtd. Toneladas Vendidas no Período
      nCstCva := 0 //Custo s/Impostos      
      nVdaCva := 0 //Vendas s/Impostos
      nMrgCva := 0 //Margem
   
      For nImp := 1 To Len(aCurvaC)
        
         IncRegua()
   
         //ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
         //³ Verifica o cancelamento pelo usuario...                             ³
         //ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

         If lAbortPrint
            @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
            Exit
         Endif

         //ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
         //³ Impressao do cabecalho do relatorio. . .                            ³
         //ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

         If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
      
         //aAdd(aCurvaA,{TRLST->CODPROD,TRLST->DESCPRO,TRLST->QUANCLI,TRLST->QTDTONE,TRLST->CUSTOPR,TRLST->VLRVEND,TRLST->NMARGEM})
        //          10        20        30        40        50        60        70        80        90       100       110       120       130       140       150       160       170       180       190       200       210       220
        // 01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
        //"Codigo           Descrição                                                       Qtd. Clientes     Qtd. Toneladas  Custo s/Impostos  Vendas s/Impostos          Margem  Curva"
        // xxxxxxxxxXxxxxx  xxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxX  999,999,999,999   999,999,999.9999    999,999,999,99     999,999,999.99  999,999,999.99
     
         If nImp == 1
            @nLin,000 PSAY "Apresentando Curva C ..."
            nLin := nLin + 2 // Avanca a linha de impressao
         EndIf

         If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
            Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
            nLin := 8
         Endif
         
         @nLin,000 PSAY Substr(aCurvaC[nImp,1],1,15)                     //Codigo do Produto
         @nLin,017 PSAY Substr(aCurvaC[nImp,2],1,60)                     //Descrição do Produto
         @nLin,079 PSAY Transform(aCurvaC[nImp,3],"@E 999,999,999,999")  //Qtd. Clientes que comprarameste produto
         @nLin,097 PSAY Transform(aCurvaC[nImp,4],"@E 999,999,999.9999") //Qtd. Toneladas Vendidas no Período
         @nLin,117 PSAY Transform(aCurvaC[nImp,5],"@E 999,999,999.99")   //Custo s/Impostos      
         @nLin,136 PSAY Transform(aCurvaC[nImp,6],"@E 999,999,999.99")   //Vendas s/Impostos
         @nLin,152 PSAY Transform(aCurvaC[nImp,7],"@E 999,999,999.99")+" %"   //Margem
         @nLin,170 PSAY "C"                                              //Curva

         nLin := nLin + 1 // Avanca a linha de impressao

         //Totalizadores por Curva
         nCliCva += aCurvaC[nImp,3] //Qtd. Clientes que comprarameste produto
         nTonCva += aCurvaC[nImp,4] //Qtd. Toneladas Vendidas no Período
         nCstCva += aCurvaC[nImp,5] //Custo s/Impostos      
         nVdaCva += aCurvaC[nImp,6] //Vendas s/Impostos
         nMrgCva += aCurvaC[nImp,7] //Margem

         //Totalizadores Geral
         nGrCliCva += aCurvaC[nImp,3] //Qtd. Clientes que comprarameste produto
         nGrTonCva += aCurvaC[nImp,4] //Qtd. Toneladas Vendidas no Período
         nGrCstCva += aCurvaC[nImp,5] //Custo s/Impostos      
         nGrVdaCva += aCurvaC[nImp,6] //Vendas s/Impostos
         nGrMrgCva += aCurvaC[nImp,7] //Margem
   
      Next nImp
   
      nLin := nLin + 1 // Avanca a linha de impressao 
   
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif   
   
      @nLin,000 PSAY "Totais Curva C ..."
      @nLin,079 PSAY Transform(nCliCva,"@E 999,999,999,999")  //Qtd. Clientes que comprarameste produto
      @nLin,097 PSAY Transform(nTonCva,"@E 999,999,999.9999") //Qtd. Toneladas Vendidas no Período
      @nLin,117 PSAY Transform(nCstCva,"@E 999,999,999.99")   //Custo s/Impostos      
      @nLin,136 PSAY Transform(nVdaCva,"@E 999,999,999.99")   //Vendas s/Impostos
      @nLin,152 PSAY Transform(nMrgCva,"@E 999,999,999.99")+" %"   //Margem

      nLin := nLin + 2 // Avanca a linha de impressao 
   
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif   
               
      lRoda := .F.
   
   EndDo

   nLin := nLin + 2 // Avanca a linha de impressao 
   
   If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 8
   Endif   
   
   @nLin,000 PSAY "Totais Gerais ..."
   @nLin,079 PSAY Transform(nGrCliCva,"@E 999,999,999,999")  //Qtd. Clientes que comprarameste produto
   @nLin,097 PSAY Transform(nGrTonCva,"@E 999,999,999.9999") //Qtd. Toneladas Vendidas no Período
   @nLin,117 PSAY Transform(nGrCstCva,"@E 999,999,999.99")   //Custo s/Impostos      
   @nLin,136 PSAY Transform(nGrVdaCva,"@E 999,999,999.99")   //Vendas s/Impostos
   @nLin,152 PSAY Transform(nGrMrgCva,"@E 999,999,999.99")+" %"   //Margem

Else

   
   aRanking := {}
   DbSelectArea("TRLST")
   nReg  := RecCount()
   SetRegua(nReg)
   DbGoTop()
   While TRLST->(!Eof())
      IncRegua()
      //                     1            2              3              4                5            6               7
      aAdd(aRanking,{TRLST->CODPROD,TRLST->DESCPRO,TRLST->QUANCLI,TRLST->QTDTONE,TRLST->CUSTOPR,TRLST->VLRVEND,TRLST->NMARGEM})
      DbSelectArea("TRLST")
      DbSkip()
   EndDo
   DbSelectArea("TRLST")
   DbCloseArea()

   If nOrdRkg == 1 //Quantidade(Tonelada Vendida)
      If nTpOrde == 1
         aSort(aRanking,,, {|x,y| x[4] < y[4]}) //Ordena Ascendente por Qtd. Tonelada Vendida
      Else
         aSort(aRanking,,, {|x,y| x[4] > y[4]}) //Ordena Decrescente por Qtd. Tonelada Vendida
      EndIf
   ElseIf nOrdRkg == 2 //Valor
      If nTpOrde == 1
         aSort(aRanking,,, {|x,y| x[6] < y[6]}) //Ordena Ascendente por Valor Vendido
      Else
         aSort(aRanking,,, {|x,y| x[6] > y[6]}) //Ordena Decrescente por Valor Vendido
      EndIf
   ElseIf nOrdRkg == 3 //Margem
      If nTpOrde == 1
         aSort(aRanking,,, {|x,y| x[7] < y[7]}) //Ordena Ascendente por Margem
      Else
         aSort(aRanking,,, {|x,y| x[7] > y[7]}) //Ordena Decrescente por Margem
      EndIf
   EndIf

   nReg  := Len(aRanking)
   SetRegua(nReg)
   
   //Totalizadores Geral
   nGrCliCva := 0 //Qtd. Clientes que comprarameste produto
   nGrTonCva := 0 //Qtd. Toneladas Vendidas no Período
   nGrCstCva := 0 //Custo s/Impostos      
   nGrVdaCva := 0 //Vendas s/Impostos
   nGrMrgCva := 0 //Margem

   For nImp := 1 To Len(aRanking)
         
      IncRegua()
      
      //ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
      //³ Verifica o cancelamento pelo usuario...                             ³
      //ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

      If lAbortPrint
         @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
         Exit
      Endif

      //ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
      //³ Impressao do cabecalho do relatorio. . .                            ³
      //ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      //aAdd(aCurvaA,{TRLST->CODPROD,TRLST->DESCPRO,TRLST->QUANCLI,TRLST->QTDTONE,TRLST->CUSTOPR,TRLST->VLRVEND,TRLST->NMARGEM})
      //          10        20        30        40        50        60        70        80        90       100       110       120       130       140       150       160       170       180       190       200       210       220
      // 01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
      //"Codigo           Descrição                                                       Qtd. Clientes     Qtd. Toneladas  Custo s/Impostos  Vendas s/Impostos          Margem  Curva"
      // xxxxxxxxxXxxxxx  xxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxX  999,999,999,999   999,999,999.9999    999,999,999,99     999,999,999.99  999,999,999.99
     
      @nLin,000 PSAY Substr(aRanking[nImp,1],1,15)                     //Codigo do Produto
      @nLin,017 PSAY Substr(aRanking[nImp,2],1,60)                     //Descrição do Produto
      @nLin,079 PSAY Transform(aRanking[nImp,3],"@E 999,999,999,999")  //Qtd. Clientes que comprarameste produto
      @nLin,097 PSAY Transform(aRanking[nImp,4],"@E 999,999,999.9999") //Qtd. Toneladas Vendidas no Período
      @nLin,117 PSAY Transform(aRanking[nImp,5],"@E 999,999,999.99")   //Custo s/Impostos      
      @nLin,136 PSAY Transform(aRanking[nImp,6],"@E 999,999,999.99")   //Vendas s/Impostos
      @nLin,152 PSAY Transform(aRanking[nImp,7],"@E 999,999,999.99")+" %"   //Margem

      nLin := nLin + 1 // Avanca a linha de impressao      
      
      //Totalizadores Geral
      nGrCliCva += aRanking[nImp,3] //TRLST->QUANCLI //Qtd. Clientes que comprarameste produto
      nGrTonCva += aRanking[nImp,4] //TRLST->QTDTONE //Qtd. Toneladas Vendidas no Período
      nGrCstCva += aRanking[nImp,5] //TRLST->CUSTOPR //Custo s/Impostos      
      nGrVdaCva += aRanking[nImp,6] //TRLST->VLRVEND //Vendas s/Impostos
      nGrMrgCva += aRanking[nImp,7] //TRLST->NMARGEM //Margem
         
     
   Next nImp
         
   nLin := nLin + 2 // Avanca a linha de impressao 
   
   If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 8
   Endif   
   
   @nLin,000 PSAY "Totais Gerais ..."
   @nLin,079 PSAY Transform(nGrCliCva,"@E 999,999,999,999")  //Qtd. Clientes que comprarameste produto
   @nLin,097 PSAY Transform(nGrTonCva,"@E 999,999,999.9999") //Qtd. Toneladas Vendidas no Período
   @nLin,117 PSAY Transform(nGrCstCva,"@E 999,999,999.99")   //Custo s/Impostos      
   @nLin,136 PSAY Transform(nGrVdaCva,"@E 999,999,999.99")   //Vendas s/Impostos
   @nLin,152 PSAY Transform(nGrMrgCva,"@E 999,999,999.99")+" %"   //Margem
   
EndIf


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Finaliza a execucao do relatorio...                                 ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

SET DEVICE TO SCREEN

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Se impressao em disco, chama o gerenciador de impressao...          ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

If aReturn[5]==1
   dbCommitAll()
   SET PRINTER TO
   OurSpool(wnrel)
Endif

MS_FLUSH()

Return
