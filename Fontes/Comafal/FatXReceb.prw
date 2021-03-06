#include "TopConn.Ch"
User Function FatXReceb
   Private cPerg := "FATRCB"
   MontPerg()
   If Pergunte(cPerg,.T.)
      Processa({||RunFXC()})
   EndIf
Return

Static Function RunFXC()
   
   For nFXC := 1 To 2
      
      If nFXC == 1
         cQryFXC := " SELECT COUNT(DISTINCT SF2.F2_DOC+SF2.F2_SERIE*) NTOTFXC "
      Else
         cQryFXC := " SELECT SF2.F2_CLIENTE,SF2.F2_LOJA,SF2.F2_DOC,SF2.F2_SERIE,SF2.F2_EMISSAO "
      EndIf
      
      cQryFXC += RetSqlName('SD2') + " SD2,"+RetSQLName("SB1")+" SB1,"+RetSQLName("SF2")+" SF2"
      cQryFXC += " WHERE SD2.D2_FILIAL  = '"+xFilial("SD2")+"' AND  SD2.D_E_L_E_T_ = ' ' "
      cQryFXC += " AND SB1.B1_FILIAL  = '"+xFilial("SB1")+"' AND  SB1.D_E_L_E_T_ = ' ' "
      cQryFXC += " AND SF2.F2_FILIAL  = '"+xFilial("SF2")+"' AND  SF2.D_E_L_E_T_ = ' ' "
      cQryFXC += " AND SF2.F2_CLIENTE <> '006629'" //Five Solutions Consultoria - 04/11/2009
      cQryFXC += " AND SD2.D2_COD = SB1.B1_COD"
      cQryFXC += " AND SUBSTRING(SB1.B1_GRUPO,1,2) <> '40'" //Faturamento sem Sucatas
      cQryFXC += " AND SF2.F2_DOC = SD2.D2_DOC"
      cQryFXC += " AND SF2.F2_SERIE = SD2.D2_SERIE"
      cQryFXC += " AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
      cQryFXC += " AND (SD2.D2_TIPO <> 'I')" //Five Solutions(07/04/2008)
      cQryFXC += " AND SD2.D2_EMISSAO BETWEEN '"+DTOS(dDtIni)+"' AND '"+DTOS(dDtFim)+"'"
      cQryFXC += " AND SF2.F2_CLIENTE BETWEEN '"+cCliIni+"' AND '"+cCliFim+"'"
      cQryFXC += " AND SF2.F2_LOJA BETWEEN '"+cLojIni+"' AND '"+cLojFim+"'"
      
      MemoWrite("FaturamentoXBaixas.SQL",cQryFXC)
      
      TCQuery cQryFXC NEW ALIAS "TFXC"
      
      If nFXC == 1
         DbSelectArea("TFXC")
         nRegFXC := TFXC->NTOTFXC
         DbCloseArea()
      EndIf
      
   Next nFXC
   
   TCSetField("TFXC","F2_EMISSAO","D",08,00)
   
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

   Aadd( aHelpPor, 'Informe o Codigo de Cliente Inicial')
   Aadd( aHelpPor, 'da Faixa de clientes do Relatorio.   ')
 //PutSx1(cGrupo,cOrdem,cPergunt       ,""           ,""           ,cVar    ,cTipo,nTamanho ,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
   PutSx1(cPerg ,"03"  ,"Do Cliente"  ,"Do Cliente","Do Cliente","mv_ch3","C"  ,6       ,0       ,0      ,"G" ,""    ,"SA1" ,""     ,"","MV_PAR03","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a loja inicial do cliente')
   Aadd( aHelpPor, 'da faixa de lojas do clientes do Relatorio.   ')
 //PutSx1(cGrupo,cOrdem,cPergunt       ,""           ,""           ,cVar    ,cTipo,nTamanho ,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
   PutSx1(cPerg ,"04"  ,"Da Loja"  ,"Da Loja","Da Loja","mv_ch4","C"  ,2       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR04","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe o Codigo do Cliente Final')
   Aadd( aHelpPor, 'da Faixa de Clientes do Relatorio.   ')
 //PutSx1(cGrupo,cOrdem,cPergunt       ,""           ,""           ,cVar    ,cTipo,nTamanho ,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
   PutSx1(cPerg ,"05"  ,"At? Cliente"  ,"At? Cliente","At? Cliente","mv_ch5","C"  ,6       ,0       ,0      ,"G" ,""    ,"SA1" ,""     ,"","MV_PAR05","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a loja Final do cliente')
   Aadd( aHelpPor, 'da faixa de lojas do clientes do Relatorio.   ')
 //PutSx1(cGrupo,cOrdem,cPergunt       ,""           ,""           ,cVar    ,cTipo,nTamanho ,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
   PutSx1(cPerg ,"06"  ,"Da Loja"  ,"Da Loja","Da Loja","mv_ch6","C"  ,2       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR06","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   RestArea(aArea)

Return(.T.)