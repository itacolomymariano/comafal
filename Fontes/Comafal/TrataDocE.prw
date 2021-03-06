#include "Protheus.Ch"
#INCLUDE "TopConn.Ch"

/*
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????ͻ??
???Programa  ?TRATADOCE ?Autor  ?Five Solutions      ? Data ?  06/10/2010 ???
?????????????????????????????????????????????????????????????????????????͹??
???Desc.     ? Processo de tratamento para ajustes de lotes automaticos   ???
???          ? de produtos gerados em NF com erro de dados na emiss?o.    ???
???          ? A Rotina far? ajuste com base nos seguintes crit?rios.     ???
???          ?                                                            ???
???          ? SD1 - ?tem NFs de Entrada                                  ??? 
???          ? Trocar o n?mero dos lotes, ficando a NF CORRETA com o lote ??? 
???          ? da NF ERRADA  e vice-versa                                 ??? 
???          ? SD5 - Movimenta??es por Lote                               ???  
???          ? - Trocar o n?mero e s?rie da NF que gerou o Lote, ficando  ??? 
???          ?   o Lote da NF CORRETA vinculado ? NF ERRADA  e vice-versa ??? 
???          ? SB8 - Saldos por Lote                                      ??? 
???          ? - Trocar N?mero, S?rie e data da gera??o do Lote conforme  ??? 
???          ?   crit?rio acima.                                          ???
???          ?                                                            ???
?????????????????????????????????????????????????????????????????????????͹??
???Uso       ? COMAFAL - PE,SP e RS                                      ???
?????????????????????????????????????????????????????????????????????????ͼ??
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
*/

User Function TrataDocE
   Private cPerg := "TRADOE"
   MontPerg()
   If MsgYesNo("Deseja Processar De-Para de Lotes de Produtos das NFs de Entrada Errada p/ NF Corretas - BrasMundi ?")
      If Pergunte(cPerg,.T.)
         Private cArqPth := Alltrim(MV_PAR01)
         Processa({||RTrtDOE()},"Processando Corre??o de Lotes das NFEs, Aguarde ...")
         MsgInfo("Processo Concluido com Sucesso!!!")
      EndIf
   EndIf
Return
Static Function RTrtDOE()
   If File(cArqPth)
      DbUseArea(.T.,"DBFCDX",cArqPth,"TDOE")
      DbSelectArea("TDOE")
      nReg := RecCount()
      nCont:= 1
      DbGoTop()
      While TDOE->(!Eof())
         
         IncProc("NF/S?rie ["+TDOE->DOCERRO+"/"+TDOE->DOCSERIE+"] "+Alltrim(Str(nCont))+" / "+Alltrim(Str(nReg)))
         
         cQryNFErro := " SELECT SD1.D1_FILIAL,SD1.D1_DOC,SD1.D1_SERIE,SD1.D1_COD,SD1.D1_LOTECTL, "
         cQryNFErro += "        SD1.D1_EMISSAO,SD1.D1_DTDIGIT,SD1.R_E_C_N_O_ RECSD1"
         cQryNFErro += "   FROM "+RetSQLName("SD1")+" SD1 "
         cQryNFErro += "  WHERE SD1.D1_FILIAL = '"+xFilial("SD1")+"'"
         cQryNFErro += "    AND SD1.D1_DOC = '"+TDOE->DOCERRO+"'"
         cQryNFErro += "    AND SD1.D1_SERIE = '"+TDOE->DOCSERIE+"'"
         cQryNFErro += "    AND SD1.D_E_L_E_T_ <> '*'"
         MemoWrite("TratDOCE_NFErrada.SQL",cQryNFErro)
         MemoWrite("C:\TEMP\TratDOCE_NFErrada.SQL",cQryNFErro)
         TCQuery cQryNFErro NEW ALIAS "NFERRO"
         TCSetField("NFERRO","RECSD1","N",10,00)
         TCSetField("NFERRO","D1_EMISSAO","D",08,00)
         TCSetField("NFERRO","D1_DTDIGIT","D",08,00)
         
         DbSelectArea("NFERRO")
         While NFERRO->(!Eof())
            
            nGoESD1 := NFERRO->RECSD1
            cLteENF := NFERRO->D1_LOTECTL
            dEmiENF := NFERRO->D1_DTDIGIT //EMISSAO
                        
            cQryCertoNF := " SELECT TMP.D1_LOTECTL,TMP.D1_EMISSAO,TMP.D1_DTDIGIT,TMP.R_E_C_N_O_ RECTMP "
            cQryCertoNF += "   FROM "+RetSQLName("SD1")+" TMP "
            cQryCertoNF += "  WHERE TMP.D1_FILIAL = '"+xFilial("SD1")+"'"
            cQryCertoNF += "    AND TMP.D1_DOC = '"+TDOE->DOCERTO+"'"
            cQryCertoNF += "    AND TMP.D1_COD = '"+NFERRO->D1_COD+"'"
            cQryCertoNF += "    AND TMP.D1_SERIE = '"+TDOE->DOCSERIE+"'"
            cQryCertoNF += "    AND TMP.D_E_L_E_T_ <> '*'"
            MemoWrite("TratDOCE_NFCerta.SQL",cQryCertoNF)
            MemoWrite("C:\TEMP\TratDOCE_NFCerta.SQL",cQryCertoNF)
            TCQuery cQryCertoNF NEW ALIAS "NFCERTO"
            TCSetField("NFCERTO","RECTMP","N",10,00)
            TCSetField("NFCERTO","D1_EMISSAO","D",08,00)
            TCSetField("NFCERTO","D1_DTDIGIT","D",08,00)
            
            DbSelectArea("NFCERTO")
            While NFCERTO->(!Eof())
               nGoCSD1 := NFCERTO->RECTMP
               cLteCNF := NFCERTO->D1_LOTECTL
               dEmiCNF := NFCERTO->D1_DTDIGIT //EMISSAO

               BEGIN TRANSACTION
                  
                  //Atualizando Lote da NF Certa com Lote da NF Errada
                  DbSelectArea("SD1")
                  DbGoTo(nGoCSD1) //Posiciona no Produto da NF Certa
                  RecLock("SD1",.F.)
                     SD1->D1_LOTECTL := cLteENF //Grava Lote da NF Errada
                  MsUnLock()
               
                  //Atualizando Lote da NF Errada com Lote da NF Certa
                  DbSelectArea("SD1")
                  DbGoTo(nGoESD1) //Posiciona no Produto da NF Errada
                  RecLock("SD1",.F.)
                     SD1->D1_LOTECTL := cLteCNF //Grava Lote da NF Certa
                  MsUnLock()
                  
                  dDtLte := fTrtSD5(cLteENF,TDOE->DOCERTO,TDOE->DOCSERIE,NFERRO->D1_COD,TDOE->DOCERRO,dEmiCNF)              
                  fSB8Trt(cLteENF,TDOE->DOCERTO,TDOE->DOCSERIE,NFERRO->D1_COD,dEmiCNF)
                  
                  dDtLte := fTrtSD5(cLteCNF,TDOE->DOCERRO,TDOE->DOCSERIE,NFERRO->D1_COD,TDOE->DOCERTO,dEmiENF)              
                  fSB8Trt(cLteCNF,TDOE->DOCERRO,TDOE->DOCSERIE,NFERRO->D1_COD,dEmiENF)
               
               END TRANSACTION
               
               DbSelectArea("NFCERTO")
               DbSkip()
               
            EndDo
            DbSelectArea("NFCERTO")
            DbCloseArea()
            
            DbSelectArea("NFERRO")
            DbSkip()
         EndDo
         DbSelectArea("NFERRO")
         DbCloseArea()
         
         DbSelectArea("TDOE")
         DbSkip()
         nCont ++
      EndDo
      DbSelectArea("TDOE")
      DbCloseArea()
   Else
      Alert("O Arquivo "+cArqPth+" N?o foi localizado na pasta, Opera??o ser? cancelada")
      Return
   EndIf
Return
Static Function MontPerg

   Local aArea := GetArea()

   //PutSx1(cGrupo,cOrdem,cPergunt  ,"","",cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
 
   Local aHelpPor := {}
   Local aHelpEng := {}
   Local aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a localiza??o do arquivo De-Para')
                                                                                                           //F3
   PutSx1(cPerg ,"01","Local/Arquivo De-Para"  ,"Local/Arquivo De-Para","Local/Arquivo De-Para","mv_ch1","C"  ,90       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR01","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   RestArea(aArea)

Return(.T.)

Static Function fTrtSD5(cLteSD5,cDOCNFE,cSERNFE,cProdNFE,cDOCCNF,dEmiSD5)
   //fTrtSD5(cLteENF,TDOE->DOCERTO,TDOE->DOCSERIE,NFERRO->D1_COD,TDOE->DOCERRO)              
   //fTrtSD5(cLteCNF,TDOE->DOCERRO,TDOE->DOCSERIE,NFERRO->D1_COD,TDOE->DOCERTO)              
   cQrySD5 := " SELECT SD5.D5_DATA,SD5.R_E_C_N_O_ RECSD5 "
   cQrySD5 += "   FROM "+RetSQLName("SD5")+" SD5 "
   cQrySD5 += "  WHERE SD5.D5_FILIAL = '"+xFilial("SD5")+"'"
   cQrySD5 += "    AND SD5.D5_DOC = '"+cDOCCNF+"'"
   cQrySD5 += "    AND SD5.D5_LOTECTL = '"+cLteSD5+"'"
   cQrySD5 += "    AND SD5.D5_PRODUTO = '"+cProdNFE+"'"
   cQrySD5 += "    AND SD5.D5_DATA >= '20100101'"
   cQrySD5 += "    AND SD5.D_E_L_E_T_ <> '*'"
   MemoWrite("MovimentosLoteSD5.SQL",cQrySD5)
   MemoWrite("C:\TEMP\MovimentosLoteSD5.SQL",cQrySD5)
   TCQuery cQrySD5 NEW ALIAS "TSD5"
   TCSetField("TSD5","RECSD5","N",10,00)
   TCSetField("TSD5","D5_DATA","D",08,00)
   DbSelectArea("TSD5")
   dRetDat := TSD5->D5_DATA
   While TSD5->(!Eof())
      nGoSD5 := TSD5->RECSD5
      DbSelectArea("SD5")
      DbGoTo(nGoSD5)
      RecLock("SD5",.F.)
         SD5->D5_DOC   := cDOCNFE
         SD5->D5_SERIE := cSERNFE
         //SD5->D5_DATA  := dEmiSD5   - N?o modifica data do SD5 para preservar hist?rico de movimenta??o.
      MsUnLock()
      
      DbSelectArea("TSD5")
      DbSkip()
   EndDo
   DbSelectArea("TSD5")
   DbCloseArea()
   
Return(dRetDat)
Static Function fSB8Trt(cLteSB8,cDOCSB8,cSERSB8,cProdSB8,dDatLte)
   //fSB8Trt(cLteENF,TDOE->DOCERTO,TDOE->DOCSERIE,NFERRO->D1_COD,dDtLte)
   //fSB8Trt(cLteCNF,TDOE->DOCERRO,TDOE->DOCSERIE,NFERRO->D1_COD,dDtLte)
   cQrySB8 := " SELECT SB8.R_E_C_N_O_ RECSB8 "
   cQrySB8 += "   FROM "+RetSQLName("SB8")+" SB8 "
   cQrySB8 += "  WHERE SB8.B8_FILIAL = '"+xFilial("SB8")+"'"
   cQrySB8 += "    AND SB8.B8_PRODUTO = '"+cProdSB8+"'"
   cQrySB8 += "    AND SB8.B8_LOTECTL = '"+cLteSB8+"'"
   cQrySB8 += "    AND SB8.D_E_L_E_T_ <> '*'"     
   MemoWrite("SaldosSB8.SQL",cQrySB8)
   MemoWrite("C:\TEMP\SaldosSB8.SQL",cQrySB8)
   TCQuery cQrySB8 NEW ALIAS "TSB8"
   TCSetField("TSB8","RECSB8","N",10,00)
   DbSelectArea("TSB8")
   While TSB8->(!Eof())
      nGoSB8 := TSB8->RECSB8
      DbSelectArea("SB8")
      DbGoTo(nGoSB8)
      RecLock("SB8",.F.)
         SB8->B8_DOC   := cDOCSB8
         SB8->B8_SERIE := cSERSB8
         SB8->B8_DATA  := dDatLte
      MsUnLock()
      DbSelectArea("TSB8")
      DbSkip()
   EndDo
   DbSelectArea("TSB8")
   DbCloseArea()
Return