#INCLUDE "TopConn.Ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³PE_MT100AGRºAutor  ³Five Solutions      º Data ³  10/03/2009 º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ PE Executado após a inclusão da NF de Entrada, usado para  º±±
±±º          ³ impressão das Etiquetas de produtos.                       º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ COMAFAL PE/RS/SP                                          º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function MT100AGR
   
   Private cAreaAki := GetArea()
   
   Private cNumNF   := SF1->F1_DOC
   Private cSerNF   := SF1->F1_SERIE
   Private cForNF   := SF1->F1_FORNECE
   Private cLojNF   := SF1->F1_LOJA
   
   Private cProdCod := ""
   Private nPrdPeso := 0
   Private cLtePrd  := ""
   
   Private xPerg := "PRGETQ" //Five Solutions Consultoria - 07/05/2010
   MtPergEtq()
   
   cQrySF1 := "  SELECT SD1.D1_COD,SD1.D1_LOTECTL,SUM(SD1.D1_QUANT) D1_QUANT "
   cQrySF1 += "    FROM "+RetSQLName("SD1")+" SD1"
   cQrySF1 += "   WHERE SD1.D1_FILIAL = '"+xFilial("SD1")+"'"
   cQrySF1 += "     AND SD1.D1_DOC = '"+cNumNF+"'"
   cQrySF1 += "     AND SD1.D1_SERIE = '"+cSerNF+"'"
   cQrySF1 += "     AND SD1.D1_TP = 'MP'"
   cQrySF1 += "     AND SD1.D_E_L_E_T_ <> '*' "
   cQrySF1 += "   GROUP BY SD1.D1_COD,SD1.D1_LOTECTL"
   cQrySF1 += "   ORDER BY SD1.D1_COD,SD1.D1_LOTECTL"
   
   TCQuery cQrySF1 NEW ALIAS "TSD1"
   
   TCSetField("TSD1","D1_QUANT","N",11,03)
   
   DbSelectArea("TSD1")
   While !Eof()
      cProdCod := Alltrim(TSD1->D1_COD)
      nPrdPeso := TSD1->D1_QUANT
      cLtePrd  := TSD1->D1_LOTECTL
      RImpEtq()
      DbSelectArea("TSD1")
      DbSkip()
   EndDo
   DbSelectArea("TSD1")
   DbCloseArea()
   
   RestArea(cAreaAki)
   
Return

Static Function RImpEtq
   
   //PADRÃO DATAMAX - ARGOX
   cCodUser := Alltrim(cUserName)
   cCodBarr := Alltrim(cProdCod) + Alltrim(Str(nPrdPeso)) + cCodUser + DtoS(dDataBase)
   cHoraProd:= Substr(Time(),1,5)

   cPedidNF := Posicione("SD1",1,xFilial("SD1")+cNumNF+cSerNF+cForNF+cLojNF,"D1_PEDIDO")
   cCodConf := Posicione("SC7",1,xFilial("SC7")+cPedidNF,"C7_USER")
   cNomeConf:= ""
   
   PswOrder(1)
   If PswSeek(cCodConf)
      aUsuario  := PswRet()
      cNomeConf := Alltrim(aUsuario[1][2])
   EndIf
   
   cGrpProd   := Posicione("SB1",1,xFilial("SB1")+cProdCod,"B1_GRUPO")
   cDescPrd   := Posicione("SB1",1,xFilial("SB1")+cProdCod,"B1_DESC")
   
   //cDescMaq := Alltrim(Posicione("SH1",1,xFilial("SH1")+cMaqProd,"H1_DESCRI"))
   cTpPeso  := "Ton" //IIF(nPrdPeso>1000,"Ton","Kg") 
   
   nQtdEtq := 1 //Quantidade de Etiquetas a Serem Impressas
   lSliter    := .F.                                                                       //Controla impressão de informações específicas do SLITER. 
   /*
   If Alltrim(cGrpProd) == "20" //Se o Produto for do Grupo Sliter

      lSliter     := .F. //.T. 
      cDestProdu  := ALLTRIM(SC2->C2_NUM) + ALLTRIM(SC2->C2_ITEM) + Substr(ALLTRIM(SC2->C2_SEQUEN),1,2)+"1"  //Número da Ordem de Produção(O.P.) do Produto destino(Especifico para SLITERS"
      cPrdDestino := Posicione("SC2",1,xFilial("SC2")+cDestProdu,"C2_PRODUTO") //Pega o Código do Produto Destino do SLITER
      cDescDest   := Posicione("SB1",1,xFilial("SB1")+cPrdDestino,"B1_DESC") //Pega o Descrição do Produto Destino do SLITER
      
      VldPerg()
      
      If Pergunte(cPerg,.T.)
         nQtdEtq := MV_PAR01
         nPesoImp := nPrdPeso / nQtdEtq
         //If MsgYesNo("Confirma a Impressão de "+Alltrim(Str(nQtdEtq))+" Etiquesta(s) com "+Alltrim(Str(nPesoImp))+" Ton cada ?")
         If MsgYesNo("Confirma a Impressão de "+Alltrim(Str(nQtdEtq))+" Etiquesta(s) com "+Ltrim(Transform(nPesoImp,"@E 999,999,999.999"))+" Ton cada ?")
         Else
            Return
         EndIf
      Else 
         Return
      EndIf
   Else
   */
      nPesoImp := nPrdPeso / nQtdEtq
   //EndIf
   
   
   If Pergunte(xPerg,.T.) //Five Solutions Consultoria - 07/05/2010
      
      nEtqOpc := MV_PAR01
   
      If nEtqOpc == 1 //Impressão ARGOX   
          
         For nImpEtq := 1 To nQtdEtq
   
            cPorta := "LPT1:9600,e,7,2"
            MSCBPRINTER("DATAMAX",cPorta,NIL,145,.F.,,,,,,.T.)
            MSCBCHKSTATUS(.F.)               // STATUS DA IMPRESSORA
   
            MSCBBEGIN(1,6, 145) 
   
            MSCBSAY(05,114,Substr(cDescPrd,1,14),"N","4","01,01")
            MSCBSAY(05,107,Substr(cDescPrd,15,30),"N","4","01,01")
      
            MSCBSAY(05,098,"Codigo: "+Alltrim(cProdCod),"N","3","01,01") 
      
            //If lSliter 
            //   MSCBSAY(05,093,Alltrim(cDescDest),"N","3","01,01") 
            //   MSCBSAY(05,088,"N.OP ou OC: "+Alltrim(cCorteOrd),"N","3","01,01") 
            //Else
            //MSCBSAY(05,093,"Quant. de Pecas: "+Alltrim(Str(nQtdPecas)),"N","3","01,01") 
            //MSCBSAY(05,088,"N.OP ou OC: "+Alltrim(cProducao),"N","3","01,01") 
            MSCBSAY(05,093,"Lote:            "+Alltrim(cLtePrd),"N","3","01,01") 
            MSCBSAY(05,088,"NF/Série:   "+Alltrim(cNumNF)+"/"+Alltrim(cSerNF),"N","3","01,01")          
            //EndIf
      
            //MSCBSAY(05,083,"Operador.: "+Alltrim(cUserName),"N","3","01,01") 
            //If lSliter
            //   MSCBSAY(05,078,"Conf: "+cNomeConf+" Peso("+cTpPeso+"): "+Ltrim(Transform(nPesoImp,"@E 999,999,999.999")),"N","3","01,01") 
            //   MSCBSAY(05,073,"Lote: "+Alltrim(cNumLte),"N","3","01,01")       
            //Else
            MSCBSAY(05,078,"Conferente: "+cNomeConf,"N","3","01,01") 
            MSCBSAY(05,073,"Peso("+cTpPeso+"): "+Ltrim(Transform(nPesoImp,"@E 999,999,999.999")),"N","3","01,01") 
            //EndIf
            MSCBSAY(05,068,"Data: "+DtoC(dDataBase)+" Hora: "+cHoraProd,"N","3","01,01") 
            //MSCBSAY(05,063,"Turno: ("+Alltrim(cTurno)+") AMARRADO: "+StrZero(nSeqProd,3),"N","3","01,01") 
            MSCBLineH(05,060,66,1) 
            MSCBLineH(05,043,66,1) 
            MSCBSAYBAR(11,028,cCodBarr,"N","MB07",10,.F.,.T.,nil,'C',3,2,.f.)
            MSCBSAYBAR(11,008,cCodBarr,"N","MB07",10,.F.,.T.,nil,'C',3,2,.f.)
            //MSCBSAY(11,004,"Cod: "+Alltrim(cProdCod)+" / Peso: "+Alltrim(Str(nPesoImp)),"N","2","01,01") 
            MSCBSAY(11,004,"Cod: "+Alltrim(cProdCod)+" / Peso: "+Ltrim(Transform(nPesoImp,"@E 999,999,999.999")),"N","2","01,01") 
   
            MSCBEND()
            MSCBCLOSEPRINTER()   
   
         Next nImpEtq
   
      Else //Impressão ZEBRA
         
         For nImpEtq := 1 To nQtdEtq
            

            cPorta := "LPT1"
            MSCBPRINTER("S4M",CPORTA,,120,.F.)
            MSCBCHKSTATUS(.F.)               // STATUS DA IMPRESSORA
            MSCBBEGIN(1,6) 

            /*
            MSCBSAY(05,114,Substr(cDescPrd,1,14),"N","E","01,01")
            MSCBSAY(05,107,Substr(cDescPrd,15,30),"N","E","01,01")
            MSCBSAY(05,098,"Codigo: "+Alltrim(cProdCod),"N","3","01,01") 
            */
           MSCBSAY(05,037,Substr(cDescPrd,1,14),"N","D","040,020")
           MSCBSAY(05,044,Substr(cDescPrd,15,30),"N","D","040,020")
           MSCBSAY(05,051,"Codigo: "+Alltrim(cProdCod),"N","B","30,14")             
      
            /*
            MSCBSAY(05,093,"Lote:            "+Alltrim(cLtePrd),"N","3","01,01") 
            MSCBSAY(05,088,"NF/Série:   "+Alltrim(cNumNF)+"/"+Alltrim(cSerNF),"N","3","01,01")          
            */
           MSCBSAY(05,058,"Lote:"+(cLtePrd),"N","B","30,14") 
           MSCBSAY(05,065,"NF/Serie:   "+Alltrim(cNumNF)+"/"+Alltrim(cSerNF),"N","B","30,14")                      
            
            /*
            MSCBSAY(05,078,"Conferente: "+cNomeConf,"N","3","01,01") 
            MSCBSAY(05,073,"Peso("+cTpPeso+"): "+Ltrim(Transform(nPesoImp,"@E 999,999,999.999")),"N","3","01,01") 
            */
      	   MSCBSAY(05,072,"Conferente: "+cNomeConf,"N","B","30,14") 
       	   MSCBSAY(05,079,"Peso("+cTpPeso+"): "+Ltrim(Transform(nPesoImp,"@E 999,999,999.999")),"N","B","30,14")                      

            /*
            MSCBSAY(05,068,"Data: "+DtoC(dDataBase)+" Hora: "+cHoraProd,"N","3","01,01") 
            MSCBLineH(05,060,66,1) 
            MSCBLineH(05,043,66,1) 
            MSCBSAYBAR(11,028,cCodBarr,"N","MB07",10,.F.,.T.,nil,'C',3,2,.f.)
            MSCBSAYBAR(11,008,cCodBarr,"N","MB07",10,.F.,.T.,nil,'C',3,2,.f.)
            MSCBSAY(11,004,"Cod: "+Alltrim(cProdCod)+" / Peso: "+Ltrim(Transform(nPesoImp,"@E 999,999,999.999")),"N","2","01,01") 
            */
            
            MSCBSAY(05,093,"Data: "+DtoC(dDataBase)+" Hora: "+cHoraProd,"N","B","30,14") 
            MSCBLineH(05,112,86,1) 
            MSCBLineH(05,136,86,1)             

          //MSCBSAYBAR(nXmm, nYmm, cConteudo, cRotacao, cTypePrt, nAltura, lDigVer, lLinha,
          //           lLinBaixo, cSubSetIni, nLargura, nRelacao, lCompacta, lSerial, cIncr, lZerosL)
            MSCBSAYBAR(012,140,cCodBarr,"N","MB07",14 ,.F. ,.T.,nil,'C',2  ,2.2  ,.f.)
            MSCBSAYBAR(012,166,cCodBarr,"N","MB07",14 ,.F. ,.T.,nil,'C',2  ,2.2  ,.f.)
          //            1   2      3     4    5     6   7    8   9  10  11 12   13
            
            MSCBSAY(012,184,"Cod: "+Alltrim(cProdCod)+" / Peso: "+Ltrim(Transform(nPesoImp,"@E 999,999,999.999")),"N","C","015,008") 
            
            MSCBEND()
            MSCBCLOSEPRINTER()   
   
         Next nImpEtq
            
      EndIf
   
   EndIf
   /*
   nNumAmar ++ //Incrementa o numero do amarrado - Temporário
   oNumAmar:Refresh()
   */
   
Return

Static Function MtPergEtq //Five Solutions Consultoria - 07/05/2010

   Local aArea := GetArea()

  
 
   Local aHelpPor := {}
   Local aHelpEng := {}
   Local aHelpSpa := {}

   Aadd( aHelpPor, 'Informe o Modelo de Etiqueta           ')
   Aadd( aHelpPor, 'A informação do modelo correto possibi ')
   Aadd( aHelpPor, 'lita o formato compativel a impressora ')
   Aadd( aHelpPor, 'e o layout da etiqueta.                ')


 //PutSx1(cGrupo,cOrdem,cPergunt              ,""                  ,""                  ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"01"  ,"Modelo da Etiqueta"  ,"Modelo da Etiqueta","Modelo da Etiqueta","MV_CH1","N"  ,1       ,0       ,1      ,"C" ,""    ,"" ,""     ,"","MV_PAR01","ARGOX","","",""    ,"ZEBRA"      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
   
   RestArea(aArea)

Return(.T.)   