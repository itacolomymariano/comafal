#include "TopConn.Ch"
/*
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????ͻ??
???Programa  ?CALCDISPEST?Autor  ?Five Solutions      ? Data ?  10/01/2008 ???
?????????????????????????????????????????????????????????????????????????͹??
???Desc.     ? Valida??o de Usu?rio do campo C6_PRODUTO.                  ???
???          ? Sistema far? consist?ncia entre a quantidade digitada no   ???
???          ? pedido de pedido de vendas em rela??o a disponibilidade do ???
???          ? saldo em estoque.                                          ???
?????????????????????????????????????????????????????????????????????????͹??
???Uso       ? COMAFAL                                                   ???
?????????????????????????????????????????????????????????????????????????ͼ??
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
*/

User Function CalcDispEst
   Local lRetVld   := .T.
   Local nAreaAki  := GetArea()
   Local cNumSC5   := M->C5_NUM //Five Solutions Consultoria - 16/12/2009
   Local nQtdC6    := M->C6_QTDVEN
   Local nPosLoc   := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_LOCAL"  })  
   Local cLocPrd   := aCols[n,nPosLoc] //GdFieldGet("C6_LOCAL",n)
   Local cProduto  := GdFieldGet("C6_PRODUTO",n)
   Local cNomePrd  := Posicione("SB1",1,xFilial("SB1")+cProduto,"B1_DESC")
   Local cTESPed   := GdFieldGet("C6_TES",n)
   Local lEstoque  := IIF(Posicione("SF4",1,xFilial("SF4")+cTESPed,"F4_ESTOQUE") == "S",.T.,.F.)
   Local nSaldoPrd := 0
   Local nQtVend   := GdFieldGet("C6_QTDVEN",n) //Five Solutions Consultoria - 23/02/2010
   Local cItemC6   := GdFieldGet("C6_ITEM",n)
   
   //Five Solutions Consultoria - 16/12/2009 
   cNomeFunc := Alltrim(FunName())
   Private lConfEtq  := .F.
   If Alltrim(cNomeFunc) == "CFPVXETQ"
      lConfEtq  := .T.
      DbSelectArea("SC6")
      DbSetOrder(2)
      If DbSeek(xFilial("SC6")+cProduto+cNumSC5+cItemC6)
         cLocPrd   := SC6->C6_LOCAL
         nQtVend   := SC6->C6_QTDVEN //Five Solutions Consultoria - 23/02/2010
      EndIf
   EndIf


   //Five Solutions Consultoria - 23/02/2010
   //Implementado Query para checar se o pedido est? liberado.
    
   cQryPV := "  SELECT COUNT(*) TOTREG"

   cQryPV += "    FROM "+RetSQlName("SC5")+" SC5,"+RetSQLName("SA1")+" SA1,"+RetSQLName("SB1")+" SB1,"
   cQryPV +=             RetSQLName("SF4")+" SF4,"+RetSQLName("SC6")+" SC6, "+RetSQLName("SC9")+" SC9"
            
   cQryPV += "   WHERE SA1.A1_FILIAL = '"+xFilial("SA1")+"'"
   cQryPV += "     AND SC5.C5_FILIAL = '"+xFilial("SC5")+"'"
   cQryPV += "     AND SC6.C6_FILIAL = '"+xFilial("SC6")+"'"
   cQryPV += "     AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"
   cQryPV += "     AND SF4.F4_FILIAL = '"+xFilial("SF4")+"'"

   cQryPV += "    AND SC9.C9_FILIAL = '"+xFilial("SC9")+"'"
   cQryPV += "    AND SC6.C6_PRODUTO = SC9.C9_PRODUTO"
   cQryPV += "    AND SC6.C6_NUM = SC9.C9_PEDIDO"
   cQryPV += "    AND SC6.C6_CLI = SC9.C9_CLIENTE"
   cQryPV += "    AND SC6.C6_LOJA = SC9.C9_LOJA"
   cQryPV += "    AND SC9.D_E_L_E_T_ <> '*'"
      
   //Verifica Pedidos s? liberados
   cQryPV += "    AND ( (SC9.C9_BLEST = '  ' AND SC9.C9_BLCRED ='  ') AND (SC9.C9_BLWMS >= '05' OR SC9.C9_BLWMS = '  ')"+IIF(SC9->((FieldPos('C9_BLTMS') > 0))," AND SC9.C9_BLTMS = '  ')",")")
   cQryPV += "    AND SUBSTRING(SC6.C6_NOTA,1,2) = '  '"
      
   cQryPV += "     AND SC5.C5_NUM = '"+cNumSC5+"'"
    
   cQryPV += "     AND SC5.C5_NUM = SC6.C6_NUM"
   cQryPV += "     AND SC5.C5_CLIENTE = SC6.C6_CLI"
   cQryPV += "     AND SC5.C5_LOJACLI = SC6.C6_LOJA"
   
   cQryPV += "     AND SC5.C5_CLIENTE = SA1.A1_COD"
   cQryPV += "     AND SC5.C5_LOJACLI = SA1.A1_LOJA"
   cQryPV += "     AND SC6.C6_PRODUTO = SB1.B1_COD"
   cQryPV += "     AND SC6.C6_TES = SF4.F4_CODIGO"
      
   //S? apresentar Vendas e/ou Bonifica??o.
      
   cQryPV += "     AND SC6.C6_CF IN ('5101','6101','5102','6102','5910','6910')"
   cQryPV += "     AND SUBSTRING(SB1.B1_TIPO,1,2) = 'PA'"
      
   cQryPV += "     AND SC5.D_E_L_E_T_ <> '*'"
   cQryPV += "     AND SC6.D_E_L_E_T_ <> '*'"
   cQryPV += "     AND SA1.D_E_L_E_T_ <> '*'"
   cQryPV += "     AND SB1.D_E_L_E_T_ <> '*'"
   cQryPV += "     AND SF4.D_E_L_E_T_ <> '*'"
   
   TCQuery cQryPV NEW ALIAS "PVLB"
   
   DbSelectArea("PVLB")
   lPdLiber := IIF(PVLB->TOTREG > 0,.T.,.F.)
   DbCloseArea()
      

   DbSelectArea("SB2")
   DbSetOrder(1)
   ctBaB := Alias()
   cFilSB2 := xFilial("SB2")
   /*
   iF lEstoque
      mSGiNFO(" o TES "+cTESPed+" Atualiza Estoque")
   Else
      mSGiNFO(" o TES "+cTESPed+" NAO Atualiza Estoque")
   eNDiF
   MsgInfo("Tabela Aberta: "+ctBaB+" pROCURANDO Produto: "+cProduto+" Local "+cLocPrd+" fILIAL B2 "+cFilSB2)
   */
   If DbSeek(xFilial("SB2")+cProduto+cLocPrd)
       
      //Five Solutions Consultoria - 23/02/2010 - nSaldoPrd := SaldoSB2()
      //nSaldoPrd := SB2->B2_QATU - (IIF(lPdLiber,SB2->B2_RESERVA,SB2->B2_QPEDVEN))
      //nSaldoPrd := SB2->B2_QATU - (SB2->B2_RESERVA + SB2->B2_QPEDVEN)
      //nSaldoPrd := SB2->B2_QATU - (SB2->B2_RESERVA + SB2->B2_QPEDVEN) - nQtVend //Five Solutions - 24/03/2010 - Subtrair do saldo a quantidade digitada no pr?prio pedido.
      cQryQt := " SELECT C6_QTDVEN "
      cQryQt += "   FROM "+RetSQLName("SC6")
      cQryQt += "  WHERE C6_FILIAL = '"+xFilial("SC6")+"'"
      cQryQt += "    AND C6_NUM = '"+cNumSC5+"'"
      cQryQt += "    AND C6_ITEM = '"+cItemC6+"'"
      cQryQt += "    AND C6_PRODUTO = '"+cProduto+"'"
      cQryQt += "    AND C6_LOCAL = '"+cLocPrd+"'"
      cQryQt += "    AND D_E_L_E_T_ <> '*'"
      
      MemoWrite("QtAntProd.SQL",cQryQt)
      TCQuery cQryQt NEW ALIAS "XC6Q"
      TCSetField("XC6Q","C6_QTDVEN","N",15,05)
      DbSelectArea("XC6Q")
      nQtAntPr := XC6Q->C6_QTDVEN
      DbCloseArea()
      DbSelectArea("SB2")
      
      If lConfEtq //Five Solutions Consultoria - 31/03/2010
         //nSaldoPrd := (SB2->B2_QATU) - (SB2->B2_RESERVA + SB2->B2_QPEDVEN)  //Five Solutions - 25/03/2010 - Somar do saldo a quantidade digitada no pr?prio pedido.
         //nSaldoPrd := (SB2->B2_QATU + nQtVend) - (SB2->B2_RESERVA + SB2->B2_QPEDVEN)  //Five Solutions - 25/03/2010 - Somar do saldo a quantidade digitada no pr?prio pedido. 31/03/2010 - S? ocorrerem caso de inclus?o
         nSaldoPrd := (SB2->B2_QATU + nQtAntPr) - (SB2->B2_RESERVA + SB2->B2_QPEDVEN)  //Five Solutions - 25/03/2010 - Somar do saldo a quantidade digitada no pr?prio pedido. 31/03/2010 - S? ocorrerem caso de inclus?o
      Else 
         If INCLUI
            //nSaldoPrd := (SB2->B2_QATU + nQtVend) - (SB2->B2_RESERVA + SB2->B2_QPEDVEN)  //Five Solutions - 25/03/2010 - Somar do saldo a quantidade digitada no pr?prio pedido. 31/03/2010 - S? ocorrerem caso de inclus?o
            nSaldoPrd := (SB2->B2_QATU) - (SB2->B2_RESERVA + SB2->B2_QPEDVEN)  //Five Solutions - 31/03/2010
         Else
            //nSaldoPrd := (SB2->B2_QATU) - (SB2->B2_RESERVA + SB2->B2_QPEDVEN)  //Five Solutions - 31/03/2010
            //nSaldoPrd := (SB2->B2_QATU + nQtVend) - (SB2->B2_RESERVA + SB2->B2_QPEDVEN)  //Five Solutions - 25/03/2010 - Somar do saldo a quantidade digitada no pr?prio pedido. 31/03/2010 - S? ocorrerem caso de inclus?o
            nSaldoPrd := (SB2->B2_QATU + nQtAntPr) - (SB2->B2_RESERVA + SB2->B2_QPEDVEN)  //Five Solutions - 25/03/2010 - Somar do saldo a quantidade digitada no pr?prio pedido. 31/03/2010 - S? ocorrerem caso de inclus?o
         EndIf
      EndIf
      
      //nSaldoPrd := (SB2->B2_QATU + nQtVend) - (SB2->B2_RESERVA + SB2->B2_QPEDVEN)  //Five Solutions - 25/03/2010 - Somar do saldo a quantidade digitada no pr?prio pedido. 31/03/2010 - S? ocorrerem caso de inclus?o
       
      //MsgInfo("Achei sALDO DO Produto: "+SB2->B2_COD+" no Local "+SB2->B2_LOCAL+" B2_QATU "+Alltrim(Str(SB2->B2_QATU))+" B2_RESERVA "+Alltrim(Str(SB2->B2_RESERVA))+" B2_QPEDVEN "+Alltrim(Str(SB2->B2_QPEDVEN))+" A Qdt Digitada foi "+Alltrim(Str(nQtVend))+" Portando o Sld. Disponivel ?: "+Alltrim(Str(nSaldoPrd)))
      
   //Else
   //  MsgInfo("O Produto "+cProduto+" Local "+cLocPrd+" NAO FOI ENCONTRADO NO SB2 Filial "+cFilSB2)
   EndIf
   //mSGiNFO(" nQtdC6 "+aLLTRIM(Str(nQtdC6))+" nQtVend "+Alltrim(Str(nQtVend))+" nSaldoPrd "+Alltrim(Str(nSaldoPrd)))
   If nQtdC6 > nSaldoPrd .And. lEstoque
      lRetVld   := .F.
      Alert("*** A Quantidade informada "+Alltrim(Transform(nQtdC6,"@E 999,999,999.9999"))+" para o produto "+Alltrim(cNomePrd)+" ? maior que a quantidade dispon?vel no Armaz?m "+cLocPrd+" que ? de "+Alltrim(Transform(nSaldoPrd,"@E 999,999,999.9999"))+" Favor iformar uma quantidade menor. ***")
   EndIf
   RestArea(nAreaAki)
Return(lRetVld)