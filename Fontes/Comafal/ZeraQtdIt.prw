User Function ZeraQtdIt
   Local nAreaAki  := GetArea()
   Local cRetOper   := M->UA_OPER
   If M->UA_OPER <> "2"
      If Len(aCols) > 0
         Alert("Foi modificado o tipo de operação do pedido, as quantidade dos itens serão zeradas, favor digitar novamente")
      EndIf
      For nIt := 1 To Len(aCols)
         GdFieldPut("UB_QUANT",0,nIt)
      Next nIt
   EndIf
   RestArea(nAreaAki)
Return(cRetOper)