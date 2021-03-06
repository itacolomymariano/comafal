
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �ClcCalDispE�Autor  �Five Solutions      � Data �  10/01/2008 ���
�������������������������������������������������������������������������͹��
���Desc.     � Valida��o de Usu�rio do campo UP_QUANT.                    ���
���          � Sistema far� consist�ncia entre a quantidade digitada no   ���
���          � pedido de pedido de vendas(Call Center) em rela��o a       ���
���          � disponibilidade do saldo em estoque.                       ���
�������������������������������������������������������������������������͹��
���Uso       � COMAFAL                                                   ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

User Function ClcCalDispE
   Local lRetVld   := .T.
   Local nAreaAki  := GetArea()
   Local nQtdC6    := M->UB_QUANT
   Local cLocPrd   := GdFieldGet("UB_LOCAL",n)
   Local cProduto  := GdFieldGet("UB_PRODUTO",n)
   Local cNomePrd  := Posicione("SB1",1,xFilial("SB1")+cProduto,"B1_DESC")
   Local nSaldoPrd := 0
   DbSelectArea("SB2")
   If DbSeek(xFilial("SB2")+cProduto+cLocPrd)
      nSaldoPrd := SaldoSB2()
   EndIf
   If nQtdC6 > nSaldoPrd
      lRetVld   := .F.
      Alert("A Quantidade informada "+Alltrim(Transform(nQtdC6,"@E 999,999,999.9999"))+" para o produto "+Alltrim(cNomePrd)+" � maior que a quantidade dispon�vel no Armaz�m "+cLocPrd+" que � de "+Alltrim(Transform(nSaldoPrd,"@E 999,999,999.9999"))+" Favor iformar uma quantidade menor.")
   EndIf
   RestArea(nAreaAki)
Return(lRetVld)