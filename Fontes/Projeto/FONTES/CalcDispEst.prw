
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �CALCDISPEST�Autor  �Five Solutions      � Data �  10/01/2008 ���
�������������������������������������������������������������������������͹��
���Desc.     � Valida��o de Usu�rio do campo C6_PRODUTO.                  ���
���          � Sistema far� consist�ncia entre a quantidade digitada no   ���
���          � pedido de pedido de vendas em rela��o a disponibilidade do ���
���          � saldo em estoque.                                          ���
�������������������������������������������������������������������������͹��
���Uso       � COMAFAL                                                   ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

User Function CalcDispEst
   Local lRetVld   := .T.
   Local nAreaAki  := GetArea()
   Local nQtdC6    := M->C6_QTDVEN
   Local cLocPrd   := GdFieldGet("C6_LOCAL",n)
   Local cProduto  := GdFieldGet("C6_PRODUTO",n)
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