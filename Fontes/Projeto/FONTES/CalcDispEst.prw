
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³CALCDISPESTºAutor  ³Five Solutions      º Data ³  10/01/2008 º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Validação de Usuário do campo C6_PRODUTO.                  º±±
±±º          ³ Sistema fará consistência entre a quantidade digitada no   º±±
±±º          ³ pedido de pedido de vendas em relação a disponibilidade do º±±
±±º          ³ saldo em estoque.                                          º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ COMAFAL                                                   º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
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
      Alert("A Quantidade informada "+Alltrim(Transform(nQtdC6,"@E 999,999,999.9999"))+" para o produto "+Alltrim(cNomePrd)+" é maior que a quantidade disponível no Armazém "+cLocPrd+" que é de "+Alltrim(Transform(nSaldoPrd,"@E 999,999,999.9999"))+" Favor iformar uma quantidade menor.")
   EndIf
   RestArea(nAreaAki)
Return(lRetVld)