
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �ChkQtdTp  �Autor  �Five Solutions      � Data �  07/05/2010 ���
�������������������������������������������������������������������������͹��
���Desc.     �Valida��o para permitir ou n�o edi��o do campo Quantidade   ���
���          �D1_QUANT no Documento de Entrada.                           ���
���          �S� ser� permitido Edi��o se o produto N�O for materia prima.���
�������������������������������������������������������������������������͹��
���Uso       � COMAFAL PE/RS/SP.                                         ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

User Function ChkQtdTp
   lAltQtd := .T.
   cProdNF := GdFieldGet("D1_COD",n)
   cTpPrd  := Posicione("SB1",1,xFilial("SB1")+cProdNF,"B1_TIPO")
   If cTpPrd == "MP"
      lAltQtd := .F.
   EndIf
Return(lAltQtd)

User Function CQtdTpPd
   lAltQtd := .T.
   cProdNF := GdFieldGet("C7_PRODUTO",n)
   cTpPrd  := Posicione("SB1",1,xFilial("SB1")+cProdNF,"B1_TIPO")
   If cTpPrd == "MP"
      lAltQtd := .F.
   EndIf
Return(lAltQtd)