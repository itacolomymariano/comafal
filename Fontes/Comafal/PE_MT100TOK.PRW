#include "TopConn.Ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณPE_MT100TOKบAutor  ณFive Solutions      บ Data ณ  15/09/2010 บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ PE utilizado para validar numero/sperie da NF consistindo  บฑฑ
ฑฑบ          ณ existencia de pesagem na entrada do controle de balan็a.   บฑฑ
ฑฑบ          ณ A NF deverแ estแ pr้-cadastrada na tabela SZB.             บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ COMAFAL PE,SP e RS - MATA103                              บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/

User Function MT100TOK
   
/*
PRIVATE cNFiscal	:= ""
PRIVATE cSerie		:= ""
PRIVATE cA100For	:= ""
PRIVATE cLoja		:= ""
PRIVATE cEspecie	:= ""
PRIVATE cCondicao	:= ""
PRIVATE cForAntNFE	:= ""
PRIVATE dDEmissao	:= dDataBase
*/
   lOkNFE := .T.
   
   cQrySZB := " SELECT COUNT(*) NQTDREG "
   cQrySZB += "   FROM "+RetSQLName("SZB")
   cQrySZB += "  WHERE ZB_FILIAL = '"+xFilial("SZB")+"'"
   cQrySZB += "    AND ZB_DOC = '"+cNFiscal+"'"       //Variaveis cNFiscal e cSerie sใo Privates do MATA103.
   cQrySZB += "    AND ZB_SERIE = '"+cSerie+"'"
   cQrySZB += "    AND D_E_L_E_T_ <> '*'"
   
   MemoWrite("C:\TEMP\QtdNFESZB.SQL",cQrySZB)
   
   TCQuery cQrySZB NEW ALIAS "XSZB"
   TCSetField("XSZB","NQTDREG","N",10,00)
   DbSelectArea("XSZB")
      nQuanSZB := XSZB->NQTDREG
   DbCloseArea()
   
   If nQuanSZB <= 0
      Alert("A Nota Fiscal/S้rie "+cNFiscal+" / "+cSerie+" nใo estแ registrada no Controle de Balan็a, nใo serแ permitido esta entrada ")
      Return(.F.)
   EndIf
   
Return(lOkNFE)