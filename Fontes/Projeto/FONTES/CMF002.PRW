#INCLUDE "RWMAKE.CH"
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³ CMF002   ³ Autor ³ Zanardo             ³ Data ³ 05.11.03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ AJUSTE DE CUSTO NA NF DE VENDA							³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Cliente   ³ Comafal                                                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function CMF002()

Local aArea  := GetArea()
Local cPerg		:= "CMF002"

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Verifica as perguntas selecionadas     ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
AjustaSX1(@cPerg)

Pergunte(cPerg,.T.)

Processa({|| CMF002A()}, "Aguarde... Processando..."  )

RestArea(aArea)

Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³ CMF002   ³ Autor ³ Zanardo             ³ Data ³ 05.11.03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ AJUSTE DE CUSTO NA NF DE VENDA							³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Cliente   ³ Comafal                                                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

STATIC Function CMF002A()

Local aArea  := GetArea()
Local aStrucSD2	:= SD2->(dbStruct()) 
Local aStrucSB1	:= SB1->(dbStruct())
Local cAliasSD2 := "cAliasSD2"
Local cQuery    := ""
Local nSD2 		:= 0 
Local nSB1 		:= 0

cQuery := "SELECT SD2.D2_FILIAL,SD2.D2_COD,SD2.D2_LOCAL,SD2.D2_NUMSEQ,SB1.B1_X_CST2 "
cQuery += "From " + RetSqlName("SD2")+ " SD2, " + RetSqlName("SB1")+" SB1 "
cQuery += "WHERE SD2.D_E_L_E_T_ <> '*' AND "
cQuery += "SD2.D2_EMISSAO BETWEEN '"+dtos(MV_PAR01)+"'  AND '"+dtos(MV_PAR02)+"' AND " 
cQuery += "SD2.D2_COD BETWEEN '"	+MV_PAR03+"'  AND '"+MV_PAR04+"' AND "
cQuery += "SD2.D2_DOC BETWEEN '"	+MV_PAR05+"'  AND '"+MV_PAR06+"' AND " 
cQuery += "SD2.D2_CLIENTE BETWEEN '"+MV_PAR07+"'  AND '"+MV_PAR08+"' AND "  
cQuery += "SD2.D2_TIPO = 'N' 										AND "
cQuery += "SD2.D2_FILIAL = '"+xFilial("SD2") + "' 					AND "
cQuery += "SD2.D2_FILIAL = SB1.B1_FILIAL 							AND "
cQuery += "SD2.D2_COD = SB1.B1_COD 									AND "
cQuery += "SB1.D_E_L_E_T_ <> '*' "
cQuery += "ORDER BY SD2.D2_FILIAL,SD2.D2_COD,SD2.D2_LOCAL,SD2.D2_NUMSEQ "

cQuery := ChangeQuery(cQuery)

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSD2,.F.,.T.)

For nSD2 := 1 To Len(aStrucSD2)
	If aStrucSD2[nSD2][2] <> "C" .And. FieldPos(aStrucSD2[nSD2][1]) <> 0
		TcSetField(cAliasSD2,aStrucSD2[nSD2][1],aStrucSD2[nSD2][2],aStrucSD2[nSD2][3],aStrucSD2[nSD2][4])
	EndIf
Next nSD2

For nSB1 := 1 To Len(aStrucSB1)
	If aStrucSB1[nSB1][2] <> "C" .And. FieldPos(aStrucSB1[nSB1][1]) <> 0
		TcSetField(cAliasSD2,aStrucSB1[nSB1][1],aStrucSB1[nSB1][2],aStrucSB1[nSB1][3],aStrucSB1[nSB1][4])
	EndIf
Next nSB1


DbSelectArea(cAliasSD2)
While (cAliasSD2)->(!Eof())
	DbSelectArea("SD2")
	DbSetOrder(1)
	If DbSeek((cAliasSD2)->D2_FILIAL+(cAliasSD2)->D2_COD+(cAliasSD2)->D2_LOCAL+(cAliasSD2)->D2_NUMSEQ)
		RecLock("SD2",.F.)
			SD2->D2_X_CST2 := SD2->D2_QUANT * Iif(MV_PAR09 = 0,(cAliasSD2)->B1_X_CST2,MV_PAR09)
		SD2->(MsUnlock())
	EndIf
	(cAliasSD2)->(DbSkip())
EndDo
(cAliasSD2)->(dbCloseArea())

RestArea(aArea)

Return
                  

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³AjustaSX1 ºAutor  ³Eduardo Zanardo     º Data ³ 28/04/06    º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Grava as perguntas no SX1                                   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³AP7                                                         º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/                         
STATIC FUNCTION AjustaSX1(cPerg)

Local aArea := GetArea()

cPerg := Padr(cPerg,6)

//PutSx1(cGrupo,cOrdem,cPergunt              ,"","",cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,"","",cHelp)
  PutSX1(cPerg ,"01"  ,"Da Emissao ?"        ,"","","mv_ch1","D"  ,8       ,0       ,0      ,"G"  ,""   ,""   ,""     ,"","mv_par01","","","","","","","",""    ,"","",""    ,"","",""    ,"","",""      ,"","","")
  PutSX1(cPerg,	"02"  ,"Ate Emissao ?"       ,"","","mv_ch2","D"  ,8       ,0       ,0      ,"G"  ,""   ,""   ,""     ,"","mv_par02","","","","","","","",""    ,"","",""    ,"","",""    ,"","",""      ,"","","")
  PutSX1(cPerg,	"03"  ,"Do Produto ?"        ,"","","mv_ch3","C"  ,15      ,0       ,0      ,"G" ,""    ,"SB1",""     ,"","mv_par03","","","","","","","",""    ,"","",""    ,"","",""    ,"","",""      ,"","","")	 
  PutSX1(cPerg,	"04"  ,"Ate Produto ?"       ,"","","mv_ch4","C"  ,15      ,0       ,0      ,"G" ,""    ,"SB1",""     ,"","mv_par04","","","","","","","",""    ,"","",""    ,"","",""    ,"","",""      ,"","","")	
  PutSX1(cPerg,	"05"  ,"Da Nota Fiscal ?"    ,"","","mv_ch5","C"  ,9       ,0       ,0      ,"G" ,""    ,""   ,""     ,"","mv_par05","","","","","","","",""    ,"","",""    ,"","",""    ,"","",""      ,"","","")	 
  PutSX1(cPerg,	"06"  ,"Ate a Nota Fiscal ?" ,"","","mv_ch6","C"  ,9       ,0       ,0      ,"G" ,""    ,""   ,""     ,"","mv_par06","","","","","","","",""    ,"","",""    ,"","",""    ,"","",""      ,"","","")	
  PutSX1(cPerg,	"07"  ,"Do Cliente ?"        ,"","","mv_ch7","C"  ,6       ,0       ,0      ,"G" ,""    ,"SA1",""     ,"","mv_par07","","","","","","","",""    ,"","",""    ,"","",""    ,"","",""      ,"","","")	 
  PutSX1(cPerg,	"08"  ,"Ate o Cliente ?"     ,"","","mv_ch8","C"  ,6       ,0       ,0      ,"G" ,""    ,"SA1",""     ,"","mv_par08","","","","","","","",""    ,"","",""    ,"","",""    ,"","",""      ,"","","")	 
  PutSX1(cPerg,	"09"  ,"Custo         ?"     ,"","","mv_ch9","N"  ,14      ,2       ,0      ,"G" ,""    ,""   ,""     ,"","mv_par09","","","","","","","",""    ,"","",""    ,"","",""    ,"","",""      ,"","","")	
RestArea(aArea)

Return
