#INCLUDE "PROTHEUS.CH"  
#INCLUDE "topconn.ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³CRF010  ºAutor  ³Eduardo Zanardo     º Data ³  03/31/03   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ RESUMO GERENCIAL                   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP7                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function CRF010()

Local aOrd		:= {}
Local cDesc1    := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2    := "resumo gerencial, de acordo com os parametros informados pelo usuario."
Local cDesc3    := ""
Local cPict     := ""
Local titulo    := "RESUMO GERENCIAL " + SM0->M0_NOME + " " + SM0->M0_FILIAL
Local nLin      := 80
Local Cabec1    := ""
Local Cabec2    := ""
Local imprime   := .T.
Local aArea		:= GetArea()

Private CbTxt       := ""
Private lEnd        := .F.
Private lAbortPrint := .F.
Private limite     	:= 220
Private Tamanho    	:= "G"
Private nomeprog   	:= "CRF010"
Private nTipo      	:= 18
Private aReturn    	:= { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey   	:= 0
Private cPerg      	:= "CRF010"
Private cbtxt      	:= Space(10)
Private cbcont     	:= 00
Private CONTFL     	:= 01
Private m_pag      	:= 01
Private wnrel      	:= "CRF010"
Private cString    	:= "SC5"
Private	aL			:= {}                      
Private aMeses := {"Janeiro","Fevereiro","Marco","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"}  
Private aMes   := {"JAN1","FEV1","MAR1","ABR1","MAI1","JUN1","JUL1","AGO1","SET1","OUT1","NOV1","DEZ1"} 
Private aFatBruto := {}
Private aDevInt  := {}  
Private aDevExt  := {}
Private aSucata := {}

//SAIDAS
Private cCFOPVenda  := "" //vendas
Private cCFOPVdInt  := "" //vendas mercado interno
Private cCFOPVdExt  := "'7101','7102','7105','7106','7127'"     //venda mercado externo
Private cCFOPDvMc   := "'5556','6556'" //devolucao de COMPRA de Material de Consumo 
Private cCFOPDvMp   := "'5201','6201'" //devolucao de COMPRA de Material de Prima 
Private cCFOPDvAt   := "'5553','6553'" //devolucao de COMPRA de Ativo imobilizado
Private cCFOPTfMp	:= "'5151','6151'" //transf de materia prima 
Private cCFOPTfTc	:= "'5152','6152'" //transf de mercadoria adquirida e/ou recebida de terceiros 
Private cCFOPTAi	:= "'5552','6652'" //transf de ATIVO IMOBLIZADO
Private cCFOPBn     := "'5901','6901'" //REMESSA PARA BENEFICIAMENTO
Private cCFOPCs     := "'5915','6915'" //REMESSA PARA CONSERTO
Private cCFOPTr     := "'5924','6924'" //REMESSA OPERAÇÃO TRIANGULAR
Private cCFORd      := "'5913','6913'" //RETORNO DE DEMONSTRAÇÃO

//ENTRADAS              
Private cCFOPCpMP	:= "'1101','2101','3101'" //COMPRA DE MATERIA PRIMA
Private cCFOPCpCo	:= "'1102','2102','3102'" //COMPRA P/ COMERCIALIZAÇAO
Private cCFOPCpUc	:= "'1556','2556','3556'" //COMPRA P/ uso e consumo
Private cCFOPCpAt	:= "'1551','2551','3551'" //COMPRA P/ Ativo Imobilizado

Private cCFOPDvVd   := "'1201','2201','3201','1202','2202','3202'" //devolucao de venda

Private cCFOPCpEg   := "'1252','2252'" //Energia eletrica 
Private cCFOPCpTl   := "'1302','2302'" //Telefonia

Private cCFOPTfMpE   := "'1151','2151'"//transferencia de MP 
Private cCFOPTfMcE   := "'1152','2152'"//transferencia de Material de consumo 
Private cCFOPTfAtE   := "'1552','2552'"//transferencia de Ativo Imobilizado

Private cCFOPRtSv   := "'1124','2124'"//Retorno de Beneficiamento serviço cobrado
Private cCFOPRtBn   := "'1902','2902'"//Retorno de Beneficiamento
Private cCFOPRtCs   := "'1916','2916'"//Retorno de conserto

//NATUREZAS                                              
PRIVATE cNATCOMMT := ""
PRIVATE cNATADMGE := "" 
PRIVATE cNATADMMT := ""
PRIVATE cNATCOMMTEX := "" 
PRIVATE cNATGEEX := ""
PRIVATE cNATINDMT := "" 
PRIVATE cNATINDGE := ""

//COMERCIAL IN GERAL DA MITTO
PRIVATE cNATCOMGE := "'100201','100202','100203','100204','100205','100206','100207','100208','100209','100210',"
	    cNATCOMGE += "'100211','100212','100213','100214','100221','100222','200201','200202','200203','200204',"
	    cNATCOMGE += "'200205','200206','200207','200208','200211','200212'"

//COMERCIAL IN META DA MITTO
cNATCOMMT := "'100204','100205','100206','100207','100208','100209','100210','100212','100214','200201',"
cNATCOMMT += "'200202','200203','200204','200205','200206','200207','200212'"

// ADMINISTRATIVO GERAL DA MITTO
cNATADMGE := "'100101','100102','100103','100104','100105','100106','100107','100108','100109','100110',"
cNATADMGE += "'100111','100112','100113','100114','100115','100116','100117','100118','100119','100121',"
cNATADMGE += "'100122','100123','200101','200102','200103','200104','200105','200106','200107','200108',"
cNATADMGE += "'200109','200110','200111','200112','200113','200114','200115','200116','200117','200118',"
cNATADMGE += "'200119','200120','200121','200122','200123','200124','200126','200127','200129','200130',"
cNATADMGE += "'200131','200132','200133','200134','200135','200136','200137','200158','200159','200160',"
cNATADMGE += "'200501','200502','200503','200504','200505','200506','200507','200508','200509','200510',"
cNATADMGE += "'500101','500102','500103','500104','600302','700206'"     

//ADMINISTRATIVO META DA MITTO
cNATADMMT := "'100103','100104','100105','100106','100107','100108','100111','100112','100113','100114',"
cNATADMMT += "'100116','100117','100121','100123','200101','200102','200103','200104','200105','200106',"
cNATADMMT += "'200107','200108','200112','200113','200114','200115','200116','200117','200118','200121',"
cNATADMMT += "'200122','200123','200124','200126','200129','200130','200131','200134','200136','200137',"
cNATADMMT += "'200501','200502','200503','200504','200505','200506','200507','200508','200509','200510',"
cNATADMMT += "'500102','500103','500104','600302'"

//COMERCIAL EX META DA MITTO
cNATCOMMTEX := "'100223','100224','100225','100226','100227','100229','100231','100233','200213','200214',"
cNATCOMMTEX += "'200215','200216','200217','200218'" 

//COMERCIAL EX GERAL DA MITTO
cNATGEEX := "'100223','100224','100225','100226','100227','100228','100229','100230','100231','100232',"
cNATGEEX += "'100233','100234','100235','200213','200214','200215','200216','200217','200218'"

//INDUSTRIA META DA MITTO
cNATINDMT := "'100302','100303','100304','100305','100306','100307','100308','100311','100312','200301',"
cNATINDMT += "'200302','200303','200304','200305','200308','200309','200310','200311','200312','200314',"
cNATINDMT += "'200315','200316','200317','200318','200319','200320','200322','200323','200324','200325',"
cNATINDMT += "'200326','200327','200328','600301'"  


//INDUSTRIA GERAL DA MITTO
cNATINDGE := "'100301','100302','100303','100304','100305','100306','100307','100308','100309','100310',"
cNATINDGE += "'100311','100312','100313','100317','200301','200302','200303','200304','200305','200306',"
cNATINDGE += "'200307','200308','200309','200310','200311','200312','200313','200314','200315','200316',"
cNATINDGE += "'200317','200318','200319','200320','200321','200322','200323','200324','200325','200326',"
cNATINDGE += "'200327','200328','200329','200361','200362','200363','200365','200367','600301','700303',"
cNATINDGE += "'700309','700310','700311'"



cCFOPVenda  	    := "'5101','6101','5102','6102','5103','6103','5104','6104','5105','6105',"
cCFOPVenda  	    += "'5106','6106','6107','6108','5109','6109','5110','6110','5111','6111','5112','6112',"
cCFOPVenda  	    += "'5113','6113','5114','6114','5115','6115','5116','6116','5117','6117','5118','6118','5119','6119',"
cCFOPVenda  	    += "'5120','6120','5122','6122','5123','6123','7101','7102','7105','7106','7127'"

cCFOPVdInt  	    := "'5101','6101','5102','6102','5103','6103','5104','6104','5105','6105',"
cCFOPVdInt  	    += "'5106','6106','6107','6108','5109','6109','5110','6110','5111','6111','5112','6112',"
cCFOPVdInt  	    += "'5113','6113','5114','6114','5115','6115','5116','6116','5117','6117','5118','6118','5119','6119',"
cCFOPVdInt  	    += "'5120','6120','5122','6122','5123','6123'"

dbSelectArea(cString)
dbSetOrder(1)
AjustaSx1()
Pergunte(cPerg,.F.)

wnrel := SetPrint(cString,NomeProg,cPerg,@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,.T.,Tamanho,,.F.)

If nLastKey == 27
	Return
Endif

SetDefault(aReturn,cString)

If nLastKey == 27
	Return
Endif

titulo += " 01/"+strzero(month(MV_PAR01),2)+"/" + substr(DTOS(MV_PAR01),1,4)+"/"+ " ate " + substr(DTOS(MV_PAR02),7,2)+"/"+substr(DTOS(MV_PAR02),5,2)+"/" + substr(DTOS(MV_PAR02),1,4)+"/"

RptStatus({|| U_CRF010B(Cabec1,Cabec2,Titulo,@nLin) },Titulo)
Eject
SET DEVICE TO SCREEN
If aReturn[5] == 1
	dbCommitAll()
	SET PRINTER TO
	OurSpool(wnrel)
Endif
MS_FLUSH()

RestArea(aArea)

Return


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³AJUSTASX1 ºAutor  ³Eduardo Zanardo     º Data ³  19/03/03   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Ajusta o SX1                                                º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AFATR001                                                   º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function AjustaSx1()

Local aArea := GetArea()

//PutSx1(cGrupo,cOrdem,cPergunt  ,"","",cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
  PutSx1(cPerg ,"01","Data de ?" ,"","","mv_ch1","D"  ,8       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR01",""     ,"","",""    ,""           ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"02","Data ate ?","","","mv_ch2","D"  ,8       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR02",""     ,"","",""    ,""           ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"03","Custo ?"   ,"","","mv_ch3","C"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR03","Medio","","",""    ,"Reposicao"  ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"04","Analitico?","","","mv_ch4","C"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR04","Sim"  ,"","",""    ,"Nao"        ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"05","Depto.?"   ,"","","mv_ch5","C"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria","","","Financeiro","","","","","")

RestArea(aArea)

Return(.T.)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³AJUSTASX1 ºAutor  ³Eduardo Zanardo     º Data ³  19/03/03   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Ajusta o SX1                                                º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AFATR001                                                   º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function Pedidos(aPedido,nQtdAtivo,nMes)

Local cQuery 	:= ""
Local cAliasPED := "cAliasPED"
Local nSC5	 	:= 0 
Local nSC6	 	:= 0
Local aStruSC5 	:= SC5->(dbStruct())
Local aStruSC6 	:= SC6->(dbStruct()) 
Local cAliasSA1 := "cAliasSA1"
Local nSA1	 	:= 0
Local aStruSA1 	:= SA1->(dbStruct()) 

cQuery := "SELECT SUBSTRING(C5.C5_EMISSAO,1,6) as MES,"
cQuery += "SUM(C6.C6_VALOR) as PedBrt, "
cQuery += "COUNT(DISTINCT C5.C5_NUM) as QtdPED "
cQuery += " FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 "
cQuery += "WHERE "
cQuery += "C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += "C5.C5_TIPO = 'N' AND "
cQuery += "C5.D_E_L_E_T_ = ' ' AND "
cQuery += "C6.C6_FILIAL = '"+xFilial("SC6")+"' AND "
cQuery += "C6.C6_NUM = C5.C5_NUM AND "
cQuery += "SUBSTRING(C5.C5_EMISSAO,1,6) >= '"+ Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +  "' AND "
cQuery += "SUBSTRING(C5.C5_EMISSAO,1,6) <= '"+ SUBSTR(DTOS(MV_PAR02),1,6) +"' AND "
cQuery += "C6.C6_LOCAL IN ('03','04') AND "
cQuery += "C6.D_E_L_E_T_ = ' ' "          
cQuery += "GROUP BY SUBSTRING(C5.C5_EMISSAO,1,6) "
cQuery += "ORDER BY SUBSTRING(C5.C5_EMISSAO,1,6) ASC"

cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasPED,.T.,.T.)},"Aguarde...","Processando Pedidos Emitidos...")

For nSC5 := 1 To len(aStruSC5)
	If aStruSC5[nSC5][2] <> "C" .And. FieldPos(aStruSC5[nSC5][1])<>0
		TcSetField(cAliasPED,aStruSC5[nSC5][1],aStruSC5[nSC5][2],aStruSC5[nSC5][3],aStruSC5[nSC5][4])
	EndIf
Next nSC5

For nSC6 := 1 To len(aStruSC6)
	If aStruSC6[nSC6][2] <> "C" .And. FieldPos(aStruSC6[nSC6][1])<>0
		TcSetField(cAliasPED,aStruSC6[nSC6][1],aStruSC6[nSC6][2],aStruSC6[nSC6][3],aStruSC6[nSC6][4])
	EndIf
Next nSC6

dbSelectArea(cAliasPED)

aadd(aPedido,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0,0,0})
aadd(aPedido,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0,0,0})
aadd(aPedido,{SUBSTR(DTOS(MV_PAR02),1,6),0,0,0,0})

while (cAliasPED)->(!eof())
	nReg := aScan(aPedido,{|x| x[1]==(cAliasPED)->MES})
	If nReg > 0
		aPedido[nReg,2] :=(cAliasPED)->QtdPED
		aPedido[nReg,3] :=(cAliasPED)->PedBrt
	EndIf
	(cAliasPED)->(dbskip())
enddo
DbSelectArea(cAliasPED)
(cAliasPED)->(DbCloseArea())


cQuery := "SELECT SUBSTRING(C5.C5_EMISSAO,1,6) as MES,"
cQuery += "COUNT(DISTINCT C5.C5_CLIENTE) as CLIENTE "
cQuery += " FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 "
cQuery += "WHERE "
cQuery += "C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += "C5.C5_TIPO = 'N' AND "
cQuery += "C5.D_E_L_E_T_ = ' ' AND "
cQuery += "C6.C6_FILIAL = '"+xFilial("SC6")+"' AND "
cQuery += "C6.C6_NUM = C5.C5_NUM AND "
cQuery += "SUBSTRING(C5.C5_EMISSAO,1,6) >= '"+ Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND "
cQuery += "SUBSTRING(C5.C5_EMISSAO,1,6) <= '"+ SUBSTR(DTOS(MV_PAR02),1,6) +"' AND "
cQuery += "C6.C6_LOCAL IN ('03','04') AND "
cQuery += "C6.D_E_L_E_T_ = ' ' "          
cQuery += "GROUP BY SUBSTRING(C5.C5_EMISSAO,1,6) "
cQuery += "ORDER BY SUBSTRING(C5.C5_EMISSAO,1,6) ASC"

cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasPED,.T.,.T.)},"Aguarde...","Processando Pedidos Emitidos...")

For nSC5 := 1 To len(aStruSC5)
	If aStruSC5[nSC5][2] <> "C" .And. FieldPos(aStruSC5[nSC5][1])<>0
		TcSetField(cAliasPED,aStruSC5[nSC5][1],aStruSC5[nSC5][2],aStruSC5[nSC5][3],aStruSC5[nSC5][4])
	EndIf
Next nSC5

For nSC6 := 1 To len(aStruSC6)
	If aStruSC6[nSC6][2] <> "C" .And. FieldPos(aStruSC6[nSC6][1])<>0
		TcSetField(cAliasPED,aStruSC6[nSC6][1],aStruSC6[nSC6][2],aStruSC6[nSC6][3],aStruSC6[nSC6][4])
	EndIf
Next nSC6

dbSelectArea(cAliasPED)

while (cAliasPED)->(!eof())
	nReg := aScan(aPedido,{|x| x[1]==(cAliasPED)->MES})
	If nReg > 0
		aPedido[nReg,04] := (cAliasPED)->CLIENTE
	EndIf	
	(cAliasPED)->(dbskip())
enddo

DbSelectArea(cAliasPED)
(cAliasPED)->(DbCloseArea())
cQuery := " SELECT SUBSTRING(A1_DTCAD,1,6) AS NOVOS,COUNT(*) AS QTD "
cQuery += " FROM "
cQuery += RetSqlName('SA1')
cQuery += " WHERE "
cQuery += " A1_FILIAL = '"+xFilial("SA1")+"' AND "
cQuery += " D_E_L_E_T_ = ' ' AND "     
cQuery += "SUBSTRING(A1_DTCAD,1,6) >= '"+ Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND "
cQuery += "SUBSTRING(A1_DTCAD,1,6) <= '"+ SUBSTR(DTOS(MV_PAR02),1,6) +"'  "
cQuery += "GROUP BY SUBSTRING(A1_DTCAD,1,6) "
cQuery += "ORDER BY SUBSTRING(A1_DTCAD,1,6) ASC"

cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSA1,.T.,.T.)},"Aguarde...","Processando Clientes...")

For nSA1 := 1 To len(aStruSA1)
	If aStruSA1[nSA1][2] <> "C" .And. FieldPos(aStruSA1[nSA1][1])<>0
		TcSetField(cAliasSA1,aStruSA1[nSA1][1],aStruSA1[nSA1][2],aStruSA1[nSA1][3],aStruSA1[nSA1][4])
	EndIf
Next nSA1

dbSelectArea(cAliasSA1)

while (cAliasSA1)->(!eof())
	nReg := aScan(aPedido,{|x| x[1]==(cAliasSA1)->NOVOS})
	If nReg > 0
		aPedido[nReg,05] := (cAliasSA1)->QTD
	EndIf	
	(cAliasSA1)->(dbskip())
enddo

DbSelectArea(cAliasSA1)
(cAliasSA1)->(DbCloseArea())

If len(aPedido)  <= 0 
	aadd(aPedido,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0,0,0}) 
	aadd(aPedido,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0,0,0})
	aadd(aPedido,{SUBSTR(DTOS(MV_PAR02),1,6),0,0,0,0})
EndIf


Return(.t.)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³FatFin    ºAutor  ³Eduardo Zanardo     º Data ³  19/03/03   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³FatFin                                                      º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ crf010                                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/


User Function FatFin(aFatFin,nMes)   

Local cALiasSD2 := "cALiasSD2"
Local nSD2		:= 0
Local aStruSD2  := SD2->(dbStruct())
Local cQuery 	:= ""	
Local aArea		:= GetArea()
Local nSA1		:= 0
Local aStruSA1  := SA1->(dbStruct())
Local nSB1		:= 0
Local aStruSB1  := SB1->(dbStruct())

//aL[23]	:=		"|Faturamento que nao entraram no Financeiro                                                                                                                                                            |"
//aL[24]	:=		"|Cliente| Loja | Nome                                   |Nota Fiscal | Serie | Produto       | Descricao                                         | QTD           | Armazem | Valor                     |"		
//aL[25]	:=		"|###### | ##   | ###################################### | ######     |   ### | ##############| ################################################# |###############|   ##    | ###############           |"

cQuery := " SELECT D2_CLIENTE,D2_LOJA,A1_NREDUZ,D2_DOC,D2_SERIE,D2_COD,B1_DESC,D2_QUANT,D2_LOCAL,SUM(D2_TOTAL+D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA) AS VALOR "
cQuery += " FROM "         
cQuery += RetSqlName('SD2') + " D2, "
cQuery += RetSqlName('SA1') + " A1, "
cQuery += RetSqlName('SB1') + " B1 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND SUBSTRING(D2_EMISSAO,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "'  "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ' "
cQuery += " AND D2.D2_DOC+D2.D2_SERIE NOT IN (SELECT E1_NUM+E1_PREFIXO "
cQuery += " FROM "
cQuery += RetSqlName('SE1')
cQuery += " WHERE "
cQuery += " D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND E1_ORIGEM = 'MATA460') "
cQuery += " AND D2.D2_CLIENTE+D2.D2_LOJA = A1.A1_COD+A1.A1_LOJA "
cQuery += " AND A1.D_E_L_E_T_ = ' ' "
cQuery += " AND D2.D2_COD = B1.B1_COD " 
cQuery += " AND B1.B1_FILIAL = '"+xFilial("SB1")+"' "
cQuery += " AND B1.D_E_L_E_T_ = ' ' "
cQuery += " GROUP BY D2_CLIENTE,D2_LOJA,A1_NREDUZ,D2_DOC,D2_SERIE,D2_COD,B1_DESC,D2_QUANT,D2_LOCAL "
cQuery += " ORDER BY D2_CLIENTE,D2_LOJA,A1_NREDUZ,D2_DOC,D2_SERIE,D2_COD,B1_DESC,D2_QUANT,D2_LOCAL "


cQuery:=ChangeQuery(cQuery)
	
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSD2,.T.,.T.)},"Aguarde...","Processando Faturamentos que nao geraram Financeiro")

For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField(cAliasSD2,aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2

For nSA1 := 1 To len(aStruSA1)
	If aStruSA1[nSA1][2] <> "C" .And. FieldPos(aStruSA1[nSA1][1])<>0
		TcSetField(cAliasSD2,aStruSA1[nSA1][1],aStruSA1[nSA1][2],aStruSA1[nSA1][3],aStruSA1[nSA1][4])
	EndIf
Next nSA1

For nSB1 := 1 To len(aStruSB1)
	If aStruSB1[nSB1][2] <> "C" .And. FieldPos(aStruSB1[nSB1][1])<>0
		TcSetField(cAliasSD2,aStruSB1[nSB1][1],aStruSB1[nSB1][2],aStruSB1[nSB1][3],aStruSB1[nSB1][4])
	EndIf
Next nSB1

	
DbSelectArea(cAliasSD2)
Do While (cAliasSD2)->(!EOF())
	aadd(aFatFin,{(cAliasSD2)->D2_CLIENTE,(cAliasSD2)->D2_LOJA,(cAliasSD2)->A1_NREDUZ,(cAliasSD2)->D2_DOC,(cAliasSD2)->D2_SERIE,;
	(cAliasSD2)->D2_COD,(cAliasSD2)->B1_DESC,(cAliasSD2)->D2_QUANT,(cAliasSD2)->D2_LOCAL,(cAliasSD2)->VALOR})
	(cAliasSD2)->(DbSkip())
EndDo
DbSelectArea(cAliasSD2)
dbCloseArea()

RestArea(aArea)


Return(.T.)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³FatEst    ºAutor  ³Eduardo Zanardo     º Data ³  19/03/03   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³FatEst                                                      º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ crf010                                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function FatEst(aFatEst,nMes)   

Local cALiasSD2 := "cALiasSD2"
Local nSD2		:= 0
Local aStruSD2  := SD2->(dbStruct())
Local cQuery 	:= ""	
Local aArea		:= GetArea()
Local nSA1		:= 0
Local aStruSA1  := SA1->(dbStruct())
Local nSB1		:= 0
Local aStruSB1  := SB1->(dbStruct())

cQuery := " SELECT D2_CLIENTE,D2_LOJA,A1_NREDUZ,D2_DOC,D2_SERIE,D2_COD,B1_DESC,D2_QUANT,D2_LOCAL,SUM(D2_TOTAL+D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA) AS VALOR "
cQuery += " FROM "         
cQuery += RetSqlName('SD2') + " D2, "
cQuery += RetSqlName('SA1') + " A1, "
cQuery += RetSqlName('SB1') + " B1 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND SUBSTRING(D2_EMISSAO,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "'  "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ' "
cQuery += " AND D2.D2_TES IN (SELECT F4_CODIGO FROM "
cQuery +=  RetSqlName('SF4') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND F4_TIPO = 'S' AND F4_ESTOQUE = 'N') "
cQuery += " AND D2.D2_CLIENTE+D2.D2_LOJA = A1.A1_COD+A1.A1_LOJA "
cQuery += " AND A1.D_E_L_E_T_ = ' ' "
cQuery += " AND D2.D2_COD = B1.B1_COD " 
cQuery += " AND B1.B1_FILIAL = '"+xFilial("SB1")+"' "
cQuery += " AND B1.D_E_L_E_T_ = ' ' "
cQuery += " GROUP BY D2_CLIENTE,D2_LOJA,A1_NREDUZ,D2_DOC,D2_SERIE,D2_COD,B1_DESC,D2_QUANT,D2_LOCAL "
cQuery += " ORDER BY D2_CLIENTE,D2_LOJA,A1_NREDUZ,D2_DOC,D2_SERIE,D2_COD,B1_DESC,D2_QUANT,D2_LOCAL "


cQuery:=ChangeQuery(cQuery)
	
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSD2,.T.,.T.)},"Aguarde...","Processando Faturamentos que nao baixaram Estoque")

For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField(cAliasSD2,aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2

For nSA1 := 1 To len(aStruSA1)
	If aStruSA1[nSA1][2] <> "C" .And. FieldPos(aStruSA1[nSA1][1])<>0
		TcSetField(cAliasSD2,aStruSA1[nSA1][1],aStruSA1[nSA1][2],aStruSA1[nSA1][3],aStruSA1[nSA1][4])
	EndIf
Next nSA1

For nSB1 := 1 To len(aStruSB1)
	If aStruSB1[nSB1][2] <> "C" .And. FieldPos(aStruSB1[nSB1][1])<>0
		TcSetField(cAliasSD2,aStruSB1[nSB1][1],aStruSB1[nSB1][2],aStruSB1[nSB1][3],aStruSB1[nSB1][4])
	EndIf
Next nSB1
DbSelectArea(cAliasSD2)
Do While (cAliasSD2)->(!EOF())
	aadd(aFatEst,{(cAliasSD2)->D2_CLIENTE,(cAliasSD2)->D2_LOJA,(cAliasSD2)->A1_NREDUZ,(cAliasSD2)->D2_DOC,(cAliasSD2)->D2_SERIE,;
	(cAliasSD2)->D2_COD,(cAliasSD2)->B1_DESC,(cAliasSD2)->D2_QUANT,(cAliasSD2)->D2_LOCAL,(cAliasSD2)->VALOR})
	(cAliasSD2)->(DbSkip())
EndDo
DbSelectArea(cAliasSD2)
dbCloseArea()
RestArea(aArea)


Return(.T.)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³FatEst    ºAutor  ³Eduardo Zanardo     º Data ³  19/03/03   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³FatEst                                                      º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ crf010                                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function Inativos(cData)   

Local nQtdDest  := 0
Local nSA1		:= 0
Local aStruSA1  := SA1->(dbStruct())
Local cAliasSA1 := "cAliasSA1"
Local cQuery    := "" 

cQuery := " SELECT COUNT(*) AS GERAL "
cQuery += " FROM "
cQuery +=   RetSqlName('SA1') + " A1  "
cQuery += " WHERE "
cQuery += " A1.A1_FILIAL = '"+xFilial("SA1")+"' AND "  
cQuery += " A1_ULTCOM < '"+ cData +"' AND "
cQuery += " A1.D_E_L_E_T_ = ' '  "

cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSA1,.T.,.T.)},"Aguarde...","Processando Clientes Inativos...")

For nSA1 := 1 To len(aStruSA1)
	If aStruSA1[nSA1][2] <> "C" .And. FieldPos(aStruSA1[nSA1][1])<>0
		TcSetField(cAliasSA1,aStruSA1[nSA1][1],aStruSA1[nSA1][2],aStruSA1[nSA1][3],aStruSA1[nSA1][4])
	EndIf
Next nSA1

dbSelectArea(cAliasSA1)

while (cAliasSA1)->(!eof())
	nQtdDest := (cAliasSA1)->GERAL
	(cAliasSA1)->(dbskip())
enddo

DbSelectArea(cAliasSA1)
(cAliasSA1)->(DbCloseArea())

Return(nQtdDest)
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³FatEst    ºAutor  ³Eduardo Zanardo     º Data ³  19/03/03   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³FatEst                                                      º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ crf010                                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function Ativos(cData)   

Local nQtdAtivo := 0
Local nSA1		:= 0
Local aStruSA1  := SA1->(dbStruct())
Local cAliasSA1 := "cAliasSA1"
Local cQuery    := "" 

cQuery := " SELECT COUNT(*) AS GERAL "
cQuery += " FROM "
cQuery +=   RetSqlName('SA1') + " A1  "
cQuery += " WHERE "
cQuery += " A1_ULTCOM >= '"+ cData +"' AND "
cQuery += " A1.D_E_L_E_T_ = ' '  "

cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSA1,.T.,.T.)},"Aguarde...","Processando Clientes Ativos...")

For nSA1 := 1 To len(aStruSA1)
	If aStruSA1[nSA1][2] <> "C" .And. FieldPos(aStruSA1[nSA1][1])<>0
		TcSetField(cAliasSA1,aStruSA1[nSA1][1],aStruSA1[nSA1][2],aStruSA1[nSA1][3],aStruSA1[nSA1][4])
	EndIf
Next nSA1

dbSelectArea(cAliasSA1)

while (cAliasSA1)->(!eof())
	nQtdAtivo := (cAliasSA1)->GERAL
	(cAliasSA1)->(dbskip())
enddo

DbSelectArea(cAliasSA1)
(cAliasSA1)->(DbCloseArea())

Return(nQtdAtivo)
