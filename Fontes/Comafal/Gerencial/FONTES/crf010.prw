#INCLUDE "PROTHEUS.CH"  
#INCLUDE "topconn.ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณCRF010  บAutor  ณEduardo Zanardo     บ Data ณ  03/31/03   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ RESUMO GERENCIAL                   บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ AP7                                                        บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
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
Local nMes 		:= 0

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
Private aSucata	:= {}
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
Private cCFOPTr     := "'5924','6924'" //REMESSA OPERAวรO TRIANGULAR
Private cCFORd      := "'5913','6913'" //RETORNO DE DEMONSTRAวรO

//ENTRADAS              
Private cCFOPCpMP	:= "'1101','2101','3101'" //COMPRA DE MATERIA PRIMA
Private cCFOPCpCo	:= "'1102','2102','3102'" //COMPRA P/ COMERCIALIZAวAO
Private cCFOPCpUc	:= "'1556','2556','3556'" //COMPRA P/ uso e consumo
Private cCFOPCpAt	:= "'1551','2551','3551'" //COMPRA P/ Ativo Imobilizado

Private cCFOPDvVd   := "'1201','2201','3201','1202','2202','3202'" //devolucao de venda

Private cCFOPCpEg   := "'1252','2252'" //Energia eletrica 
Private cCFOPCpTl   := "'1302','2302'" //Telefonia

Private cCFOPTfMpE   := "'1151','2151'"//transferencia de MP 
Private cCFOPTfMcE   := "'1152','2152'"//transferencia de Material de consumo 
Private cCFOPTfAtE   := "'1552','2552'"//transferencia de Ativo Imobilizado

Private cCFOPRtSv   := "'1124','2124'"//Retorno de Beneficiamento servi็o cobrado
Private cCFOPRtBn   := "'1902','2902'"//Retorno de Beneficiamento
Private cCFOPRtCs   := "'1916','2916'"//Retorno de conserto

//NATUREZAS                                              
PRIVATE cNATCOMMT := ""
PRIVATE cNATADMGE := "" 
PRIVATE cNATADMMT := ""
PRIVATE cNATCOMMTEX := "" 
PRIVATE cNATGEEX  := ""
PRIVATE cNATINDMT := "" 
PRIVATE cNATINDGE := ""
PRIVATE cCUSTPROD := "" 
PRIVATE cGASTGERIND := ""
PRIVATE aPRODUZIDO := {}
PRIVATE aCSTBruto := {}
//COMERCIAL IN GERAL DA MITTO
PRIVATE cNATCOMGE := "'100201','100202','100203','100204','100205','100206','100207','100208','100209','100210',"
	    cNATCOMGE += "'100211','100212','100213','100214','100221','100222','200201','200202','200203','200204',"
	    cNATCOMGE += "'200205','200206','200207','200208','200211','200212'"

//NATUREZAS PARA CUSTO DE PRODUCAO
cCUSTPROD := "'200308','200408','100301','100302','100303','100304','100305','100306','100307','100309','100310',"
cCUSTPROD += "'100311','100312','100313','100317','100401','100402','100403','100404','100405','100406','100407','100409','100410',"
cCUSTPROD += "'100411','100412','100417','200314','600601','200317','200318','200319','200417','200418','200419','200320','200420',"
cCUSTPROD += "'200311','200411','200310','200410','200309'"
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

//GASTOS GERAIS IND.
cGASTGERIND := "'100301','100302','100303','100304','100305','100306','100307','100308','100309','100310','100311','100312','100313','100317','100401','100402','100403','100404','100405','100406','100407',"
cGASTGERIND += "'100408','100409','100410','100411','100412','100417','200109','200110','200111','200132','200135','200301','200302','200303','200304','200305','200306','200307','200308','200309',"
cGASTGERIND += "'200310','200311','200312','200313','200314','200315','200316','200317','200318','200319','200320','200321','200322','200360','200361','200362','200363','200364','200365','200366',"
cGASTGERIND += "'200367','200368','200369','200370','200372','200373','200374','200375','200376','200377','200401','200402','200403','200404','200405','200406','200407','200408','200410','200411',"
cGASTGERIND += "'200412','200416','200417','200418','200419','200420','200421','600201','600301','600501','600601','700301','700303','700304','700306','700307','700308','700309'"
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

AjustaSx1()
Pergunte(cPerg,.F.)
wnrel := SetPrint(cString,NomeProg,cPerg,@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,.T.,Tamanho,,.F.)

If nLastKey == 27
	Return
Endif

titulo += " 01/"+strzero(month(MV_PAR01),2)+"/" + substr(DTOS(MV_PAR01),1,4)+"/"+ " ate " + substr(DTOS(MV_PAR02),7,2)+"/"+substr(DTOS(MV_PAR02),5,2)+"/" + substr(DTOS(MV_PAR02),1,4)+"/"

DbSelectArea("SC5")
DbSetOrder(2)
SC5->(MsSeek(xFilial('SC5')+Dtos(MV_PAR01),.T.))
while SC5->(!EOF() .and. C5_FILIAL == xFilial('SC5') .and. C5_EMISSAO <= MV_PAR02 )
	//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
	//ณRecalculo da margem de lucro do Pedidoณ
	//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู
	If SC5->C5_TIPO = "N"
		DbSelectArea("SC6")
		DbSetOrder(1)
		If DbSeek(xFilial("SC6")+SC5->C5_NUM)
			Do While SC6->(!EOF()) .AND. SC6->C6_NUM == SC5->C5_NUM .AND. xFilial("SC6") == SC6->C6_FILIAL
				DbSelectArea("SB1")
				DbSetOrder(1)
				DbSeek(xFilial("SB1")+SC6->C6_PRODUTO)
				nPrecoVen := xMoeda(SC6->C6_PRCVEN,SC5->C5_MOEDA,1,SC5->C5_EMISSAO)
				RecLock("SC6",.F.)
				If SC6->C6_X_CUSTO == 0
					SC6->C6_X_CUSTO := SB1->B1_CUSTD  
				EndIf	
				If NoRound(((nPrecoVen/SC6->C6_X_CUSTO)-1)*100,2) > 999
					SC6->C6_X_MARGE := 999
				Else	
					SC6->C6_X_MARGE := NoRound(((nPrecoVen/SC6->C6_X_CUSTO)-1)*100,2)
				EndIf	
				MsUnlock()
			    DbSelectArea("SC6")
				SC6->(DbSkip())	
			EndDo                  
		EndIf	         
	EndIf	
	DbSelectArea("SC5")
	SC5->(DbSkip())	
EndDo 
nMes := val(Substr(dtos(MV_PAR01),5,2))
For nMes:=nMes to val(Substr(dtos(MV_PAR02),5,2))
	U_CMB007(nMes)
Next nMes
nMes := val(Substr(dtos(MV_PAR02),5,2))
aadd(aPRODUZIDO,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0}) 
aadd(aPRODUZIDO,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0})
aadd(aPRODUZIDO,{SUBSTR(DTOS(MV_PAR02),1,6),0,0})
If ( nLastKey == 27 )
	dbSelectArea(cString)
	dbSetOrder(1)
	Return
Endif
SetDefault(aReturn,cString)
If ( nLastKey == 27 )
	dbSelectArea(cString)
	dbSetOrder(1)
	Return
Endif

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
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณAJUSTASX1 บAutor  ณEduardo Zanardo     บ Data ณ  19/03/03   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณAjusta o SX1                                                บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ AFATR001                                                   บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function AjustaSx1()

Local aArea := GetArea()

//PutSx1(cGrupo,cOrdem,cPergunt  ,"","",cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
  PutSx1(cPerg ,"01","Data de ?" ,"","","mv_ch1","D"  ,8       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR01",""     ,"","",""    ,""           ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"02","Data ate ?","","","mv_ch2","D"  ,8       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR02",""     ,"","",""    ,""           ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"03","Custo ?"   ,"","","mv_ch3","C"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR03","Medio","","",""    ,"Reposicao"  ,"","","Calculado","","",""          ,"","","","","")
  PutSx1(cPerg ,"04","Analitico?","","","mv_ch4","C"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR04","Sim"  ,"","",""    ,"Nao"        ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"05","Depto.?"   ,"","","mv_ch5","C"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria","","","Financeiro","","","Gerencial","","")

RestArea(aArea)

Return(.T.)

