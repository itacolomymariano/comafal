#include "TopConn.Ch"
#include "Color.Ch"
#include "Colors.Ch"

/*
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????ͻ??
???Programa  ?RGERANLT  ?Autor  ?Five Solutions      ? Data ?  06/06/2010 ???
?????????????????????????????????????????????????????????????????????????͹??
???Desc.     ? Relatorio Gerencial Analitico - Vis?es Comercial, Industrial??
???          ? Administrativo e Resumo de Despesas/Receitas.              ???
???          ? Relat?rio Desenvolvido em Modelo Grafico: TReport          ???
?????????????????????????????????????????????????????????????????????????͹??
???Uso       ? COMAFAL - PE,SP e RS.                                     ???
?????????????????????????????????????????????????????????????????????????ͼ??
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
*/

User Function RGerAnlt
   
   Private aCampos   := {}
           aAdd(aCampos, {"INDICADO","C",50,00})
           aAdd(aCampos, {"SUBINDIC","C",50,00})
           aAdd(aCampos, {"INTRAIND","C",50,00})
           aAdd(aCampos, {"E5_FILIAL","C",04,00})
           aAdd(aCampos, {"MES","C",12,00})
           aAdd(aCampos, {"E5_NATUREZ","C",10,00})
           aAdd(aCampos, {"ED_DESCRIC","C",30,00})
           aAdd(aCampos, {"E5_PREFIXO","C",03,00})
           aAdd(aCampos, {"E5_NUMERO","C",09,00})
           aAdd(aCampos, {"E5_PARCELA","C",01,00})
           aAdd(aCampos, {"E5_TIPO","C",03,00})
           aAdd(aCampos, {"E5_BENEF","C",30,00})
           aAdd(aCampos, {"E5_HISTOR","C",40,00})
           aAdd(aCampos, {"E5_DATA","D",08,00})
           aAdd(aCampos, {"E5_VALOR","N",17,02})
           aAdd(aCampos, {"FILMVM","C",02,00})
           aAdd(aCampos, {"E5_CLIFOR","C",6,00})
           aAdd(aCampos, {"E5_LOJA","C",2,00})
   Private cArqTRB   := CriaTrab(aCampos)
   DbUseArea(.T.,"DBFCDX",cArqTRB,"XVIS")
   
   Private cPerg     := "GERANL"

   //TOTAL DE GASTOS DA ?REA COMERCIAL
   
   //DESPESA FIXA
   Private cGCOMMODF := "33101001,33101002,33101004,33101005,33101010,33101011,33101012,33101013,33101014,33101015,33101500"
   //DESPESA PERI?DICA
   Private cGCOMMODP := "33101003,33101007,33101008,33101016,33101030,33101034,33101035,33101046,33101501,33101504" 
   //RESCIS?O
   Private cGCOMRC   := "33101009,33101503"  
   //PREMIA??O
   Private cGCOMPR   := "33101006" 
   //MARKETING
   Private cGCOMMK   := "33101025,33101045"
   //DESPESAS DE VIAGENS
   Private cGCOMDV   := "33101019,33101057"
   //FRETE VENDAS
   Private cGCOMFV   := "33101505"
   //SERVI?OS TERCEIRIZADOS
   Private cGCOMST   := "33101017,33101023,33101026,33101027,33101047,33101051,33101052"
   //OUTRAS NATUREZAS
   Private cGCOMON   := "33101,33101018,33101020,33101021,33101022,33101028,33101029,33101031,33101032,33101033,33101036,33101037,33101038,"
           cGCOMON   += "33101039,33101040,33101041,33101042,33101048,33101049,33101050,33101053,33101054,33101055,33101058,33101502,33101506,"
           cGCOMON   += "33101507"
   //GASTOS DA META OR?AMENT?RIA DA ?REA COMERCIAL        
   Private cGCOMSTMO := "33101,33101002,33101004,33101005,33101012,33101013,33101014,33101015,33101016,33101018,33101019,33101020,33101021,"
           cGCOMSTMO += "33101022,33101028,33101029,33101030,33101032,33101033,33101034,33101035,33101036,33101037,33101038,33101039,33101040,"
           cGCOMSTMO += "33101041,33101042,33101046,33101047,33101048,33101049,33101050,33101051,33101052,33101053,33101054,33101055,33101057,"
           cGCOMSTMO += "33101500,33101501,33101504"

   //TOTAL DE GASTOS DA IND?STRIA
   //DESPESAS FINANCEIRAS DE IMPORTA??O
   //FINIMP
   Private cGINDFINI := "34103,34103001"
   //FLAT
   Private cGINDFLAT := "34103002"
   //LIBOR
   Private cGINDLIBO := "34103003"
   //BANQUEIRO
   Private cGINDBANQ := "34103004"
   //DISCREP?NCIA 
   Private cGINDDISC := "34103005"
   //IR - IMPORTA??O  
   Private cGINDIRIM := "34103006"
   //IOF - IMPORTA??O 
   Private cGINDIOF  := "34103007"
   //HIPOTECA - REGISTRO 
   Private cGINDHIPO := "34103008"
   //MONITORAMENTO - ESTOQUE 
   Private cGINDMONI := "34103009" 
   //IMPOSTOS
   //ICMS S/ IMOBILIZADO
   Private cGINDICMIM:= "11205001"
   //ICMS
   Private cGINDICMS := "11205,11205002"
   //IPI  
   Private cGINDIPI  := "11205003"
   //PIS 
   Private cGINDPIS  := "11205004"
   //COFINS  
   Private cGINDCOF  := "11205005"
   //CSLL  
   Private cGINDCSL  := "11205006"
   //IRPJ
   Private cGINDIRPJ := "11205007"
   //IRRF "11205008"
   Private cGINDIRRF := "11205008"
   //II 
   Private cGINDII   := "11205009"
   //MATERIA PRIMA
   //MATERIA PRIMA 
   Private cGINDMATMP:= "11250001"
   //AFRMM 
   Private cGINDAFRM := "32101001"
   //ARMAZENAGEM 
   Private cGINDARMZ := "32101002"
   //OPERA??O PORTU?RIA  
   Private cGINDOPORT:= "32101003"
   //DESPACHANTE ADUANEIRO  
   Private cGINDDADU := "32101004"
   //DESPESAS DE IMPORTA??O 
   Private cGINDDIMP := "32101,32101005"
   //FRETE IMPORTA??O 
   Private cGINDFRTI := "32101006"
                         
   //M?O DE OBRA DIRETA
   //DESPESA FIXA 
   Private cGINDMODF := "32102,32102001,32102002,32102004,32102005,32102010,32102011,32102012,32102013,32102014,32102015,32102500" 
   //DESPESA PERI?DICA  
   Private cGINDMOPD := "32102003,32102007,32102008,32102016,32102501"
   //RESCIS?O  
   Private cGINDRCS  := "32102009,32102503"
   //PREMIA??ES  
   Private cGINDPRM  := "32102006"
   //M?O DE OBRA INDIRETA
   //DESPESA FIXA 
   Private cGINDMIDF := "32103,32103001,32103002,32103004,32103005,32103010,32103011,32103012,32103013,32103014,32103015,32103500" 
   //DESPESA PERI?DICA  
   Private cGINDDPMI := "32103003,32103007,32103008,32103016,32103501,32104030,32104034,32104035,32104046,32104504" 
   //RESCIS?O  
   Private cGINDRMI  := "32103009,32103503"
   //PREMIA??ES  
   Private cGINDPMI  := "32103006" 
   //BENEFICIAMENTO  
   Private cGINDBNF  := "32104507,32104508"
   //INSUMOS 
   Private cGINDINSUM:= "11250005"
   //EMBALAGEM 
   Private cGINDEMBA := "11250007"
   //SEGURAN?A E MEDICINA DO TRABALHO 
   Private cGINDSMT  := "11250006"
   //MANUTEN??O
   //MANUTEN??O  ELETRICA  
   Private cGINDMELE := "32104511"  
   //MANUTEN??O FERRAMENTAL   
   Private cGINDMFERR:= "32104512"
   //MANUTEN??O MECANICA 
   Private cGINDMECA := "32104513"
   //ENERGIA 
   Private cGINDENER := "32104033"
   //AGUA  
   Private cGINDAGUA := "32104018"
   //ALUGUEL 
   Private cGINDALG  := "32104020"
   //SERVI?OS TERCEIRIZADOS  
   Private cGINDSTRZ := "32104017,32104023,32104026,32104027,32104047,32104051,32104052"
   //DESPESAS DE VIAGENS   
   Private cGINDDVG  := "32104057"
   //OUTRAS NATUREZAS
   Private cGINDONAT := "11250,11250002,11250003,11250004,11250009,11250010,11250011,321,32102502,32103502,32104,32104019,"
	       cGINDONAT += "32104021,32104022,32104025,32104028,32104029,32104031,32104032,32104036,32104037,32104038,32104039,32104040,32104041,"
           cGINDONAT += "32104042,32104045,32104048,32104049,32104050,32104053,32104054,32104055,32104058,32104509"
   //GASTOS DA META OR?AMENT?RIA DA INDUSTRIA        
   Private cGINDGMO  := "11250005,11250006,11250007,11250009,11250010,11250011,321,32102,32102002,32102004,32102005,32102012,32102013,32102014,"
           cGINDGMO  += "32102015,32102016,32102500,32102501,32103,32103002,32103004,32103005,32103012,32103013,32103014,32103015,32103016,32103500,"
           cGINDGMO  += "32103501,32104,32104018,32104019,32104020,32104021,32104022,32104028,32104029,32104030,32104032,32104033,32104034,32104035,"
           cGINDGMO  += "32104036,32104037,32104038,32104039,32104040,32104041,32104042,32104046,32104047,32104048,32104049,32104050,32104051,"
           cGINDGMO  += "32104052,32104053,32104054,32104055,32104057,32104504,32104509,32104511,32104512,32104513"
   
   //TOTAL DE GASTOS DA ?REA ADMINISTRATIVA
   //M?O DE OBRA
   //DESPESA FIXA
   Private cGADMDF   := "33102001,33102002,33102004,33102005,33102010,33102011,33102012,33102013,33102014,33102015,33102500"
   //DESPESA PERI?DICA
   Private cGADMDP   := "33102003,33102007,33102008,33102016,33102030,33102034,33102035,33102046,33102501,33102504"  
   //RESCIS?O
   Private cGADMRCS  := "33102009,33102503"   
   //PREMIA??O
   Private cGADMPRM  := "33102006"
   //CAIXA GERAL   
   Private cGADMCXG  := "11101001"
   //CAIXINHA   
   Private cGADMCXNH := "11101002"
   //INVESTIMENTOS
   Private cGADMINVST:= "11206001"
   //EMPRESTIMOS
   //EMPRESTIMO TOMADO  
   Private cGADMET   := "21101001,21101003,21101005,21101007"  
   //EMPRESTIMO CONCEDIDO    
   Private cGADMEC   := "21101002,21101004,21101006"
   //IMOBILIZADO
   //BENS IMOVEIS  
   Private cGADMIBI  := "13201,13201001,13201002,13201003,13201004"  
   //BENS MOVEIS     
   Private cGADMIBM  := "13202,13202001,13202002,13202003,13202004,13202005"
   //DESPESAS DE VIAGENS   
   Private cGADMDVG  := "33102057"
   //ADVOGADOS    
   Private cGADMADV  := "33102017,33102601"
   //SERVI?OS TERCEIRIZADOS 
   Private cGADMSTRZ := "33102023,33102026,33102027,33102047,33102051,33102052"
   //DESPESAS FINANCEIRAS   
   Private cXGADMDFIN := "34102,34102001,34102002,34102003,34102004,34102005,34102006" 
   //OUTRAS NATUREZAS   
   Private cGADMOUTNT:= "11101,11206,132,21101,33102,33102018,33102019,33102020,33102021,33102022,"
	       cGADMOUTNT+= "33102025,33102028,33102029,33102031,33102032,33102033,33102036,33102037,33102038,33102039,33102040,33102041,33102042,"
           cGADMOUTNT+= "33102045,33102048,33102049,33102050,33102053,33102054,33102055,33102058,33102502,33102505,33102506,33102600"
   //GASTOS DA META OR?AMENT?RIA DO ADMINISTRATIVO       
   Private cGADMMTOR := "11101,11101002,33102,33102002,33102004,33102005,33102012,33102013,33102014,33102015,"
	       cGADMMTOR += "33102016,33102018,33102019,33102020,33102021,33102022,33102028,33102029,33102030,33102032,33102033,33102034,33102035,"
           cGADMMTOR += "33102036,33102037,33102038,33102039,33102040,33102041,33102042,33102046,33102047,33102048,33102049,33102050,33102051,"
           cGADMMTOR += "33102052,33102053,33102054,33102055,33102057,33102500,33102501,33102504,33102505,33102506"

   //FRETES
   //FRETE VENDAS 
   Private cFRTVND   := "33101505"
   //FRETE IMPORTA??O 
   Private cFRTIMP   := "32101006" //"32104506"
   //FRETE EXPORTA??O 
   Private cFRTEXP   := "33101506"   
   //FRETE BENEFICIAMENTO 
   Private cFRTBENF  := "32104508"
   //FRETE TRANSFERENCIA
   Private cFRTTRANSF:= "32104509"
   //FRETE COM?RCIO  
   Private cFRTCOM   := "33101036"  
   //FRETE IND?STRIA  
   Private cFRTIND := "32104036"  
   //FRETE ADMINISTRATIVO  
   Private cFRTADM := "33102036"  
   //CUSTO DOS PRODUTOS VENDIDOS
   //CUSTO DIRETO
   //SALARIO DA PRODU??O  
   Private cCUSPVAP:= "32102,32102001,32102002,32102003,32102004,32102005,32102006,32102007,32102008,32102009,32102010,"
	       cCUSPVAP+= "32102011,32102012,32102013,32102014,32102015,32102016,32102500,32102501,32102503"
   //BENEFICIAMENTO   
   Private cCUSBENF:= "32104507,32104508"
   //INSUMOS 
   Private cCUSINSU:= "11250005"   
   //EMBALAGEM 
   Private cCUSEMBA:= "11250007"
   //CUSTO INDIRETO
   //SALARIO DO ADM INDUSTRIA  
   Private cCUSSI  := "32103,32103001,32103002,32103003,32103004,32103005,32103006,32103007,32103008,32103009,32103010,32103011,32103012,"
	       cCUSSI  += "32103013,32103014,32103015,32103016,32103500,32103501,32103503,32104030,32104034,32104035,32104046,32104504"
   //MANUTEN??O    
   Private cCUSMANU:= "32104511,32104512,32104513"
   //SERVI?OS TERCEIRIZADOS  
   Private cCUSSTRZ:= "32104017,32104023,32104026,32104027,32104047,32104051,32104052"   
   //SEGURAN?A E MEDICINA DO TRABALHO  
   Private cCUSSMT := "11250006"
   //ENERGIA    
   Private cCUSENRG:= "32104033"
   //AGUA    
   Private cCUSAGUA:= "32104018"  
   //ALUGUEL  
   Private cCUSALG := "32104020"
   //OUTRAS DESPESAS DA IND?STRIA  
   Private cCUSODIN:= "11250,11250002,11250003,11250004,11250009,11250010,11250011,321,32102502,32103502,32104,32104019,"
	       cCUSODIN+= "32104021,32104022,32104025,32104028,32104029,32104031,32104032,32104036,32104037,32104038,32104039,32104040,32104041,"
	       cCUSODIN+= "32104042,32104045,32104048,32104049,32104050,32104053,32104054,32104055,32104056,32104057,32104058,32104509"
   //(-)DESPESAS/RECEITAS OPERACIONAIS
   //COMERCIAIS
   Private cDRCOM := "33101,33101001,33101002,33101003,33101004,33101005,33101006,33101007,33101008,33101009,33101010,33101011,33101012,"
	       cDRCOM += "33101013,33101014,33101015,33101016,33101017,33101018,33101019,33101020,33101021,33101022,33101023,33101025,33101026,"
	       cDRCOM += "33101027,33101028,33101029,33101030,33101031,33101032,33101033,33101034,33101035,33101036,33101037,33101038,33101039,"
	       cDRCOM += "33101040,33101041,33101042,33101045,33101046,33101047,33101048,33101049,33101050,33101051,33101052,33101053,33101054,"
	       cDRCOM += "33101055,33101056,33101057,33101058,33101500,33101501,33101502,33101503,33101504,33101505,33101506,33101507"
   //ADMINISTRATIVAS
   Private cDRADM := "11101,11101001,11101002,33102,33102001,33102002,33102003,33102004,"
	       cDRADM += "33102005,33102006,33102007,33102008,33102009,33102010,33102011,33102012,33102013,33102014,33102015,33102016,33102017,"
	       cDRADM += "33102018,33102019,33102020,33102021,33102022,33102023,33102025,33102026,33102027,33102028,33102029,33102030,33102031,"
	       cDRADM += "33102032,33102033,33102034,33102035,33102036,33102037,33102038,33102039,33102040,33102041,33102042,33102045,33102046,"
	       cDRADM += "33102047,33102048,33102049,33102050,33102051,33102052,33102053,33102054,33102055,33102056,33102057,33102058,33102500,"
	       cDRADM += "33102501,33102502,33102503,33102504,33102505,33102506,33102600"
   //FINANCEIRAS
   //DESPESAS FINANCEIRAS
   Private cDRDFIN := "34102,34102001,34102002,34102003,34102004,34102005,34102006"
   //RECEITAS FINANCEIRAS
   Private cDRRFIN := "34101,34101001,34101002,34101003,34101004"


   MontPerg()
   
   If Pergunte(cPerg,.T.)
      
      Private dDataIni := MV_PAR01
      Private dDataFim := MV_PAR02
      Private nVisoes  := MV_PAR03 //1=Todas;2=Comercial;3=Industrial;4=Administrativo;5=Fret/Custo-Desp/Rec
      Private DeFilial := MV_PAR04
      Private AteFilial:= MV_PAR05
      Private nHistoric:= MV_PAR07
      Private cNatIni  := MV_PAR08
      Private cNatFim  := MV_PAR09
      Private cExclFil := Alltrim(MV_PAR06)
      
      Processa({||RGeraGA()},"Preparando Gerencial Analitico, Aguarde ...")
      
   EndIf
   
Return

Static Function RGeraGA()
   
   //1=Todas;2=Comercial;3=Industrial;4=Administrativo;5=Despesas/Receitas 
      
   If nVisoes == 1 .Or. nVisoes == 2
      Processa({||RNatCom(2,10)},"Processando Vis?o Comercial, Aguarde ...")
   EndIf
   If nVisoes == 1 .Or. nVisoes == 3
      Processa({||RNatCom(3,47)},"Processando Vis?o Industrial, Aguarde ...")
   EndIf
   If nVisoes == 1 .Or. nVisoes == 4
      Processa({||RNatCom(4,17)},"Processando Vis?o Administrativa, Aguarde ...")
   EndIf
   If nVisoes == 1 .Or. nVisoes == 5
      Processa({||RNatCom(5,24)},"Processando Vis?o Frete/Custos Desp/Receitas, Aguarde ...")
   EndIf
   Processa({|lEnd|RImpRGA()},"Imprimindo Relatorio, Aguarde ...")
   
Return

Static Function RNatCom(nOpVis,nProc)

   aMsg := {}
   
   If nOpVis == 2 //Vis?o Comercial
   
      aAdd(aMsg,"Despesa Fixa")
      aAdd(aMsg,"Despesa Periodica")
      aAdd(aMsg,"Recis?o")
      aAdd(aMsg,"Premia??o")
      aAdd(aMsg,"Marketing")
      aAdd(aMsg,"Despesas Viagens")
      aAdd(aMsg,"Frete Vendas")
      aAdd(aMsg,"Servi?os Terceirizados")
      aAdd(aMsg,"Outras Naturezas")
      aAdd(aMsg,"Gastos Metas")
   
   ElseIf nOpVis == 3 //Vis?o Ind?strial
   
      aAdd(aMsg,"FINIMP")
      aAdd(aMsg,"FLAT")
      aAdd(aMsg,"LIBOR")
      aAdd(aMsg,"BANQUEIRO")
      aAdd(aMsg,"DISCREP?NCIA")
      aAdd(aMsg,"IR-IMPORTA??O")
      aAdd(aMsg,"IOF-IMPORTA??O")
      aAdd(aMsg,"HIPOTECA-REGISTRO") 
      aAdd(aMsg,"MONIT-ESTOQUE")
      aAdd(aMsg,"ICMS S/IMOBILIZADO")
      aAdd(aMsg,"ICMS")
      aAdd(aMsg,"IPI")
      aAdd(aMsg,"PIS")
      aAdd(aMsg,"COFINS")
      aAdd(aMsg,"CSLL")
      aAdd(aMsg,"IRPJ")
      aAdd(aMsg,"IRRF 11205008")
      aAdd(aMsg,"II")
      aAdd(aMsg,"MATERIA PRIMA")
      aAdd(aMsg,"AFRMM")
      aAdd(aMsg,"ARMAZENAGEM")
      aAdd(aMsg,"OP.PORTU?RIA")
      aAdd(aMsg,"DESP ADUANEIRO")
      aAdd(aMsg,"DESP IMPORTA??O")
      aAdd(aMsg,"FRETE IMPORTA??O")
      aAdd(aMsg,"DESPESA FIXA")
      aAdd(aMsg,"DESPESA PERI?DICA")
      aAdd(aMsg,"RESCIS?O")
      aAdd(aMsg,"PREMIA??ES")  
      aAdd(aMsg,"MOD-DESPESA FIXA")
      aAdd(aMsg,"MOD-DESPESA PERI?DICA")
      aAdd(aMsg,"MOD-RESCIS?O")
      aAdd(aMsg,"MOD-PREMIA??ES")  
      aAdd(aMsg,"BENEFICIAMENTO")
      aAdd(aMsg,"INSUMOS")
      aAdd(aMsg,"EMBALAGEM") 
      aAdd(aMsg,"SEG MED TRABALHO")
      aAdd(aMsg,"MANUTEN??O  ELETRICA")
      aAdd(aMsg,"MANUTEN??O FERRAMENTAL")   
      aAdd(aMsg,"MANUTEN??O MECANICA")
      aAdd(aMsg,"ENERGIA")
      aAdd(aMsg,"AGUA")
      aAdd(aMsg,"ALUGUEL")
      aAdd(aMsg,"SERV TERCEIRIZADOS")
      aAdd(aMsg,"DESPESAS DE VIAGENS")  
      aAdd(aMsg,"OUTRAS NATUREZAS")
      aAdd(aMsg,"META ORC INDUSTRIA")

   ElseIf nOpVis == 4 //Vis?o Administrativa
   

      aAdd(aMsg,"DESPESA FIXA")
      aAdd(aMsg,"DESPESA PERI?DICA")
      aAdd(aMsg,"RESCIS?O")
      aAdd(aMsg,"PREMIA??O")
      aAdd(aMsg,"CAIXA GERAL")   
      aAdd(aMsg,"CAIXINHA")
      aAdd(aMsg,"INVESTIMENTOS")
      aAdd(aMsg,"EMPRESTIMO TOMADO")
      aAdd(aMsg,"EMPRESTIMO CONCEDIDO")
      aAdd(aMsg,"BENS IMOVEIS")
      aAdd(aMsg,"BENS MOVEIS")
      aAdd(aMsg,"DESPESAS DE VIAGENS")
      aAdd(aMsg,"ADVOGADOS")
      aAdd(aMsg,"SERVI?OS TERCEIRIZADOS")
      aAdd(aMsg,"DESPESAS FINANCEIRAS")
      aAdd(aMsg,"OUTRAS NATUREZAS")
      aAdd(aMsg,"META ORC ADMIN")

   ElseIf nOpVis == 5 //Frete/Custos Produtos Vendidos - Despesas/Receitas Operacionais
   
      aAdd(aMsg,"FRETE VENDAS") 
      aAdd(aMsg,"FRETE IMPORTA??O") 
      aAdd(aMsg,"FRETE EXPORTA??O") 
      aAdd(aMsg,"FRETE BENEFICIAMENTO") 
      aAdd(aMsg,"FRETE TRANSFERENCIA") 
      aAdd(aMsg,"FRETE COM?RCIO") 
      aAdd(aMsg,"FRETE IND?STRIA") 
      aAdd(aMsg,"FRETE ADMINISTRATIVO") 
      aAdd(aMsg,"SALARIO DA PRODU??O") 
      aAdd(aMsg,"BENEFICIAMENTO") 
      aAdd(aMsg,"INSUMOS") 
      aAdd(aMsg,"EMBALAGEM") 
      aAdd(aMsg,"SALARIO DO ADM INDUSTRIA") 
      aAdd(aMsg,"MANUTEN??O") 
      aAdd(aMsg,"SERVI?OS TERCEIRIZADOS") 
      aAdd(aMsg,"SEGURAN?A E MEDICINA DO TRABALHO") 
      aAdd(aMsg,"ENERGIA") 
      aAdd(aMsg,"AGUA") 
      aAdd(aMsg,"ALUGUEL") 
      aAdd(aMsg,"OUTRAS DESP IND?STRIA") 
      aAdd(aMsg,"COMERCIAIS") 
      aAdd(aMsg,"ADMINISTRATIVA") 
      aAdd(aMsg,"DESPESAS FINANCEIRAS") 
      aAdd(aMsg,"RECEITAS FINANCEIRAS") 
   
   EndIf
   
   ProcRegua(nProc)
   
   For nCom := 1 To nProc
      
      IncProc(aMsg[nCom]+" "+Alltrim(Str(nCom))+" / "+Alltrim(Str(nProc)))
   
      cQryCOM := " SELECT CASE WHEN SE5.E5_FILIAL = '02' THEN 'COSP'"
      cQryCOM += "             WHEN SE5.E5_FILIAL = '03' THEN 'CORS'"
      cQryCOM += "             WHEN SE5.E5_FILIAL = '04' THEN 'COPB'"
      cQryCOM += "             WHEN SE5.E5_FILIAL = '06' THEN 'COPE' END AS E5_FILIAL,"
      
      If nOpVis == 2 //Vis?o Comercial

         If nCom == 1 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA COMERCIAL' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA' xSUBINDI,"
            cQryCOM += "        'DESPESA FIXA' xINTRAIN,"
         ElseIf nCom == 2
            cQryCOM += "        'TOTAL DE GASTOS DA AREA COMERCIAL' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA' xSUBINDI,"
            cQryCOM += "        'DESPESA PERIODICA' xINTRAIN,"
         ElseIf nCom == 3
            cQryCOM += "        'TOTAL DE GASTOS DA AREA COMERCIAL' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA' xSUBINDI,"
            cQryCOM += "        'RESCISAO' xINTRAIN,"            
         ElseIf nCom == 4
            cQryCOM += "        'TOTAL DE GASTOS DA AREA COMERCIAL' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA' xSUBINDI,"
            cQryCOM += "        'PREMIACAO' xINTRAIN,"
         ElseIf nCom == 5
            cQryCOM += "        'TOTAL DE GASTOS DA AREA COMERCIAL' xINDICAD,"
            cQryCOM += "        'MARKETING' xSUBINDI,"
            cQryCOM += "        '                 ' xINTRAIN,"
         ElseIf nCom == 6
            cQryCOM += "        'TOTAL DE GASTOS DA AREA COMERCIAL' xINDICAD,"
            cQryCOM += "        'DESPESA DE VIAGEM' xSUBINDI,"
            cQryCOM += "        '                 ' xINTRAIN,"
         ElseIf nCom == 7
            cQryCOM += "        'TOTAL DE GASTOS DA AREA COMERCIAL' xINDICAD,"
            cQryCOM += "        'FRETE DE VENDAS' xSUBINDI,"
            cQryCOM += "        '                 ' xINTRAIN,"
         ElseIf nCom == 8
            cQryCOM += "        'TOTAL DE GASTOS DA AREA COMERCIAL' xINDICAD,"
            cQryCOM += "        'SERVI?OS TERCEIRIZADOS' xSUBINDI,"
            cQryCOM += "        '                 ' xINTRAIN,"
         ElseIf nCom == 9
            cQryCOM += "        'TOTAL DE GASTOS DA AREA COMERCIAL' xINDICAD,"
            cQryCOM += "        'OUTRAS NATUREZAS' xSUBINDI,"
            cQryCOM += "        '                 ' xINTRAIN,"
         ElseIf nCom == 10
            cQryCOM += "        'GASTOS DA META ORCAMENTARIA DA AREA COMERCIAL' xINDICAD,"
            //cQryCOM += "        'OUTRAS NATUREZAS' xSUBINDI,"
            cQryCOM += "        '                ' xSUBINDI,"
            cQryCOM += "        '                 ' xINTRAIN,"            
         EndIf
      
      ElseIf nOpVis == 3 //Vis?o Ind?strial

         If nCom == 1 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'DESPESAS FINANCEIRAS DE IMPORTACAO' xSUBINDI,"
            cQryCOM += "        'FINIMP' xINTRAIN,"            
         ElseIf nCom == 2 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'DESPESAS FINANCEIRAS DE IMPORTACAO' xSUBINDI,"
            cQryCOM += "        'FLAT' xINTRAIN,"            
         ElseIf nCom == 3 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'DESPESAS FINANCEIRAS DE IMPORTACAO' xSUBINDI,"
            cQryCOM += "        'LIBOR' xINTRAIN,"            
         ElseIf nCom == 4 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'DESPESAS FINANCEIRAS DE IMPORTACAO' xSUBINDI,"
            cQryCOM += "        'BANQUEIRO' xINTRAIN,"            
         ElseIf nCom == 5 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'DESPESAS FINANCEIRAS DE IMPORTACAO' xSUBINDI,"
            cQryCOM += "        'DISCREPANCIA' xINTRAIN,"            
         ElseIf nCom == 6 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'DESPESAS FINANCEIRAS DE IMPORTACAO' xSUBINDI,"
            cQryCOM += "        'IR - IMPORTACAO' xINTRAIN,"            
         ElseIf nCom == 7 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'DESPESAS FINANCEIRAS DE IMPORTACAO' xSUBINDI,"
            cQryCOM += "        'IOF - IMPORTACAO' xINTRAIN,"            
         ElseIf nCom == 8 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'DESPESAS FINANCEIRAS DE IMPORTACAO' xSUBINDI,"
            cQryCOM += "        'HIPOTECA - REGISTRO' xINTRAIN,"            
         ElseIf nCom == 9 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'DESPESAS FINANCEIRAS DE IMPORTACAO' xSUBINDI,"
            cQryCOM += "        'MONITORAMENTO - ESTOQUE' xINTRAIN,"            
         ElseIf nCom == 10 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'IMPOSTOS' xSUBINDI,"
            cQryCOM += "        'ICMS S/ IMOBILIZADO' xINTRAIN,"            
         ElseIf nCom == 11 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'IMPOSTOS' xSUBINDI,"
            cQryCOM += "        'ICMS' xINTRAIN,"            
         ElseIf nCom == 12 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'IMPOSTOS' xSUBINDI,"
            cQryCOM += "        'IPI' xINTRAIN,"            
         ElseIf nCom == 13 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'IMPOSTOS' xSUBINDI,"
            cQryCOM += "        'PIS' xINTRAIN,"            
         ElseIf nCom == 14 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'IMPOSTOS' xSUBINDI,"
            cQryCOM += "        'COFINS' xINTRAIN,"            
         ElseIf nCom == 15 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'IMPOSTOS' xSUBINDI,"
            cQryCOM += "        'CSLL' xINTRAIN,"            
         ElseIf nCom == 16 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'IMPOSTOS' xSUBINDI,"
            cQryCOM += "        'IRPJ' xINTRAIN,"            
         ElseIf nCom == 17 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'IMPOSTOS' xSUBINDI,"
            cQryCOM += "        'IRRF 11205008' xINTRAIN,"            
         ElseIf nCom == 18 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'IMPOSTOS' xSUBINDI,"
            cQryCOM += "        'II' xINTRAIN,"            
         ElseIf nCom == 19 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MATERIA PRIMA(TOTAL)' xSUBINDI,"
            cQryCOM += "        'MATERIA PRIMA' xINTRAIN,"            
         ElseIf nCom == 20 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MATERIA PRIMA(TOTAL)' xSUBINDI,"
            cQryCOM += "        'AFRMM' xINTRAIN,"            
         ElseIf nCom == 21 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MATERIA PRIMA(TOTAL)' xSUBINDI,"
            cQryCOM += "        'ARMAZENAGEM' xINTRAIN,"            
         ElseIf nCom == 22
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MATERIA PRIMA(TOTAL)' xSUBINDI,"
            cQryCOM += "        'OPERACAO PORTUARIA' xINTRAIN,"            
         ElseIf nCom == 23
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MATERIA PRIMA(TOTAL)' xSUBINDI,"
            cQryCOM += "        'DESPACHANTE ADUANEIRO' xINTRAIN,"            
         ElseIf nCom == 24
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MATERIA PRIMA(TOTAL)' xSUBINDI,"
            cQryCOM += "        'DESPESAS DE IMPORTACAO' xINTRAIN,"            
         ElseIf nCom == 25
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MATERIA PRIMA(TOTAL)' xSUBINDI,"
            cQryCOM += "        'FRETE IMPORTACAO' xINTRAIN,"            
         ElseIf nCom == 26
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA DIRETA' xSUBINDI,"
            cQryCOM += "        'DESPESA FIXA' xINTRAIN,"            
         ElseIf nCom == 27 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA DIRETA' xSUBINDI,"
            cQryCOM += "        'DESPESA PERIODICA' xINTRAIN,"            
         ElseIf nCom == 28
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA DIRETA' xSUBINDI,"
            cQryCOM += "        'RECISOES' xINTRAIN,"            
         ElseIf nCom == 29 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA DIRETA' xSUBINDI,"
            cQryCOM += "        'PREMIACOES' xINTRAIN,"            
         ElseIf nCom == 30 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA INDIRETA' xSUBINDI,"
            cQryCOM += "        'DESPESA FIXA' xINTRAIN,"            
         ElseIf nCom == 31 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA INDIRETA' xSUBINDI,"
            cQryCOM += "        'DESPESA PERIODICA' xINTRAIN,"            
         ElseIf nCom == 32 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA INDIRETA' xSUBINDI,"
            cQryCOM += "        'RECISOES' xINTRAIN,"            
         ElseIf nCom == 33 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA INDIRETA' xSUBINDI,"
            cQryCOM += "        'PREMIACOES' xINTRAIN,"            
         ElseIf nCom == 34 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'BENEFICIAMENTO' xSUBINDI,"
            cQryCOM += "        '          ' xINTRAIN,"            
         ElseIf nCom == 35 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'INSUMOS' xSUBINDI,"
            cQryCOM += "        '          ' xINTRAIN,"            
         ElseIf nCom == 36 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'EMBALAGEM' xSUBINDI,"
            cQryCOM += "        '          ' xINTRAIN,"            
         ElseIf nCom == 37 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'SEGURANCA E MEDICINA DO TRABALHO' xSUBINDI,"
            cQryCOM += "        '          ' xINTRAIN,"            
         ElseIf nCom == 38 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MANUTENCAO' xSUBINDI,"
            cQryCOM += "        'MANUTENCAO  ELETRICA' xINTRAIN,"            
         ElseIf nCom == 39 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MANUTENCAO' xSUBINDI,"
            cQryCOM += "        'MANUTENCAO FERRAMENTAL' xINTRAIN,"            
         ElseIf nCom == 40 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'MANUTENCAO' xSUBINDI,"
            cQryCOM += "        'MANUTENCAO MECANICA' xINTRAIN,"            
         ElseIf nCom == 41 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'ENERGIA' xSUBINDI,"
            cQryCOM += "        '                   ' xINTRAIN,"            
         ElseIf nCom == 42 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'AGUA' xSUBINDI,"
            cQryCOM += "        '                   ' xINTRAIN,"            
         ElseIf nCom == 43 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'ALUGUEL' xSUBINDI,"
            cQryCOM += "        '                   ' xINTRAIN,"            
         ElseIf nCom == 44 
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'SERVICOS TERCEIRIZADOS' xSUBINDI,"
            cQryCOM += "        '                   ' xINTRAIN,"            
         ElseIf nCom == 45
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'DESPESAS DE VIAGENS' xSUBINDI,"
            cQryCOM += "        '                   ' xINTRAIN,"            
         ElseIf nCom == 46
            cQryCOM += "        'TOTAL DE GASTOS DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        'OUTRAS NATUREZAS' xSUBINDI,"
            cQryCOM += "        '                   ' xINTRAIN,"            
         ElseIf nCom == 47 
            cQryCOM += "        'GASTOS DA META OR?AMENT?RIA DA INDUSTRIA' xINDICAD,"
            cQryCOM += "        '                ' xSUBINDI,"
            cQryCOM += "        '                   ' xINTRAIN,"            
         EndIf
                
      ElseIf nOpVis == 4 //Vis?o Administrativa

         If nCom == 1 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA' xSUBINDI,"
            cQryCOM += "        'DESPESA FIXA' xINTRAIN,"            
         ElseIf nCom == 2 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA' xSUBINDI,"
            cQryCOM += "        'DESPESA PERIODICA' xINTRAIN,"            
         ElseIf nCom == 3 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA' xSUBINDI,"
            cQryCOM += "        'RESCISOES' xINTRAIN,"            
         ElseIf nCom == 4 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'MAO DE OBRA' xSUBINDI,"
            cQryCOM += "        'PREMIACOES' xINTRAIN,"            
         ElseIf nCom == 5 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'CAIXA GERAL' xSUBINDI,"
            cQryCOM += "        '          ' xINTRAIN,"            
         ElseIf nCom == 6 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'CAIXINHA' xSUBINDI,"
            cQryCOM += "        '          ' xINTRAIN,"            
         ElseIf nCom == 7 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'INVESTIMENTOS' xSUBINDI,"
            cQryCOM += "        '          ' xINTRAIN,"            
         ElseIf nCom == 8 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'EMPRESTIMOS' xSUBINDI,"
            cQryCOM += "        'EMPRESTIMO TOMADO' xINTRAIN,"            
         ElseIf nCom == 9 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'EMPRESTIMOS' xSUBINDI,"
            cQryCOM += "        'EMPRESTIMO CONCEDIDO' xINTRAIN,"            
         ElseIf nCom == 10 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'IMOBILIZADO' xSUBINDI,"
            cQryCOM += "        'BENS IMOVEIS' xINTRAIN,"            
         ElseIf nCom == 11 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'IMOBILIZADO' xSUBINDI,"
            cQryCOM += "        'BENS MOVEIS' xINTRAIN,"            
         ElseIf nCom == 12 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'DESPESAS DE VIAGENS' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         ElseIf nCom == 13 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'ADVOGADOS' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         ElseIf nCom == 14 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'SERVI?OS TERCEIRIZADOS' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         ElseIf nCom == 15 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'DESPESAS FINANCEIRAS' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         ElseIf nCom == 16 
            cQryCOM += "        'TOTAL DE GASTOS DA AREA ADMINISTRATIVA' xINDICAD,"
            cQryCOM += "        'OUTRAS NATUREZAS' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         ElseIf nCom == 17 
            cQryCOM += "        'GASTOS DA META OR?AMENTARIA DO ADMINISTRATIVO' xINDICAD,"
            cQryCOM += "        '                ' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         EndIf
         
      ElseIf nOpVis == 5 //Vis?o Frete/Custos Prod.Vendidos / Despesas/Receitas Operacionais.
         
         If nCom == 1 
            cQryCOM += "        'FRETES' xINDICAD,"
            cQryCOM += "        'FRETE VENDAS' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         ElseIf nCom == 2 
            cQryCOM += "        'FRETES' xINDICAD,"
            cQryCOM += "        'FRETE IMPORTACAO' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         ElseIf nCom == 3 
            cQryCOM += "        'FRETES' xINDICAD,"
            cQryCOM += "        'FRETE EXPORTA??O' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         ElseIf nCom == 4 
            cQryCOM += "        'FRETES' xINDICAD,"
            cQryCOM += "        'FRETE BENEFICIAMENTO' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         ElseIf nCom == 5 
            cQryCOM += "        'FRETES' xINDICAD,"
            cQryCOM += "        'FRETE TRANSFERENCIA' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         ElseIf nCom == 6 
            cQryCOM += "        'FRETES' xINDICAD,"
            cQryCOM += "        'FRETE COMERCIO' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         ElseIf nCom == 7 
            cQryCOM += "        'FRETES' xINDICAD,"
            cQryCOM += "        'FRETE INDUSTRIA' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         ElseIf nCom == 8 
            cQryCOM += "        'FRETES' xINDICAD,"
            cQryCOM += "        'FRETE ADMINISTRATIVO' xSUBINDI,"
            cQryCOM += "        '           ' xINTRAIN,"            
         ElseIf nCom == 9 
            cQryCOM += "        'CUSTO DOS PRODUTOS VENDIDOS' xINDICAD,"
            cQryCOM += "        'CUSTO DIRETO' xSUBINDI,"
            cQryCOM += "        'SALARIO DA PRODUCAO' xINTRAIN,"            
         ElseIf nCom == 10 
            cQryCOM += "        'CUSTO DOS PRODUTOS VENDIDOS' xINDICAD,"
            cQryCOM += "        'CUSTO DIRETO' xSUBINDI,"
            cQryCOM += "        'BENEFICIAMENTO' xINTRAIN,"            
         ElseIf nCom == 11
            cQryCOM += "        'CUSTO DOS PRODUTOS VENDIDOS' xINDICAD,"
            cQryCOM += "        'CUSTO DIRETO' xSUBINDI,"
            cQryCOM += "        'INSUMOS' xINTRAIN,"            
         ElseIf nCom == 12
            cQryCOM += "        'CUSTO DOS PRODUTOS VENDIDOS' xINDICAD,"
            cQryCOM += "        'CUSTO DIRETO' xSUBINDI,"
            cQryCOM += "        'EMBALAGEM' xINTRAIN,"            
         ElseIf nCom == 13 
            cQryCOM += "        'CUSTO DOS PRODUTOS VENDIDOS' xINDICAD,"
            cQryCOM += "        'CUSTO INDIRETO' xSUBINDI,"
            cQryCOM += "        'SALARIO DO ADM INDUSTRIA' xINTRAIN,"            
         ElseIf nCom == 14 
            cQryCOM += "        'CUSTO DOS PRODUTOS VENDIDOS' xINDICAD,"
            cQryCOM += "        'CUSTO INDIRETO' xSUBINDI,"
            cQryCOM += "        'MANUTENCAO' xINTRAIN,"            
         ElseIf nCom == 15 
            cQryCOM += "        'CUSTO DOS PRODUTOS VENDIDOS' xINDICAD,"
            cQryCOM += "        'CUSTO INDIRETO' xSUBINDI,"
            cQryCOM += "        'SERVICOS TERCEIRIZADOS' xINTRAIN,"            
         ElseIf nCom == 16 
            cQryCOM += "        'CUSTO DOS PRODUTOS VENDIDOS' xINDICAD,"
            cQryCOM += "        'CUSTO INDIRETO' xSUBINDI,"
            cQryCOM += "        'SEGURANCA E MEDICINA DO TRABALHO' xINTRAIN,"            
         ElseIf nCom == 17 
            cQryCOM += "        'CUSTO DOS PRODUTOS VENDIDOS' xINDICAD,"
            cQryCOM += "        'CUSTO INDIRETO' xSUBINDI,"
            cQryCOM += "        'ENERGIA' xINTRAIN,"            
         ElseIf nCom == 18 
            cQryCOM += "        'CUSTO DOS PRODUTOS VENDIDOS' xINDICAD,"
            cQryCOM += "        'CUSTO INDIRETO' xSUBINDI,"
            cQryCOM += "        'AGUA' xINTRAIN,"            
         ElseIf nCom == 19 
            cQryCOM += "        'CUSTO DOS PRODUTOS VENDIDOS' xINDICAD,"
            cQryCOM += "        'CUSTO INDIRETO' xSUBINDI,"
            cQryCOM += "        'ALUGUEL' xINTRAIN,"            
         ElseIf nCom == 20 
            cQryCOM += "        'CUSTO DOS PRODUTOS VENDIDOS' xINDICAD,"
            cQryCOM += "        'CUSTO INDIRETO' xSUBINDI,"
            cQryCOM += "        'OUTRAS DESPESAS DA IND?STRIA' xINTRAIN,"            
         ElseIf nCom == 21 
            cQryCOM += "        '(-)DESPESAS/RECEITAS OPERACIONAIS' xINDICAD,"
            cQryCOM += "        'COMERCIAIS' xSUBINDI,"
            cQryCOM += "        '                   ' xINTRAIN,"            
         ElseIf nCom == 22 
            cQryCOM += "        '(-)DESPESAS/RECEITAS OPERACIONAIS' xINDICAD,"
            cQryCOM += "        'ADMINISTRATIVAS' xSUBINDI,"
            cQryCOM += "        '                   ' xINTRAIN,"            
         ElseIf nCom == 23 
            cQryCOM += "        '(-)DESPESAS/RECEITAS OPERACIONAIS' xINDICAD,"
            cQryCOM += "        'FINANCEIRAS' xSUBINDI,"
            cQryCOM += "        'DESPESAS FINANCEIRAS' xINTRAIN,"            
         ElseIf nCom == 24 
            cQryCOM += "        '(-)DESPESAS/RECEITAS OPERACIONAIS' xINDICAD,"
            cQryCOM += "        'FINANCEIRAS' xSUBINDI,"
            cQryCOM += "        'RECEITAS FINANCEIRAS' xINTRAIN,"            
         EndIf
         
      EndIf
      
      cQryCOM += "        CASE WHEN SUBSTRING(SE5.E5_DATA,5,2) = '01' THEN SUBSTRING(SE5.E5_DATA,5,2)+'-JANEIRO'"
      cQryCOM += "             WHEN SUBSTRING(SE5.E5_DATA,5,2) = '02' THEN SUBSTRING(SE5.E5_DATA,5,2)+'-FEVEREIRO'"
      cQryCOM += "             WHEN SUBSTRING(SE5.E5_DATA,5,2) = '03' THEN SUBSTRING(SE5.E5_DATA,5,2)+'-MAR?O'"
      cQryCOM += "             WHEN SUBSTRING(SE5.E5_DATA,5,2) = '04' THEN SUBSTRING(SE5.E5_DATA,5,2)+'-ABRIL'"
      cQryCOM += "             WHEN SUBSTRING(SE5.E5_DATA,5,2) = '05' THEN SUBSTRING(SE5.E5_DATA,5,2)+'-MAIO'"
      cQryCOM += "             WHEN SUBSTRING(SE5.E5_DATA,5,2) = '06' THEN SUBSTRING(SE5.E5_DATA,5,2)+'-JUNHO'"
      cQryCOM += "             WHEN SUBSTRING(SE5.E5_DATA,5,2) = '07' THEN SUBSTRING(SE5.E5_DATA,5,2)+'-JULHO'"     
      cQryCOM += "             WHEN SUBSTRING(SE5.E5_DATA,5,2) = '08' THEN SUBSTRING(SE5.E5_DATA,5,2)+'-AGOSTO'"     
      cQryCOM += "             WHEN SUBSTRING(SE5.E5_DATA,5,2) = '09' THEN SUBSTRING(SE5.E5_DATA,5,2)+'-SETEMBRO'"     
      cQryCOM += "             WHEN SUBSTRING(SE5.E5_DATA,5,2) = '10' THEN SUBSTRING(SE5.E5_DATA,5,2)+'-OUTUBRO'"     
      cQryCOM += "             WHEN SUBSTRING(SE5.E5_DATA,5,2) = '11' THEN SUBSTRING(SE5.E5_DATA,5,2)+'-NOVEMBRO'"     
      cQryCOM += "             WHEN SUBSTRING(SE5.E5_DATA,5,2) = '12' THEN SUBSTRING(SE5.E5_DATA,5,2)+'-DEZEMBRO' END MES,"
      cQryCOM += "  SE5.E5_NATUREZ,SED.ED_DESCRIC,SE5.E5_PREFIXO,SE5.E5_NUMERO,SE5.E5_PARCELA,SE5.E5_TIPO,"
      cQryCOM += "  SE5.E5_BENEF,SE5.E5_HISTOR,SE5.E5_DATA,SE5.E5_VALOR,SE5.E5_FILIAL xFILMOV,"
      cQryCOM += "  SE5.E5_CLIFOR,SE5.E5_LOJA" //Five Solutions - 24/01/2011
      cQryCOM += "  FROM "+RetSQLName("SE5")+" SE5,"+RetSQLName("SED")+" SED"  
      cQryCOM += " WHERE SE5.E5_RECPAG = 'P'"    
      cQryCOM += "   AND SE5.E5_CLIFOR <> '006629'"    
      cQryCOM += "   AND  SUBSTRING(SE5.E5_DATA,1,6) BETWEEN '"+Substr(DTOS(dDataIni),1,6)+"' AND '"+Substr(DTOS(dDataFim),1,6)+"'"
      
      If nOpVis == 2 //Vis?o Comercial
      
         If nCom == 1 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGCOMMODF,",")
         ElseIf nCom == 2
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGCOMMODP,",")
         ElseIf nCom == 3
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGCOMRC,",")
         ElseIf nCom == 4
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGCOMPR,",")
         ElseIf nCom == 5
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGCOMMK,",")
         ElseIf nCom == 6
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGCOMDV,",")
         ElseIf nCom == 7
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGCOMFV,",")
         ElseIf nCom == 8
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGCOMST,",")
         ElseIf nCom == 9
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGCOMON,",")
         ElseIf nCom == 10
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGCOMSTMO,",")
         EndIf
      
      ElseIf nOpVis == 3 //Vis?o Ind?strial
      
         If nCom == 1 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDFINI,",")
         ElseIf nCom == 2 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDFLAT,",")
         ElseIf nCom == 3 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDLIBO,",")
         ElseIf nCom == 4 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDBANQ,",")
         ElseIf nCom == 5 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDDISC,",")
         ElseIf nCom == 6 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDIRIM,",")
         ElseIf nCom == 7 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDIOF,",")
         ElseIf nCom == 8 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDHIPO,",")
         ElseIf nCom == 9 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDMONI,",")
         ElseIf nCom == 10 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDICMIM,",")
         ElseIf nCom == 11 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDICMS,",")
         ElseIf nCom == 12 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDIPI,",")
         ElseIf nCom == 13 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDPIS,",")
         ElseIf nCom == 14 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDCOF,",")
         ElseIf nCom == 15 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDCSL,",")
         ElseIf nCom == 16 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDIRPJ,",")
         ElseIf nCom == 17 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDIRRF,",")
         ElseIf nCom == 18 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDII,",")
         ElseIf nCom == 19 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDMATMP,",")
         ElseIf nCom == 20 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDAFRM,",")
         ElseIf nCom == 21 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDARMZ,",")
         ElseIf nCom == 22
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDOPORT,",")
         ElseIf nCom == 23
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDDADU,",")
         ElseIf nCom == 24
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDDIMP,",")
         ElseIf nCom == 25
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDFRTI,",")
         ElseIf nCom == 26
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDMODF,",")
         ElseIf nCom == 27 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDMOPD,",")
         ElseIf nCom == 28
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDRCS,",")
         ElseIf nCom == 29 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDPRM,",")
         ElseIf nCom == 30 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDMIDF,",")
         ElseIf nCom == 31 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDDPMI,",")
         ElseIf nCom == 32 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDRMI,",")
         ElseIf nCom == 33 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDPMI,",")
         ElseIf nCom == 34 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDBNF,",")
         ElseIf nCom == 35 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDINSUM,",")
         ElseIf nCom == 36 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDEMBA,",")
         ElseIf nCom == 37 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDSMT,",")
         ElseIf nCom == 38 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDMELE,",")
         ElseIf nCom == 39 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDMFERR,",")
         ElseIf nCom == 40 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDMECA,",")
         ElseIf nCom == 41 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDENER,",")
         ElseIf nCom == 42 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDAGUA,",")
         ElseIf nCom == 43 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDALG,",")
         ElseIf nCom == 44 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDSTRZ,",")
         ElseIf nCom == 45
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDDVG,",")
         ElseIf nCom == 46
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDONAT,",")
         ElseIf nCom == 47 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGINDGMO,",")
         EndIf
      
      ElseIf nOpVis == 4 //Vis?o Administrativa

         If nCom == 1 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMDF,",")
         ElseIf nCom == 2 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMDP,",")
         ElseIf nCom == 3 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMRCS,",")
         ElseIf nCom == 4 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMPRM,",")
         ElseIf nCom == 5 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMCXG,",")
         ElseIf nCom == 6 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMCXNH,",")
         ElseIf nCom == 7 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMINVST,",")
         ElseIf nCom == 8 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMET,",")
         ElseIf nCom == 9 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMEC,",")
         ElseIf nCom == 10 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMIBI,",")
         ElseIf nCom == 11 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMIBM,",")
         ElseIf nCom == 12 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMDVG,",")
         ElseIf nCom == 13 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMADV,",")
         ElseIf nCom == 14 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMSTRZ,",")
         ElseIf nCom == 15 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cXGADMDFIN,",")
         ElseIf nCom == 16 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMOUTNT,",")
         ElseIf nCom == 17 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cGADMMTOR,",")
         EndIf
         
      ElseIf nOpVis == 5 //Vis?o Frete/Custos Prod.Vendidos / Despesas/Receitas Operacionais.
         
         If nCom == 1 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cFRTVND,",")
         ElseIf nCom == 2 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cFRTIMP,",")
         ElseIf nCom == 3 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cFRTEXP,",")
         ElseIf nCom == 4 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cFRTBENF,",")
         ElseIf nCom == 5 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cFRTTRANSF,",")
         ElseIf nCom == 6 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cFRTCOM,",")
         ElseIf nCom == 7 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cFRTIND,",")
         ElseIf nCom == 8 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cFRTADM,",")
         ElseIf nCom == 9 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cCUSPVAP,",")
         ElseIf nCom == 10 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cCUSBENF,",")
         ElseIf nCom == 11
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cCUSINSU,",")
         ElseIf nCom == 12
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cCUSEMBA,",")
         ElseIf nCom == 13 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cCUSSI,",")
         ElseIf nCom == 14 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cCUSMANU,",")
         ElseIf nCom == 15 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cCUSSTRZ,",")
         ElseIf nCom == 16 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cCUSSMT,",")
         ElseIf nCom == 17 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cCUSENRG,",")
         ElseIf nCom == 18 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cCUSAGUA,",")
         ElseIf nCom == 19 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cCUSALG,",")
         ElseIf nCom == 20 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cCUSODIN,",")
         ElseIf nCom == 21 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cDRCOM,",")
         ElseIf nCom == 22 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cDRADM,",")
         ElseIf nCom == 23 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cDRDFIN,",")
         ElseIf nCom == 24 
            cQryCOM += "   AND SE5.E5_NATUREZ IN "+FormatIn(cDRRFIN,",")
         EndIf

      EndIf
           
      cQryCOM += "   AND SE5.D_E_L_E_T_ = ' '"         
      cQryCOM += "   AND SED.D_E_L_E_T_ = ' '"  
      cQryCOM += "   AND SED.ED_FILIAL  = '  '"  
      cQryCOM += "   AND SE5.E5_NATUREZ = SED.ED_CODIGO"
      cQryCOM += "   AND SE5.E5_TIPODOC IN ('VL','BA','PA')"  
      cQryCOM += "   AND SE5.E5_MOTBX <> 'CMP'"
      cQryCOM += "   AND SE5.E5_MOTBX <> 'DEV'" //Five Solutions Consultoria - 25/11/2010 - Solicita??o de Patr?cia  
      cQryCOM += "   AND SE5.E5_SITUACA = ' '"  
      cQryCOM += "   AND SUBSTRING(SE5.E5_NUMERO,1,6)  <> '      '"  
      cQryCOM += "   AND SE5.E5_FILIAL BETWEEN '"+DeFilial+"' AND '"+AteFilial+"'"
      cQryCOM += "   AND SE5.E5_NATUREZ BETWEEN '"+cNatIni+"' AND '"+cNatFim+"'"
      cQryCOM += "   AND SE5.E5_FILIAL NOT IN "+FormatIn(cExclFil,",")
      cQryCOM += "   AND (SELECT COUNT(*)"         
      cQryCOM += "          FROM "+RetSQLName("SE5")+" TMP"       
      cQryCOM += "         WHERE TMP.E5_FILIAL BETWEEN '"+DeFilial+"' AND '"+AteFilial+"'"
      cQryCOM += "           AND TMP.E5_FILIAL NOT IN "+FormatIn(cExclFil,",")
      cQryCOM += "           AND TMP.E5_FILIAL = SE5.E5_FILIAL"         
      cQryCOM += "           AND TMP.E5_PREFIXO = SE5.E5_PREFIXO"         
      cQryCOM += "           AND TMP.E5_NUMERO = SE5.E5_NUMERO"         
      cQryCOM += "           AND TMP.E5_PARCELA = SE5.E5_PARCELA"         
      cQryCOM += "           AND TMP.E5_TIPO = SE5.E5_TIPO"         
      cQryCOM += "           AND TMP.E5_CLIFOR = SE5.E5_CLIFOR"
      cQryCOM += "           AND TMP.E5_LOJA = SE5.E5_LOJA"         
      cQryCOM += "           AND TMP.E5_VALOR = SE5.E5_VALOR"         
      cQryCOM += "           AND TMP.E5_SEQ = SE5.E5_SEQ"         
      cQryCOM += "           AND TMP.E5_TIPODOC = 'ES'"         
      cQryCOM += "           AND TMP.D_E_L_E_T_ <> '*') = 0"
      
      cQryCOM += " ORDER BY E5_FILIAL,xINDICAD,xSUBINDI,xINTRAIN,SE5.E5_NATUREZ,SE5.E5_DATA,"
      cQryCOM += "          E5_PREFIXO,E5_NUMERO,E5_PARCELA,E5_TIPO"
      
      MemoWrite("C:\TEMP\GerencialAnalitico.SQL",cQryCOM)
      
      TCQuery cQryCOM NEW ALIAS "RGAN"
      
      TCSetField("RGAN","E5_VALOR","N",17,02)
      TCSetField("RGAN","E5_DATA","D",08,00)
      
      DbSelectArea("RGAN")
      While RGAN->(!Eof())
         fGravVis() //Grava Vis?o no Arquivo Temporario
         DbSelectArea("RGAN")
         DbSkip()
      EndDo
      DbSelectArea("RGAN")
      DbCloseArea()
      
     
   Next nCom        
   
Return

Static Function fGravVis()
   
   DbSelectArea("XVIS")
   RecLock("XVIS",.T.)
   
      XVIS->INDICADO  := RGAN->xINDICAD
      XVIS->SUBINDIC  := RGAN->xSUBINDI
      XVIS->INTRAIND  := RGAN->xINTRAIN
      XVIS->E5_FILIAL := RGAN->E5_FILIAL
      XVIS->MES       := RGAN->MES
      XVIS->E5_NATUREZ:= RGAN->E5_NATUREZ
      XVIS->ED_DESCRIC:= RGAN->ED_DESCRIC
      XVIS->E5_PREFIXO:= RGAN->E5_PREFIXO
      XVIS->E5_NUMERO := RGAN->E5_NUMERO
      XVIS->E5_PARCELA:= RGAN->E5_PARCELA
      XVIS->E5_TIPO   := RGAN->E5_TIPO
      XVIS->E5_BENEF  := RGAN->E5_BENEF
      XVIS->E5_HISTOR := RGAN->E5_HISTOR
      XVIS->E5_DATA   := RGAN->E5_DATA
      XVIS->E5_VALOR  := RGAN->E5_VALOR
      XVIS->FILMVM    := RGAN->xFILMOV
      XVIS->E5_CLIFOR := RGAN->E5_CLIFOR
      XVIS->E5_LOJA   := RGAN->E5_LOJA
   MsUnLock()
   
Return

Static Function RImpRGA()

   Private oPrint
   oPrint:= TMSPrinter():New( "Relatorio Gerencial Analitico - COMAFAL" )
   oPrint:Setup()
   oPrint:SetLandscape() //ou SetPortrait()

   //oPrint:Setup()
   oPrint:StartPage()   // Inicia uma nova p?gina
   //Fontes
   oFont8  := TFont():New("Arial",9,8 ,.T.,.T.,5,.T.,5,.T.,.F.)
   oFont8n := TFont():New("Arial",9,8 ,.T.,.F.,5,.T.,5,.T.,.F.)
   oFont10 := TFont():New("Arial",9,10,.T.,.T.,5,.T.,5,.T.,.F.)
   oFont10n:= TFont():New("Arial",9,10,.T.,.F.,5,.T.,5,.T.,.F.)
   oFont12 := TFont():New("Arial",9,12,.T.,.T.,5,.T.,5,.T.,.F.)
   oFont12n := TFont():New("Arial",9,12,.T.,.F.,5,.T.,5,.T.,.F.)
   oFont14 := TFont():New("Arial",9,14,.T.,.T.,5,.T.,5,.T.,.F.)
   oFont14n:= TFont():New("Arial",9,14,.T.,.F.,5,.T.,5,.T.,.F.)
   oFont16 := TFont():New("Arial",9,16,.T.,.T.,5,.T.,5,.T.,.F.)
   oFont16n:= TFont():New("Arial",9,16,.T.,.F.,5,.T.,5,.T.,.F.)
   oFont24 := TFont():New("Arial",9,24,.T.,.T.,5,.T.,5,.T.,.F.)
   //oBrush := TBrush():New("",4)//Fundo Cinza
   oBrush := TBrush():New("",CLR_HGRAY)//Fundo Cinza

   DbSelectArea("XVIS")
   nRegImp := RecCount() 
   ProcRegua(nRegImp)
   DbGoTop()
   nCont   := 1
   nTitPag := 0
   nLinha  := 80
   
/*
Say(nRow,nCol,cText,oFont,nWidth,nClrText,nBkMode,nPad)
Imprime um texto no relat?rio.

nRow		Linha para impress?o do texto
nCol		Coluna para impress?o do texto
cText		Texto que sera impresso
oFont		Objeto da classe TFont
nWidth		Tamanho em pixel do texto para impress?o
nClrText	Cor da fonte
nBkMode	Compatibilidade - N?o utilizado
nPad		Compatibilidade - N?o utilizado

*/
   While XVIS->(!Eof())
      
      //IncProc("Impress?o do Registro ... "+Alltrim(Str(nCont))+" / "+Alltrim(Str(nRegImp)))

	  //Impress(oPrint,aDadImp)
	  cIndica := XVIS->INDICADO
	  //Fundo Cinza
	  aCoords := {nLinha,0100,50+nLinha,3000}
      oPrint:FillRect(aCoords,oBrush)
      //Say(nRow,nCol,cText,oFont,nWidth,nClrText,nBkMode,nPad)
	  oPrint:Say  (nLinha,0100,OemToAnsi(cIndica),oFont14n)//,,CLR_HBLUE)
	  nLinha := nLinha + 70

	  If nLinha > 2300 //2500
	     oPrint:EndPage()     // Finaliza a p?gina
	     oPrint:StartPage()   // Inicia uma nova p?gina
	     nLinha  := 80
      EndIf
      
	  //oPrint:Line (nLinha,100,nLinha,2500,oFont14n)//linha 1
	  //nLinha := nLinha + 50
	  nIndicaTt := 0

	  While XVIS->INDICADO == cIndica .And. XVIS->(!Eof())
	     
	     cSubIndc := XVIS->SUBINDIC
	     oPrint:Say  (nLinha,0100,OemToAnsi(cSubIndc),oFont14n)
	     nLinha := nLinha + 70

	     If nLinha > 2300 //2500
	        oPrint:EndPage()     // Finaliza a p?gina
	        oPrint:StartPage()   // Inicia uma nova p?gina
	        nLinha  := 80
         EndIf

	     //oPrint:Line (nLinha,100,nLinha,2500,oFont14n)//linha 1
	     //nLinha := nLinha + 50
	     nSubIndTt := 0
	     
	     While XVIS->INDICADO+XVIS->SUBINDIC == cIndica+cSubIndc .And. XVIS->(!Eof())
	     
	        cIntrIndc := XVIS->INTRAIND
	        oPrint:Say  (nLinha,0100,OemToAnsi(cIntrIndc),oFont12n)
	        nLinha := nLinha + 70
	     
	        //oPrint:Line (nLinha,100,nLinha,2500,oFont12n)//linha 1
	        //nLinha := nLinha + 50
	        nIntraTt := 0
	     	
	     	While XVIS->INDICADO+XVIS->SUBINDIC+XVIS->INTRAIND == cIndica+cSubIndc+cIntrIndc .And. XVIS->(!Eof())
	     	   
	     	   cNatIndc := XVIS->E5_NATUREZ
	     	   nNatTotal := 0
	     	   
	     	   While XVIS->INDICADO+XVIS->SUBINDIC+XVIS->INTRAIND+XVIS->E5_NATUREZ == cIndica+cSubIndc+cIntrIndc+cNatIndc .And. XVIS->(!Eof())
	     	   
	     	      IncProc("Listando Naturezas ... "+Alltrim(Str(nCont))+" / "+Alltrim(Str(nRegImp)))

	              If nLinha > 2300 //2500
	                 oPrint:EndPage()     // Finaliza a p?gina
	                 oPrint:StartPage()   // Inicia uma nova p?gina
	                 nLinha  := 80
	              EndIf
	     	   
                  oPrint:Say  (nLinha,0100,OemToAnsi(XVIS->E5_PREFIXO),oFont8)
                  oPrint:Say  (nLinha,0200,OemToAnsi(XVIS->E5_NUMERO),oFont8)
                  oPrint:Say  (nLinha,0350,OemToAnsi(XVIS->E5_PARCELA),oFont8)
                  oPrint:Say  (nLinha,0400,OemToAnsi(XVIS->E5_TIPO),oFont8)
               
                  oPrint:Say  (nLinha,0500,OemToAnsi(XVIS->E5_BENEF),oFont8)

                  oPrint:Say  (nLinha,1200,OemToAnsi(XVIS->E5_NATUREZ),oFont8)
                  
                  
                  
                  If nHistoric == 1 
                     //Five - 21/01/2011 - cHistEmis := Alltrim(Posicione("SE2",1,XVIS->(FILMVM+E5_PREFIXO+E5_NUMERO+E5_PARCELA+E5_TIPO),"E2_HIST"))
                     cQryHE2 := " SELECT E2_HIST "
                     cQryHE2 += "   FROM "+RetSQLName("SE2")
                     cQryHE2 += "  WHERE E2_FILIAL = '"+XVIS->FILMVM+"'"
                     cQryHE2 += "    AND E2_PREFIXO = '"+XVIS->E5_PREFIXO+"'"
                     cQryHE2 += "    AND E2_NUM = '"+XVIS->E5_NUMERO+"'"
                     cQryHE2 += "    AND E2_PARCELA = '"+XVIS->E5_PARCELA+"'"
                     cQryHE2 += "    AND E2_TIPO = '"+XVIS->E5_TIPO+"'"
                     cQryHE2 += "    AND E2_FORNECE = '"+XVIS->E5_CLIFOR+"'"
                     cQryHE2 += "    AND E2_LOJA = '"+XVIS->E5_LOJA+"'"
                     cQryHE2 += "    AND D_E_L_E_T_ <> '*'"
                     TCQuery cQryHE2 NEW ALIAS "THE2"
                     DbSelectArea("THE2")
                        cHistEmis := THE2->E2_HIST
                     DbCloseArea()
                     
                     //cHistEmis := Alltrim(Posicione("SE2",1,XVIS->FILMVM+XVIS->E5_PREFIXO+XVIS->E5_NUMERO+XVIS->E5_PARCELA+XVIS->E5_TIPO+XVIS->E5_CLIFOR+XVIS->E5_LOJA,"E2_HIST"))
                     //Five - 13/12/2010 - oPrint:Say  (nLinha,1500,OemToAnsi(cHistEmis),oFont8)
                     oPrint:Say  (nLinha,1350,OemToAnsi(cHistEmis),oFont8)
                  Else
                     //Five - 13/12/2010 - oPrint:Say  (nLinha,1500,OemToAnsi(XVIS->E5_HISTOR),oFont8)
                     oPrint:Say  (nLinha,1350,OemToAnsi(XVIS->E5_HISTOR),oFont8)
                  EndIf
               
                  //Five - 13/12/2010 - oPrint:Say  (nLinha,2300,OemToAnsi(DTOC(XVIS->E5_DATA)),oFont8)
                  oPrint:Say  (nLinha,2150,OemToAnsi(DTOC(XVIS->E5_DATA)),oFont8)
                  
                  
                  //oPrint:Say  (nLinha,0180,OemToAnsi(XVIS->ED_DESCRIC),oFont8)
               
                  //Five - 13/12/2010 - oPrint:Say  (nLinha,2500,OemToAnsi(Transform(XVIS->E5_VALOR,"@E 999,999,999.99")),oFont8)
                  oPrint:Say  (nLinha,2350,OemToAnsi(Transform(XVIS->E5_VALOR,"@E 999,999,999.99")),oFont8)
                  
                  oPrint:Say  (nLinha,2550,OemToAnsi(XVIS->ED_DESCRIC),oFont8) //Five - 13/12/2010
                  
                  //Acumula Totalizadores por Indicador
                   nIndicaTt += XVIS->E5_VALOR
                   
                  //Acumula Totalizadores por SubIndicador
                   nSubIndTt += XVIS->E5_VALOR
               
                  //Acumula Totalizadores do IntraIndicador
                  nIntraTt += XVIS->E5_VALOR
                  
                  //Acumula totalizadores por Natureza
                  nNatTotal += XVIS->E5_VALOR

                  
                  nLinha := nLinha + 50
                  
	       
                  DbSelectArea("XVIS")
                  DbSkip()
                  nCont ++
                  
               EndDo
	           //Imprime Totalizadores por Natureza
	           //FatLine()
	           If nLinha > 2300 //2500
	              oPrint:EndPage()     // Finaliza a p?gina
	              oPrint:StartPage()   // Inicia uma nova p?gina
	              nLinha  := 80
	           EndIf
	           
	           oPrint:Line (nLinha,100,nLinha,3000,oFont8n)//linha 1
	           nLinha := nLinha + 25
	           
	           If nLinha > 2300 //2500
	              oPrint:EndPage()     // Finaliza a p?gina
	              oPrint:StartPage()   // Inicia uma nova p?gina
	              nLinha  := 80
	           EndIf
	           
	           //Five - 13/12/2010 - oPrint:Say  (nLinha,2500,OemToAnsi(Transform(nNatTotal,"@E 999,999,999.99")),oFont8n)
	           oPrint:Say  (nLinha,2350,OemToAnsi(Transform(nNatTotal,"@E 999,999,999.99")),oFont8n)
	           nLinha := nLinha + 100
               
            EndDo
            //Imprime Totalizadores por IntraIndicador
            //FatLine()
	        If !Empty(cIntrIndc)
	           oPrint:Line (nLinha,100,nLinha,3000,oFont12n)//linha 1
	           nLinha := nLinha + 25
	           
	           If nLinha > 2300 //2500
	              oPrint:EndPage()     // Finaliza a p?gina
	              oPrint:StartPage()   // Inicia uma nova p?gina
	              nLinha  := 80
	           EndIf
	           
	           //Five - 13/12/2010 - oPrint:Say  (nLinha,2450,OemToAnsi(Transform(nIntraTt,"@E 999,999,999.99")),oFont12n)
	           oPrint:Say  (nLinha,2300,OemToAnsi(Transform(nIntraTt,"@E 999,999,999.99")),oFont12n)
	           nLinha := nLinha + 100
	           
	           If nLinha > 2300 //2500
	              oPrint:EndPage()     // Finaliza a p?gina
	              oPrint:StartPage()   // Inicia uma nova p?gina
	              nLinha  := 80
	           EndIf

	        EndIf
            
         EndDo
         //Imprime Totalizadores por SubIndicador
         //FatLine()
	     If !Empty(cSubIndc)
	        oPrint:Line (nLinha,100,nLinha,3000,oFont14n)//linha 1
	        nLinha := nLinha + 25
	        
	        If nLinha > 2300 //2500
	           oPrint:EndPage()     // Finaliza a p?gina
	           oPrint:StartPage()   // Inicia uma nova p?gina
	           nLinha  := 80
	        EndIf
	           
	        //Five - 13/12/2010 - oPrint:Say  (nLinha,2450,OemToAnsi(Transform(nSubIndTt,"@E 999,999,999.99")),oFont14n)
	        oPrint:Say  (nLinha,2300,OemToAnsi(Transform(nSubIndTt,"@E 999,999,999.99")),oFont14n)
	        nLinha := nLinha + 100

            If nLinha > 2300 //2500
	           oPrint:EndPage()     // Finaliza a p?gina
	           oPrint:StartPage()   // Inicia uma nova p?gina
	           nLinha  := 80
	        EndIf
	           
	     EndIf

         
      EndDo

      //Imprime Totalizadores por Indicador
      //FatLine()
	  oPrint:Line (nLinha,100,nLinha,3000,oFont14n)//linha 1
	  nLinha := nLinha + 25
	  If nLinha > 2300 //2500
	     oPrint:EndPage()     // Finaliza a p?gina
	     oPrint:StartPage()   // Inicia uma nova p?gina
	     nLinha  := 80
	  EndIf
	  //Five - 13/12/2010 - oPrint:Say  (nLinha,2400,OemToAnsi(Transform(nIndicaTt,"@E 999,999,999.99")),oFont14n)
	  oPrint:Say  (nLinha,2250,OemToAnsi(Transform(nIndicaTt,"@E 999,999,999.99")),oFont14n)
	  nLinha := nLinha + 100

      If nLinha > 2300 //2500
	     oPrint:EndPage()     // Finaliza a p?gina
	     oPrint:StartPage()   // Inicia uma nova p?gina
	     nLinha  := 80
	  EndIf	           
	  
   EndDo
   DbSelectArea("XVIS")
   DbCloseArea()
   Ms_Flush()
   oPrint:EndPage()     // Finaliza a p?gina
   oPrint:Preview()     // Visualiza antes de imprimir
   
Return

/*
Static Function Impress(oPrint,aDadImp)
Return
*/


Static Function MontPerg

   Local aArea := GetArea()

   //PutSx1(cGrupo,cOrdem,cPergunt  ,"","",cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
 
   Local aHelpPor := {}
   Local aHelpEng := {}
   Local aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Data Inicial do Periodo do')
   Aadd( aHelpPor, 'Relatorio                           ')
                                                                                                           //F3
   PutSx1(cPerg ,"01","Da Data"  ,"Da Data","Da Data","mv_ch1","D"  ,08       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR01","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Data Final do Periodo do')
   Aadd( aHelpPor, 'Relatorio                           ')
                                                                                                           //F3
   PutSx1(cPerg ,"02","Ate Data"  ,"Ate Data","Ate Data","mv_ch2","D"  ,08       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR02","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe quais vis?es desea visualizar')

                                                                                                           //F3
 //PutSx1(cPerg ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPerg ,"03","Vis?es"  ,"Vis?es","Vis?es","mv_ch3","N"  ,01       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR03","Todas","","",""    ,"Comercial"  ,"","","Industrial","","","Administrativo","","","Despesas/Receitas","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Filial Inicial da Faixa ')
   Aadd( aHelpPor, 'de Filiais do Relatorio.            ')
                                                                                                           //F3
   PutSx1(cPerg ,"04","Da Filial"  ,"Da Filial","Da Filial","mv_ch4","C"  ,02       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR04","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Filial Final da Faixa ')
   Aadd( aHelpPor, 'de Filiais do Relatorio.            ')
                                                                                                           //F3
   PutSx1(cPerg ,"05","Ate Filial"  ,"Ate Filial","Ate Filial","mv_ch5","C"  ,02       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR05","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe separando por virgula as')
   Aadd( aHelpPor, 'Filiais que n?o devem ser apresentadas')
                                                                                                           //F3
   PutSx1(cPerg ,"06","Exceto Filiais"  ,"Exceto Filiais","Exceto Filiais","mv_ch6","C"  ,30       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR06","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe qual historico a ser apresentado')
 
 //PutSx1(cPerg ,"05","Depto.?"            , ""                ,""                 ,"mv_ch5","C"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos"  ,"","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")                                                                                                           //F3
   PutSx1(cPerg ,"07","Imprime Historico"  ,"Imprime Historico","Imprime Historico","mv_ch7","N"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR07","Emissao","","",""    ,"Baixa","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Natureza Inicial da Faixa ')
   Aadd( aHelpPor, 'de Naturezas do Relatorio.            ')
 //PutSx1(cGrupo,cOrdem,cPergunt       ,""           ,""           ,cVar    ,cTipo,nTamanho ,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
   PutSx1(cPerg ,"08"  ,"Da Natureza"  ,"Da Natureza","Da Natureza","mv_ch8","C"  ,10       ,0       ,0      ,"G" ,""    ,"SED" ,""     ,"","MV_PAR08","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Natureza Final da Faixa ')
   Aadd( aHelpPor, 'de Naturezas do Relatorio.            ')
 //PutSx1(cGrupo,cOrdem,cPergunt       ,""           ,""           ,cVar    ,cTipo,nTamanho ,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
   PutSx1(cPerg ,"09"  ,"Ate Natureza"  ,"Ate Natureza","Ate Natureza","mv_ch9","C"  ,10       ,0       ,0      ,"G" ,""    ,"SED" ,""     ,"","MV_PAR09","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   RestArea(aArea)

Return(.T.)