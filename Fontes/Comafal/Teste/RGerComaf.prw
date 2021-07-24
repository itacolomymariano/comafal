#Include "PROTHEUS.CH"
#Include "TopConn.Ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³RGERCOMAF ºAutor  ³Five Solutions      º Data ³  05/01/2008 º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Relatório de Resumo Gerencial COMAFAL                      º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ COMAFAL - PE, SP e RS.                                    º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function RGerComaf

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Declaracao de Variaveis                                             ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

Local cDesc1         := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2         := "de acordo com os parametros informados pelo usuario."
Local cDesc3         := "Resumo Gerencial COMAFAL"
Local cPict          := ""
Local titulo       := "Resumo Gerencial COMAFAL"
Local nLin         := 80

Local Cabec1       := ""//Mês                               Mês 01      Mês 02        Mês 03"
Local Cabec2       := ""
Local imprime      := .T.
Local aOrd := {}
Private lEnd         := .F.
Private lAbortPrint  := .F.
Private CbTxt        := ""
Private limite           := 220
Private tamanho          := "G"
Private nomeprog         := "RGerComaf" // Coloque aqui o nome do programa para impressao no cabecalho
Private nTipo            := 15
Private aReturn          := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey        := 0
Private cPerg       := "RGERCO"
//Private cPerg      	:= "CRF010"
Private cbtxt      := Space(10)
Private cbcont     := 00
Private CONTFL     := 01
Private m_pag      := 01
Private wnrel      := "RGerComaf" // Coloque aqui o nome do arquivo usado para impressao em disco
Private dDataDe    := CTOD("")
Private dDataAte   := CTOD("")
Private d2AnoMes   := ""
Private d3AnoMes   := ""


Private cString := "SA1"

dbSelectArea("SA1")
dbSetOrder(1)

AjustaSx1()
pergunte(cPerg,.F.)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Monta a interface padrao com o usuario...                           ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

wnrel := SetPrint(cString,NomeProg,cPerg,@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,.F.,Tamanho,,.F.)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Grupo de Perguntas - RGERCO                                         ³
//³ MV_PAR01 - Data De                                                  ³
//³ MV_PAR02 - Data Até                                                 ³
//³ MV_PAR03 - Custo - (1.Médio, 2.Reposição                            ³
//³ MV_PAR04 - Analítico - (1.Sim, 2.Não)                               ³
//³ MV_PAR05 - Depto - (1.Todos,2.Faturamento,3.Indústria,4.Financeiro  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

dDataDe  := MV_PAR01
dDataAte := MV_PAR02
nDeptos  := MV_PAR05

lComercial := IIF(nDeptos == 1 .Or. nDeptos == 2,.T.,.F.)

titulo       := "Resumo Gerencial COMAFAL Período "+DtoC(dDataDe)+" a "+DtoC(dDataAte)

n2ResMes := Val(Substr(DTOS(dDataDe),1,6)) + 1
d2AnoMes := Alltrim(Str(n2ResMes))
If Right(Alltrim(Str(n2ResMes)),2) == "13"
   cAno := Val(Substr(DTOS(dDataDe),1,4)) + 1
   d2AnoMes   := Alltrim(Str(cAno))+"01"
EndIf

n3ResMes := Val(Substr(DTOS(dDataDe),1,6)) + 2
d3AnoMes   := Alltrim(Str(n3ResMes,2))
If Right(Alltrim(Str(n3ResMes)),2) == "13"
   cAno := Val(Substr(DTOS(dDataDe),1,4)) + 1
   d3AnoMes   := Alltrim(Str(cAno))+"01"
EndIf

If nLastKey == 27
	Return
Endif

SetDefault(aReturn,cString)

If nLastKey == 27
   Return
Endif

nTipo := If(aReturn[4]==1,15,18)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Processamento. RPTSTATUS monta janela com a regua de processamento. ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

RptStatus({|| RunReport(Cabec1,Cabec2,Titulo,nLin) },Titulo)
Return

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFun‡„o    ³RUNREPORT º Autor ³ AP6 IDE            º Data ³  05/01/08   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescri‡„o ³ Funcao auxiliar chamada pela RPTSTATUS. A funcao RPTSTATUS º±±
±±º          ³ monta a janela com a regua de processamento.               º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ Programa principal                                         º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/

Static Function RunReport(Cabec1,Cabec2,Titulo,nLin)

Local nOrdem

dbSelectArea(cString)
dbSetOrder(1)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Área de Instruções SQL - Querys.                                    ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ


If lComercial

   /**************
   *
   * Calculando Clientes (Clientes)
   *
   ***************/

   For nMes := 1 To 9
      
      If nMes <= 6
         cQrySF2 := "SELECT COUNT(*) AS CLIENTES"
      ElseIf nMes >= 7
         cQrySF2 := "SELECT COUNT(DISTINCT F2_CLIENTE+' '+F2_LOJA) AS CLIENTES"
      EndIf
      cQrySF2 += " FROM "+RetSQLName("SF2")+" SF2"
      cQrySF2 += " WHERE SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      If nMes <= 3
         If nMes == 1
            cQrySF2 += "   AND SUBSTRING(SF2.F2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         ElseIf nMes == 2
            cQrySF2 += "   AND SUBSTRING(SF2.F2_EMISSAO,1,6) = '"+d2AnoMes+"'"
         ElseIf nMes == 3
            cQrySF2 += "   AND SUBSTRING(SF2.F2_EMISSAO,1,6) = '"+d3AnoMes+"'"
         EndIf
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            cQrySF2 += "   AND SUBSTRING(SF2.F2_EMISSAO,1,6) BETWEEN '"+Substr(DTOS(dDataDe - 180),1,6)+"' AND '"+Substr(DTOS(dDataDe),1,6)+"'"
         ElseIf nMes == 5
            dData02 := CTOD("01/"+Substr(d2AnoMes,5,2)+"/"+Substr(d2AnoMes,1,4))
            cQrySF2 += "   AND SUBSTRING(SF2.F2_EMISSAO,1,6) BETWEEN '"+Substr(DTOS(dData02 - 180),1,6)+"' AND '"+Substr(DTOS(dData02),1,6)+"'"
         ElseIf nMes == 6
            dData03 := CTOD("01/"+Substr(d3AnoMes,5,2)+"/"+Substr(d3AnoMes,1,4))
            cQrySF2 += "   AND SUBSTRING(SF2.F2_EMISSAO,1,6) BETWEEN '"+Substr(DTOS(dData03 - 180),1,6)+"' AND '"+Substr(DTOS(dData03),1,6)+"'"
         EndIf      
      EndIf
      
      If nMes <= 3 
         cQrySF2 += "   AND (SELECT COUNT(*)"
         cQrySF2 += "        FROM "+RetSQLName("SF2")+" TSF2"
         cQrySF2 += "        WHERE TSF2.F2_FILIAL = '"+xFilial("SF2")+"'"
         cQrySF2 += "          AND TSF2.F2_CLIENTE = SF2.F2_CLIENTE"
         cQrySF2 += "          AND TSF2.F2_LOJA = SF2.F2_LOJA"
         If nMes == 1
            cQrySF2 += "          AND SUBSTRING(TSF2.F2_EMISSAO,1,6) < '"+Substr(DTOS(dDataDe),1,6)+"'"
         ElseIf nMes == 2
            cQrySF2 += "          AND SUBSTRING(TSF2.F2_EMISSAO,1,6) < '"+d2AnoMes+"'"
         ElseIf nMes == 3
            cQrySF2 += "          AND SUBSTRING(TSF2.F2_EMISSAO,1,6) < '"+d2AnoMes+"'"
         EndIf
         cQrySF2 += "          AND TSF2.D_E_L_E_T_ <> '*') = 0"
      EndIf
      
      If nMes >= 7
         If nMes == 7
            cQrySF2 += "   AND SUBSTRING(SF2.F2_EMISSAO,1,6) = '"+Substr(DTOS(dDataDe),1,6)+"'"
         ElseIf nMes == 8
            cQrySF2 += "   AND SUBSTRING(SF2.F2_EMISSAO,1,6) = '"+d2AnoMes+"'"
         ElseIf nMes == 9
            cQrySF2 += "   AND SUBSTRING(SF2.F2_EMISSAO,1,6) = '"+d3AnoMes+"'"
         EndIf
      EndIf
      
      cQrySF2 += "   AND SF2.D_E_L_E_T_ <> '*'"

      TCQuery cQrySF2 NEW ALIAS "TCLI"
   
      If nMes <= 3
         If nMes == 1
            DbSelectArea("TCLI")
            nCliN01Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf

         If nMes == 2
            DbSelectArea("TCLI")
            nCliN02Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf

         If nMes == 3
            DbSelectArea("TCLI")
            nCliN03Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf   
      ElseIf nMes >= 4 .And. nMes <= 6
         If nMes == 4
            DbSelectArea("TCLI")
            nClAtv1Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf

         If nMes == 5
            DbSelectArea("TCLI")
            nClAtv2Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf

         If nMes == 6
            DbSelectArea("TCLI")
            nClAtv3Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf   
      ElseIf nMes >= 7 .And. nMes <= 9
         If nMes == 7
            DbSelectArea("TCLI")
            nClAtd1Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf

         If nMes == 8
            DbSelectArea("TCLI")
            nClAtd2Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf

         If nMes == 9
            DbSelectArea("TCLI")
            nClAtd3Mes := TCLI->CLIENTES
            DbCloseArea()
         EndIf   
      EndIf

   Next nMes
   For nMes := 1 To 3
      cQrySA1 := "SELECT COUNT(*) AS CLIENTES"
      cQrySA1 += " FROM "+RetSQLName("SA1")+" SA1"
      cQrySA1 += " WHERE SA1.A1_FILIAL = '"+xFilial("SA1")+"'"
      cQrySA1 += " AND (SELECT COUNT(*)"
      cQrySA1 += "      FROM "+RetSQLName("SF2")+" SF2"
      cQrySA1 += "      WHERE SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      If nMes == 1
         cQrySA1 += "     AND SUBSTRING(SF2.F2_EMISSAO,1,6) BETWEEN '"+Substr(DTOS(dDataDe - 180),1,6)+"' AND '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQrySA1 += "     AND SUBSTRING(SF2.F2_EMISSAO,1,6) BETWEEN '"+Substr(DTOS(dData02 - 180),1,6)+"' AND '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQrySA1 += "     AND SUBSTRING(SF2.F2_EMISSAO,1,6) BETWEEN '"+Substr(DTOS(dData03 - 180),1,6)+"' AND '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQrySA1 += "        AND SF2.F2_CLIENTE = SA1.A1_COD"
      cQrySA1 += "        AND SF2.F2_LOJA = SA1.A1_LOJA"
      cQrySA1 += "        AND SF2.D_E_L_E_T_ <> '*') = 0 "

      cQrySA1 += " AND (SELECT COUNT(*)"
      cQrySA1 += "      FROM "+RetSQLName("SF2")+" SF2"
      cQrySA1 += "      WHERE SF2.F2_FILIAL = '"+xFilial("SF2")+"'"
      If nMes == 1
         cQrySA1 += "     AND SUBSTRING(SF2.F2_EMISSAO,1,6) > '"+Substr(DTOS(dDataDe),1,6)+"'"
      ElseIf nMes == 2
         cQrySA1 += "     AND SUBSTRING(SF2.F2_EMISSAO,1,6) > '"+Substr(DTOS(dData02),1,6)+"'"
      ElseIf nMes == 3
         cQrySA1 += "     AND SUBSTRING(SF2.F2_EMISSAO,1,6) > '"+Substr(DTOS(dData03),1,6)+"'"
      EndIf
      cQrySA1 += "        AND SF2.F2_CLIENTE = SA1.A1_COD"
      cQrySA1 += "        AND SF2.F2_LOJA = SA1.A1_LOJA"
      cQrySA1 += "        AND SF2.D_E_L_E_T_ <> '*') = 0 "

      cQrySA1 += " AND SA1.D_E_L_E_T_ <> '*'"
      
      //MemoWrite("cQrySA1.SQL"+Alltrim(Str(nMes)),cQrySA1)
      
      TCQuery cQrySA1 NEW ALIAS "TCLI"
      
      If nMes == 1
         DbSelectArea("TCLI")
         nClIna1Mes := TCLI->CLIENTES
         DbCloseArea()
      EndIf

      If nMes == 2
         DbSelectArea("TCLI")
         nClIna2Mes := TCLI->CLIENTES
         DbCloseArea()
      EndIf

      If nMes == 3
         DbSelectArea("TCLI")
         nClIna3Mes := TCLI->CLIENTES
         DbCloseArea()
      EndIf   
   Next nMes  

EndIf

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ SETREGUA -> Indica quantos registros serao processados para a regua ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

//SetRegua(RecCount())

//dbGoTop()
lRoda := .T.
While lRoda//!EOF()

   //ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
   //³ Verifica o cancelamento pelo usuario...                             ³
   //ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

   If lAbortPrint
      @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
      Exit
   Endif

   //ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
   //³ Impressao do cabecalho do relatorio. . .                            ³
   //ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

   If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 8
   Endif

   If lComercial
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|COMERCIAL"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //                       10        20        30        40
      //              01234567890123456789012345678901234567890
      @nLin,000 PSAY "|MES                                |"
      @nLin,037 PSAY fNomeMes(Alltrim(StrZero(Month(dDataDe),2)))
      @nLin,072 PSAY "|"
      @nLin,073 PSAY fNomeMes(Right(d2AnoMes,2))
      @nLin,108 PSAY "|"
      @nLin,109 PSAY fNomeMes(Right(d3AnoMes,2))
      @nLin,144 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"
      nLin ++
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "|CLIENTES                           |"
      @nLin,050 PSAY  (nClAtv1Mes + nClIna1Mes ) Picture "@E 999,999"
      @nLin,072 PSAY "|"
      @nLin,087 PSAY  (nClAtv2Mes + nClIna2Mes ) Picture "@E 999,999"
      @nLin,108 PSAY "|"
      @nLin,124 PSAY  (nClAtv3Mes + nClIna3Mes ) Picture "@E 999,999"
      @nLin,144 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      //                       10        20        30        40        50        60        70        80        90       100       110       120       130       140       150
      //              0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
      //              |                                   |                                   |                                   |                                   |
      //                                                                999,999                              999,999                              999,999                 
      @nLin,000 PSAY "| * Novos                           |"
      @nLin,050 PSAY nCliN01Mes Picture "@E 999,999" 
      @nLin,072 PSAY "|"
      @nLin,087 PSAY nCliN02Mes Picture "@E 999,999" 
      @nLin,108 PSAY "|"
      @nLin,124 PSAY nCliN03Mes Picture "@E 999,999" 
      @nLin,144 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Ativos                          |"
      @nLin,050 PSAY nClAtv1Mes Picture "@E 999,999" 
      @nLin,072 PSAY "|"
      @nLin,087 PSAY nClAtv2Mes Picture "@E 999,999" 
      @nLin,108 PSAY "|"
      @nLin,124 PSAY nClAtv3Mes Picture "@E 999,999" 
      @nLin,144 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Inativos                        |"
      @nLin,050 PSAY nClIna1Mes Picture "@E 999,999" 
      @nLin,072 PSAY "|"
      @nLin,087 PSAY nClIna2Mes Picture "@E 999,999" 
      @nLin,108 PSAY "|"
      @nLin,124 PSAY nClIna3Mes Picture "@E 999,999" 
      @nLin,144 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * Atendidos                       |"
      @nLin,050 PSAY nClAtd1Mes Picture "@E 999,999" 
      @nLin,072 PSAY "|"
      @nLin,087 PSAY nClAtd2Mes Picture "@E 999,999" 
      @nLin,108 PSAY "|"
      @nLin,124 PSAY nClAtd3Mes Picture "@E 999,999" 
      @nLin,144 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif      
      @nLin,000 PSAY "| * % Atendimento                   |"
      @nLin,051 PSAY nClAtd1Mes / (nClAtv1Mes + nClIna1Mes) Picture "@E 999.99" 
      @nLin,072 PSAY "|"
      @nLin,088 PSAY nClAtd2Mes / (nClAtv2Mes + nClIna2Mes) Picture "@E 999.99" 
      @nLin,108 PSAY "|"
      @nLin,125 PSAY nClAtd3Mes / (nClAtv3Mes + nClIna3Mes) Picture "@E 999.99" 
      @nLin,144 PSAY "|"
      @nLin,219 PSAY "|"
      nLin ++
      If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 8
      Endif
      @nLin,000 PSAY "|"+Replicate("=",218)+"|"

   EndIf
   
   nLin ++ // Avanca a linha de impressao

   //dbSkip() // Avanca o ponteiro do registro no arquivo
   lRoda := .F.
EndDo

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Finaliza a execucao do relatorio...                                 ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

SET DEVICE TO SCREEN

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Se impressao em disco, chama o gerenciador de impressao...          ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

If aReturn[5]==1
   dbCommitAll()
   SET PRINTER TO
   OurSpool(wnrel)
Endif

MS_FLUSH()
Return

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³                                     F U N Ç Õ E S   G E N É R I C A S                                                        ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ


Static Function AjustaSx1

Local aArea := GetArea()

//PutSx1(cGrupo,cOrdem,cPergunt  ,"","",cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
  PutSx1(cPerg ,"01","Data de ?" ,"","","mv_ch1","D"  ,8       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR01",""     ,"","",""    ,""           ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"02","Data ate ?","","","mv_ch2","D"  ,8       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR02",""     ,"","",""    ,""           ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"03","Custo ?"   ,"","","mv_ch3","C"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR03","Medio","","",""    ,"Reposicao"  ,"","","Calculado","","",""          ,"","","","","")
  PutSx1(cPerg ,"04","Analitico?","","","mv_ch4","C"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR04","Sim"  ,"","",""    ,"Nao"        ,"","",""         ,"","",""          ,"","","","","")
  PutSx1(cPerg ,"05","Depto.?"   ,"","","mv_ch5","C"  ,1       ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria","","","Financeiro","","","Gerencial","","")

RestArea(aArea)

Return(.T.)

Static Function fNomeMes(cMes)
   cNomeMes := ""
   If cMes == "01"
      cNomeMes := "JANEIRO"
   ElseIf cMes == "02"
      cNomeMes := "FEVEREIRO"
   ElseIf cMes == "03"
      cNomeMes := "MARÇO"
   ElseIf cMes == "04"
      cNomeMes := "ABRIL"
   ElseIf cMes == "05"
      cNomeMes := "MAIO"
   ElseIf cMes == "06"
      cNomeMes := "JUNHO"
   ElseIf cMes == "07"
      cNomeMes := "JULHO"
   ElseIf cMes == "08"
      cNomeMes := "AGOSTO"
   ElseIf cMes == "09"
      cNomeMes := "SETEMBBRO"
   ElseIf cMes == "10"
      cNomeMes := "OUTUBRO"
   ElseIf cMes == "11"
      cNomeMes := "NOVEMBRO"
   ElseIf cMes == "12"
      cNomeMes := "DEZEMBRO"
   EndIf
Return(cNomeMes)