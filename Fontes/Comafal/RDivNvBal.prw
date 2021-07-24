#include "Rwmake.ch"
#include "TopConn.Ch"
/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³RDivNvBal º Autor ³ Five Solutions     º Data ³  19/08/2010 º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescricao ³ Relação de Divergencia de Pesos entre informado NF(Navio) eº±±
±±º          ³ Balança COMAFAL.                                           º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ COMAFAL - PE,SP e SP                                       º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/

User Function RDivNvBal


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Declaracao de Variaveis                                             ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

Local cDesc1         := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2         := "de acordo com os parametros informados pelo usuario."
Local cDesc3         := "Divergencia Peso NF(Navio) X Peso Real(Balança COMAFAL)"
Local cPict          := ""
Local titulo       := "Divergencia Peso NF(Navio) X Peso Real(Balança COMAFAL)"
Local nLin         := 80
Local Cabec1       := "Codigo           Descrição"
Local Cabec2       := "    Navio                                    D.I.          N.Fiscal Ser Peso NF(Navio)   Peso COMAFAL      Diferença Dt.Ent(MP) Dt.Ent.NF Transportador                             Fornecedor"
Local imprime      := .T.
Local aOrd := {}
Private lEnd         := .F.
Private lAbortPrint  := .F.
Private CbTxt        := ""
Private limite           := 220
Private tamanho          := "G"
Private nomeprog         := "RDivNvBal" // Coloque aqui o nome do programa para impressao no cabecalho
Private nTipo            := 15
Private aReturn          := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey        := 0
Private cPrgD       := "RDVNNF"
Private cbtxt      := Space(10)
Private cbcont     := 00
Private CONTFL     := 01
Private m_pag      := 01
Private wnrel      := "RDivNvBal" // Coloque aqui o nome do arquivo usado para impressao em disco

Private cString := "SZB"

dbSelectArea("SZB")
dbSetOrder(1)

MTPergDV()

Pergunte(cPrgD,.F.)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Monta a interface padrao com o usuario...                           ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

wnrel := SetPrint(cString,NomeProg,cPrgD,@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,.F.,Tamanho,,.F.)

   Private DoNavio := MV_PAR01
   Private AteNavi := MV_PAR02
   Private DaDI    := MV_PAR03
   Private AteDI   := MV_PAR04
   Private DaTrans := MV_PAR05
   Private AteTran := MV_PAR06
   Private DoForne := MV_PAR07
   Private AteForn := MV_PAR08
   Private DaLoja  := MV_PAR09
   Private AteLoja := MV_PAR10
   Private DaData  := MV_PAR11
   Private AteData := MV_PAR12

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
±±ºFun‡„o    ³RUNREPORT º Autor ³ AP6 IDE            º Data ³  19/08/10   º±±
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
//³ QUERY DE SELEÇÃO DOS REGISTROS DO RELATORIO                         ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

For nImp := 1 To 2
   
   If nImp == 1
      cQrySZB := " SELECT COUNT(*) NREGSZB "
   Else
      cQrySZB := " SELECT SD1.D1_COD,SZB.ZB_DESNAV,SZB.ZB_DI,SZB.ZB_PESONF,SZ5.Z5_TRANSP,SZ5.Z5_FORNECE, "
      cQrySZB += "        SZB.ZB_DTENT,SF1.F1_DTDIGIT,SZB.ZB_DOC,SZB.ZB_SERIE,"
      cQrySZB += "        (SELECT SUM(TMD1.D1_QUANT) "
      cQrySZB += "           FROM "+RetSQLName("SD1")+" TMD1 "
      cQrySZB += "          WHERE TMD1.D1_FILIAL = '"+xFilial("SD1")+"'"
      cQrySZB += "            AND TMD1.D1_DOC = SF1.F1_DOC"
      cQrySZB += "            AND TMD1.D1_SERIE = SF1.F1_SERIE"
      cQrySZB += "            AND TMD1.D1_FORNECE = SF1.F1_FORNECE"
      cQrySZB += "            AND TMD1.D1_LOJA = SF1.F1_LOJA"
      cQrySZB += "            AND TMD1.D_E_L_E_T_ <> '*') PESOREAL "
   EndIf
   
   cQrySZB += "   FROM "+RetSQLName("SZ5")+" SZ5,"+RetSQLName("SZB")+" SZB, "+RetSQLName("SF1")+" SF1,"
   cQrySZB += RetSQLName("SD1")+" SD1"
   cQrySZB += "  WHERE SZ5.Z5_FILIAL = '"+xFilial("SZ5")+"'"
   cQrySZB += "    AND SZB.ZB_FILIAL = '"+xFilial("SZB")+"'"
   cQrySZB += "    AND SF1.F1_FILIAL = '"+xFilial("SF1")+"'"
   cQrySZB += "    AND SD1.D1_FILIAL = '"+xFilial("SD1")+"'"

   cQrySZB += "    AND SZB.ZB_NAVIO BETWEEN '"+DoNavio+"' AND '"+AteNavi+"'"
   cQrySZB += "    AND SZB.ZB_DI BETWEEN '"+DaDI+"' AND '"+AteDI+"'"
   cQrySZB += "    AND SZ5.Z5_TRANSP BETWEEN '"+DaTrans+"' AND '"+AteTran+"'"
   cQrySZB += "    AND SZ5.Z5_FORNECE BETWEEN '"+DoForne+"' AND '"+AteForn+"'"
   cQrySZB += "    AND SZ5.Z5_LOJA BETWEEN '"+DaLoja+"' AND '"+AteLoja+"'"
   cQrySZB += "    AND SF1.F1_DTDIGIT  BETWEEN '"+DTOS(DaData)+"' AND '"+DTOS(AteData)+"'"
   
   cQrySZB += "    AND SZB.ZB_DOC = SF1.F1_DOC"
   cQrySZB += "    AND SZB.ZB_SERIE = SF1.F1_SERIE" 
   cQrySZB += "    AND SZ5.Z5_PLACA = SZB.ZB_PLACA "
   cQrySZB += "    AND SZ5.Z5_DENTR = SZB.ZB_DTENT"
   cQrySZB += "    AND SZ5.Z5_HENTR = SZB.ZB_HRENT"
   cQrySZB += "    AND SF1.F1_DOC = SD1.D1_DOC"
   cQrySZB += "    AND SF1.F1_SERIE = SD1.D1_SERIE"
   cQrySZB += "    AND SF1.F1_FORNECE = SD1.D1_FORNECE"
   cQrySZB += "    AND SF1.F1_LOJA = SD1.D1_LOJA"
   
   cQrySZB += "    AND SZB.D_E_L_E_T_ <> '*'"
   cQrySZB += "    AND SZ5.D_E_L_E_T_ <> '*'"
   cQrySZB += "    AND SF1.D_E_L_E_T_ <> '*'"
   cQrySZB += "    AND SD1.D_E_L_E_T_ <> '*'"
   
   If nImp > 1
      cQrySZB += " ORDER BY SD1.D1_COD "
   EndIf
   
   MemoWrite("C:\TEMP\DivNvXBal.SQL",cQrySZB)
   
   TCQuery cQrySZB NEW ALIAS "TSZB"
   
   If nImp == 1
      TCSetField("TSZB","NREGSZB","N",10,00)
      DbSelectArea("TSZB")
      nReg := TSZB->NREGSZB
      DbCloseArea()
   EndIf
   
Next nImp

TCSetField("TSZB","ZB_PESONF","N",17,03)
TCSetField("TSZB","PESOREAL","N",17,03)
TCSetField("TSZB","ZB_DTENT","D",08,00)
TCSetField("TSZB","F1_DTDIGIT","D",08,00)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ SETREGUA -> Indica quantos registros serao processados para a regua ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

SetRegua(nReg)

//Totais Gerais
nTPESONF := 0
nTPESOREA:= 0
nTtDifPes:= 0

DbSelectArea("TSZB")
While TSZB->(!EOF())

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
      nLin := 9
   Endif
      
   cProdNv := TSZB->D1_COD
   cDscProd := Substr(Posicione("SB1",1,xFilial("SB1")+cProdNv,"B1_DESC"),1,60)
   @nLin,000 Psay cProdNv
   @nLin,017 Psay cDscProd
   nLin ++
   nLin ++

   //Totais Produto
   nPrTPESONF := 0
   nPrTPESOREA:= 0
   nPrTtDifPes:= 0
   

   While TSZB->D1_COD == cProdNv .And. TSZB->(!EOF()) 
   
      IncRegua()
   
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
         nLin := 9
      Endif

      //           10        20        30        40        50        60        70        80        90       100       110       120       130       140       150       160       170       180       190       200       210       220       230
      //  01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
      // "Codigo           Descrição"
      // "    Navio                                    D.I.          N.Fiscal Ser Peso NF(Navio)   Peso COMAFAL      Diferença Dt.Ent(MP) Dt.Ent.NF Transportador                             Fornecedor"
      //  xxxxxxxxxXxxxxx  xxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxXXxxxxxxxxxXXxxxxxxxxxX
      //      xxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxX 99/9999999-9 999999999 999  9.999,999.999  9.999,999.999  9.999,999.999   99/99/99  99/99/99 xxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxX  xxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxX 
   
      @nLin,004 PSAY Substr(TSZB->ZB_DESNAV,1,40) 
      @nLin,045 PSAY TSZB->ZB_DI
      @nLin,058 PSAY TSZB->ZB_DOC
      @nLin,068 PSAY TSZB->ZB_SERIE
      @nLin,073 PSAY TSZB->ZB_PESONF Picture "@E 9,999,999.999"
      @nLin,088 PSAY TSZB->PESOREAL Picture "@E 9,999,999.999"
      //nDifPeso := ABS(TSZB->ZB_PESONF - TSZB->PESOREAL)
      nDifPeso := (TSZB->PESOREAL - TSZB->ZB_PESONF)
      @nLin,103 PSAY nDifPeso Picture "@E 9,999,999.999"
      @nLin,119 PSAY TSZB->ZB_DTENT
      @nLin,129 PSAY TSZB->F1_DTDIGIT
      cDescTra := Alltrim(Posicione("SA4",1,xFilial("SA4")+TSZB->Z5_TRANSP,"A4_NOME"))
      @nLin,138 PSAY Substr(cDescTra,1,40)
      cDescFor := Alltrim(Posicione("SA2",1,xFilial("SA2")+TSZB->Z5_FORNECE,"A2_NOME"))
      @nLin,180 PSAY Substr(cDescFor,1,40)

      //Totais por Produto
      nPrTPESONF  += TSZB->ZB_PESONF
      nPrTPESOREA += TSZB->PESOREAL
      nPrTtDifPes += nDifPeso
            
      //Totais Gerais
      nTPESONF  += TSZB->ZB_PESONF
      nTPESOREA += TSZB->PESOREAL
      nTtDifPes += nDifPeso
   
      nLin := nLin + 1 // Avanca a linha de impressao
      DbSelectArea("TSZB") 
      DbSkip() // Avanca o ponteiro do registro no arquivo
   
   EndDo

   nLin := nLin + 1 // Avanca a linha de impressao
   If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 9
   Endif
   @nLin,000 PSAY "Totais Produto "+cProdNv+" ..."
   @nLin,073 PSAY nPrTPESONF Picture "@E 9,999,999.999"
   @nLin,088 PSAY nPrTPESOREA Picture "@E 9,999,999.999"
   //@nLin,103 PSAY nPrTtDifPes Picture "@E 9,999,999.999"
   @nLin,103 PSAY (nPrTPESOREA - nPrTPESONF) Picture "@E 9,999,999.999"
   
   nLin ++
   nLin ++
   

EndDo

nLin ++
nLin := nLin + 1 // Avanca a linha de impressao
If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
   Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
   nLin := 9
Endif
@nLin,000 PSAY "Totais Gerais ... "   
@nLin,073 PSAY nTPESONF Picture "@E 9,999,999.999"
@nLin,088 PSAY nTPESOREA Picture "@E 9,999,999.999"
//@nLin,103 PSAY nTtDifPes Picture "@E 9,999,999.999"
@nLin,103 PSAY (nTPESOREA - nTPESONF) Picture "@E 9,999,999.999"
   
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Finaliza a execucao do relatorio...                                 ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
DbSelectArea("TSZB")
DbCloseArea()

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

Static Function MTPergDV

   Local aArea := GetArea()

   //PutSx1(cGrupo,cOrdem,cPergunt  ,"","",cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
 
   Local aHelpPor := {}
   Local aHelpEng := {}
   Local aHelpSpa := {}

   Aadd( aHelpPor, 'Informe o codigo inicial do Navio da')
   Aadd( aHelpPor, 'faixa de Navios a serem apresentados.')
                                                                                                           //F3
   PutSx1(cPrgD ,"01","Do Navio"  ,"Do Navio","Do Navio","mv_ch1","C"  ,06       ,0       ,0      ,"G" ,""    ,"9B" ,""     ,"","MV_PAR01","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe o codigo final do Navio da')
   Aadd( aHelpPor, 'faixa de Navios a serem apresentados.')

                                                                                                           //F3
   PutSx1(cPrgD ,"02","Ate Navio"  ,"Ate Navio","Ate Navio","mv_ch2","C"  ,06       ,0       ,0      ,"G" ,""    ,"9B" ,""     ,"","MV_PAR02","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe o codigo D.I. inicial')
   Aadd( aHelpPor, 'Declaração de Importação inicial')


                                                                                                           //F3
 //PutSx1(cPrgD ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPrgD ,"03","Do D.I."  ,"Do D.I.","Do D.I.","mv_ch3","C"  ,06       ,0       ,0      ,"G" ,""    ,"9C" ,""     ,"","MV_PAR03","","","",""    ,""  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe o codigo D.I. final')
   Aadd( aHelpPor, 'Declaração de Importação final')

                                                                                                           //F3
 //PutSx1(cPrgD ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPrgD ,"04","Ate D.I."  ,"Ate D.I.","Ate D.I.","mv_ch4","C"  ,6       ,0       ,0      ,"G" ,""    ,"9C" ,""     ,"","MV_PAR04","","","",""    ,""  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}
   
   Aadd( aHelpPor, 'Informe o codigo da transportadora')
   Aadd( aHelpPor, 'inicial da faixa de transportadoras')

                                                                                                           //F3
 //PutSx1(cPrgD ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPrgD ,"05","Da Transportadora"  ,"Da Transportadora","Da Transportadora","mv_ch5","C"  ,6       ,0       ,0      ,"G" ,""    ,"SA4" ,""     ,"","MV_PAR05","","","",""    ,""  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}
   
   Aadd( aHelpPor, 'Informe o codigo da transportadora')
   Aadd( aHelpPor, 'final da faixa de transportadoras')

                                                                                                           //F3
 //PutSx1(cPrgD ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPrgD ,"06","Ate Transportadora"  ,"Ate Transportadora","Ate Transportadora","mv_ch6","C"  ,6       ,0       ,0      ,"G" ,""    ,"SA4" ,""     ,"","MV_PAR06","","","",""    ,""  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
      

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}
   
   Aadd( aHelpPor, 'Informe o codigo do Fornecedor')
   Aadd( aHelpPor, 'inicial da faixa de Fornecedores')

                                                                                                           //F3
 //PutSx1(cPrgD ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPrgD ,"07","Do Fornecedor"  ,"Do Fornecedor","Do Fornecedor","mv_ch7","C"  ,6       ,0       ,0      ,"G" ,""    ,"SA2" ,""     ,"","MV_PAR07","","","",""    ,""  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}
   
   Aadd( aHelpPor, 'Informe o codigo do Fornecedor')
   Aadd( aHelpPor, 'final da faixa de Fornecedores')

                                                                                                           //F3
 //PutSx1(cPrgD ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPrgD ,"08","Ate Fornecedor"  ,"Ate Fornecedor","Ate Fornecedor","mv_ch8","C"  ,6       ,0       ,0      ,"G" ,""    ,"SA2" ,""     ,"","MV_PAR08","","","",""    ,""  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}
   
   Aadd( aHelpPor, 'Informe a Loja inicial do Fornecedor')
   Aadd( aHelpPor, 'da faixa de Lojas a serem listadas')

                                                                                                           //F3
 //PutSx1(cPrgD ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPrgD ,"09","Da Loja"  ,"Da Loja","Da Loja","mv_ch9","C"  ,2       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR09","","","",""    ,""  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)


   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}
   
   Aadd( aHelpPor, 'Informe a Loja final do Fornecedor')
   Aadd( aHelpPor, 'da faixa de Lojas a serem listadas')

                                                                                                           //F3
 //PutSx1(cPrgD ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPrgD ,"10","Ate Loja"  ,"Ate Loja","Ate Loja","mv_cha","C"  ,2       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR10","","","",""    ,""  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Data Inicial de Entrada')
   Aadd( aHelpPor, 'da Nota Fiscal na COMAFAL')

                                                                                                           //F3
 //PutSx1(cPrgD ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPrgD ,"11","Data Inicial"  ,"Data Inicial","Data Inicial","mv_chb","D"  ,08       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR11","","","",""    ,""  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Data Final de Entrada')
   Aadd( aHelpPor, 'da Nota Fiscal na COMAFAL')

                                                                                                           //F3
 //PutSx1(cPrgD ,"05","Depto.?" ,""       ,""     ,"mv_ch5","C"  ,1        ,0       ,0      ,"C" ,""    ,"" ,""     ,"","MV_PAR05","Todos","","",""    ,"Faturamento","","","Industria" ,"","","Financeiro"    ,"","","Gerencial","","")
   PutSx1(cPrgD ,"12","Data final"  ,"Data final","Data final","mv_chc","D"  ,08       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR12","","","",""    ,""  ,"","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   RestArea(aArea)

Return(.T.)