#include "Protheus.Ch"
#include "TopConn.Ch"
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³AJUSTOP   ºAutor  ³Microsiga           º Data ³  30/09/2010 º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Processo de Ajuste de consumo de insumos por OP            º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ COMAFAL- PE,SP e RS - Menu                                º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function AjustOP
   
   Private cProducao  := ALLTRIM(SC2->C2_NUM) + ALLTRIM(SC2->C2_ITEM) + ALLTRIM(SC2->C2_SEQUEN)  //Número da Ordem de Produção(O.P.)
   Private cDatExt := Alltrim(Str(Day(dDataBase)))+" de "+MesExtenso(MONTH(dDataBase))+" de "+Alltrim(Str(YEAR(dDataBase)))
   Private dOrdData   := SC2->C2_EMISSAO                                                          //Data da Emissão da O.P.
   Private dIniPrev   := SC2->C2_DATPRI                                                           //Data da Previsão de Inicio da O.P. 
   Private dPrevEnt   := SC2->C2_DATPRF                                                           //Data da Previsão de Entrega da O.P.
   Private nQtdOrd    := (SC2->C2_QUANT)                                                            //Quantidade da O.P.
   Private cArmazem   := SC2->C2_LOCAL                                                            //Armazém da O.P.
   Private cTipoOp    := IIF(SC2->C2_TPOP=="F","Firme","Prevista")                                //Tipo da O.P.
   Private cOpCCusto  := SC2->C2_CC                                                               //Centro de Custo da O.P.
   Private cCorteOrd  := SC2->C2_NUMOC                                                            //Ordem de Corte que gerou a O.P.
   Private cProdCod   := SC2->C2_PRODUTO                                                          //Código do Produto a produzir.
   Private cDescPrd   := Posicione("SB1",1,xFilial("SB1")+cProdCod,"B1_DESC")                     //Descrição do Produto a produzir.
   
   
   Private aLegenda := {}

   Private	aCorAjt  := {	{ "Empty(C2_DATAJT)"	, "BR_AMARELO" ,"Ajuste NÃO Realizado"},; 	
						    { "!Empty(C2_DATAJT)", "BR_VERMELHO","Ajuste JÁ Realizado" }} 

   Private	aLegenda := {	{ "BR_AMARELO" ,"Ajuste NÃO Realizado"},; 	
						    { "BR_VERMELHO","Ajuste JÁ Realizado" }} 

   Private aRotina    := {	{ OemToAnsi("Pesquisar"       ) ,"AxPesqui"		, 0 , 1},;
                     	    { OemToAnsi("Visualizar"      ) ,"AxVisual"		, 0 , 2},;
                     	    { OemToAnsi("Ajusta Consumo"  ) ,"u_AjtConsu()"	, 0 , 3},;
                     	    { OemToAnsi("Legenda"         ) ,"u_LegenAjt"	, 0 , 2} }

   Private cCadastro  := OemToAnsi("Ajusta Ordens de Produção")
   //Private xPerg := "PRGETQ"
   //MtPergEtq() //Monta Perguntas para definição do Tipo/Modelo de Etiquetas
   mBrowse(6,1,22,75,"SC2",,,,,,aCorAjt)
Return

User Function AjtConsu
   
   If !Empty(SC2->C2_DATRF)
      Alert("A O.P. "+Alltrim(cProducao)+" do Produto "+Alltrim(cDescPrd)+" já encontra-se ENCERRADA, favor selecionar outra Ordem de Produção")
      Return
   EndIf
   
   DEFINE FONT oBold NAME "Arial" SIZE 0, -13 BOLD
   DEFINE FONT oFtTmp NAME "Arial" SIZE 0, -11 BOLD
   DEFINE FONT oFntNegrito    NAME "Arial" SIZE 0, -11 BOLD
   DEFINE FONT oFnt2Negro    NAME "Arial" SIZE 10,14 BOLD
   DEFINE FONT oFntSMedia NAME "Arial" SIZE 16,18 
   DEFINE FONT oFntMedia NAME "Arial" SIZE 16,18 BOLD
   DEFINE FONT oFntGrande NAME "Arial" SIZE 20,22 BOLD
   //oFont1 := TFont():New("Arial",0,-9,,.F.,,,,.T.,.T.,.F.) //Fonte Ítalico Sublinhado
   oFont1 := TFont():New("Arial",0,16,,.F.,,,,.T.,.T.,.F.) //Fonte Ítalico Sublinhado
   cTitulo:= "Ajuste de Consumo Efetivo de OPs"
         
   Define MsDialog oDlg Title cTitulo From 001,005 To 630,980 Of oMainWnd Pixel

   @010,100 SAY "Ajuste de Ordem de Produção:" SIZE 200,8 FONT oFntGrande COLOR CLR_HBLUE Of oDlg PIXEL 
   @010,300 SAY cProducao SIZE 200,8 FONT oFntGrande Of oDlg PIXEL 

   @ 030,007 SAY "Informações da O.P.    " Of oDlg PIXEL SIZE 130 ,9 FONT oBold
   @ 036,007 SAY Replicate("_",800) Of oDlg PIXEL SIZE 450 ,9 FONT oBold
	
   @ 045,007 SAY cDatExt + ", "+Substr(Time(),1,5)+" hs" Of oDlg PIXEL SIZE 200,88 FONT oFtTmp
   
   @ 060,007 SAY "Produto :              " Of oDlg PIXEL SIZE 130 ,9 FONT oFntSMedia
   @ 060,077 SAY cProdCod+" - "+cDescPrd Of oDlg PIXEL SIZE 430 ,9 FONT oFntMedia
   
   @ 075,007 SAY Replicate("_",800) Of oDlg PIXEL SIZE 450 ,9 FONT oBold
   
   @ 090,007 SAY "Quantidade O.P.(Ton)   " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp
   @ 090,067 MSGET oQtdOrd VAR nQtdOrd   Picture "@E 999,999,999.999" Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp WHEN .F.

   @ 090,147 SAY "Emissão da O.P.        " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp
   @ 090,207 MSGET oOrdData VAR dOrdData  Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp WHEN .F.

   @ 090,297 SAY "Previsão de Inicio     " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp
   @ 090,357 MSGET oPrevIni VAR dIniPrev  Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp WHEN .F.

      
   @ 105,007 SAY "Armazém                " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp
   @ 105,067 MSGET oArmazem VAR cArmazem   Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp WHEN .F.

   @ 105,147 SAY "Tipo                   " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp
   @ 105,207 MSGET oTipoOp VAR cTipoOp   Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp WHEN .F.

   @ 105,297 SAY "Centro de Custo        " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp
   @ 105,357 MSGET oOpCCusto VAR cOpCCusto  Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp WHEN .F.
      
   @ 120,007 SAY "Quant. Já Produzida    " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp 
   @ 120,067 MSGET oJaProduzido VAR nJaProduzido Picture "@E 999,999,999.999" Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp WHEN .F.

   @ 120,147 SAY "Saldo a Produzir       " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp
   @ 120,207 MSGET oAProduzir VAR nAProduzir  Picture "@E 999,999,999.999" Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp WHEN .F.


   @ 120,297 SAY "Peso do Amarrado       " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp COLOR CLR_HBLUE
   @ 120,357 MSGET oPesoAmar VAR nPesoAmar  Picture "@E 999,999,999.999" Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp WHEN .F.
   
   @ 135,007 SAY "Quantidade de Peças    " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp COLOR CLR_HBLUE 
   @ 135,067 MSGET oQtdPecas    VAR nQtdPecas   Picture "@E 999,999,999" Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp WHEN .T. VALID(fVldQtPc())

/*
   @ 135,147 SAY "Turno                  " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp COLOR CLR_HBLUE
   @ 135,207 MSGET oTurno     VAR cTurno    Picture "@!" F3 "9A" Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp WHEN .T. VALID(ExistCpo("SX5","9A"+cTurno) .And. fVldTrno())
   
   @ 135,297 SAY "Previsão de Entrega    " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp
   @ 135,357 MSGET oPrevEnt VAR dPrevEnt  Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp WHEN .F.   

   @ 150,007 SAY "Ordem de Corte         " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp COLOR CLR_HRED
   @ 150,067 MSGET oCorteOrd VAR cCorteOrd  Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp WHEN .F.
   
   @ 150,147 SAY "Máquina                " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp COLOR CLR_HBLUE
   @ 150,207 MSGET oMaqProd  VAR cMaqProd  Picture "@!" Of oDlg PIXEL SIZE 070 ,8 F3 "SH1" WHEN .T. VALID(ExistCpo("SH1",cMaqProd) .And. fMqVld())
   
   
   @ 150,297 SAY "Número do Lote         " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp //COLOR CLR_HBLUE
   @ 150,357 MSGET oNumLte   VAR cNumLte   Picture "@!" Of oDlg PIXEL SIZE 070 ,8 F3 "XB8"  WHEN .T. VALID(IIF(!Empty(cNumLte),ExistLte(cNumLte),.T.)) //XB8 É uma Cópia adaptada da Consulta F3 do SB8 especificamente para este programa.
   
      
   @ 165,007 SAY "Produto Origem         " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp 
   @ 165,067 MSGET ocPrdOrig VAR _ycPrdOrig  Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp F3 "SB1" WHEN .T.
   
   @ 165,147 SAY "Armazem Origem         " Of oDlg PIXEL SIZE 070 ,8 FONT oFtTmp 
   @ 165,207 MSGET ocLocOrig  VAR _ycLocOrig  Picture "@!" Of oDlg PIXEL SIZE 070 ,8 WHEN .T. 

   
   @ 188,007 BUTTON OemToAnsi("&O u t r a s   E n t r a d a s [A p o n t a m e n t o s   d e   P e r d a s]") SIZE 450,12 OF oDlg PIXEL ACTION Processa({|| OutrasENT()},"Outras Entradas/Apontamento de Perdas ") WHEN .T.
    
   @ 195,007 SAY Replicate("_",800) Of oDlg PIXEL SIZE 450 ,9 FONT oBold
    
   @ 205,007 BUTTON OemToAnsi("&C a p t u r a r   P e s o   A m a r r a d o   /   I m p r i m i r   E t i q u e t a   /   G r a v a r   P r o d u ç ã o") SIZE 450,12 OF oDlg PIXEL ACTION Processa({|| CapPesoAM()},"Capturando Peso do Amarrado ") WHEN .T.
   

   @ 225,007 LISTBOX oProds VAR cVarProds Fields HEADER "Documento","Emissão","Quantidade","Tipo de Movimentação","Armazém","Centro de Custo","Conta Contábil","Unidade","Parcial/Total","Perda","Hora","Qtd.Pecas","Operador","Turno","Maquina","Sequencia" SIZE 450,60 NOSCROLL OF oDlg PIXEL //,"Estornado"
   @ 290,007 SAY "click duplo sobre a produção selecionada reimprime Etiqueta" FONT oFont1 COLOR CLR_HBLUE OF oDlg PIXEL
      
   oProds:SetArray(aProducoes)
   oProds:bLine := { || { aProducoes[oProds:nAt][1], aProducoes[oProds:nAt,2],;
   aProducoes[oProds:nAt,3],aProducoes[oProds:nAt,4],aProducoes[oProds:nAt,5],aProducoes[oProds:nAt,6],;
   aProducoes[oProds:nAt,7],aProducoes[oProds:nAt,8],aProducoes[oProds:nAt,9],aProducoes[oProds:nAt,10],; //}}
   aProducoes[oProds:nAt,11],aProducoes[oProds:nAt,12],aProducoes[oProds:nAt,13],aProducoes[oProds:nAt,14],;//}}
   aProducoes[oProds:nAt,15],aProducoes[oProds:nAt,16]}}
   
   oProds:blDblClick := { || Etq2Zebra( aProducoes[oProds:nAt][1], aProducoes[oProds:nAt,2],;
                                        aProducoes[oProds:nAt,3],aProducoes[oProds:nAt,4],aProducoes[oProds:nAt,5],aProducoes[oProds:nAt,6],;
                                        aProducoes[oProds:nAt,7],aProducoes[oProds:nAt,8],aProducoes[oProds:nAt,9],aProducoes[oProds:nAt,10],;
                                        aProducoes[oProds:nAt,11],aProducoes[oProds:nAt,12],aProducoes[oProds:nAt,13],aProducoes[oProds:nAt,14],;
                                        aProducoes[oProds:nAt,15],aProducoes[oProds:nAt,16] )}
   */
   Activate MsDialog oDlg Center //On Init EnchoiceBar(oDlg,{|| If(CAI002TudoOk(nOpc,aBobinas,aSliter), ( nOpca:=1, oDlg:End() ), .F.) },{||oDlg:End()},,If(lEd,"",aButtons))
Return

User Function LegenAjt
   BrwLegenda(cCadastro,"Legenda",aLegenda) 
Return(.T.)

Static Function ReqEmpOP
   //12345678901
   //00373401002  
   cQrySD4 := " SELECT "
   cQrySD4 += "  FROM "+RetSQLName("SD4")+" SD4"
   cQrySD4 += " WHERE SD4.D4_FILIAL = '"+xFilial("SD4")+"'"
   cQrySD4 += "   AND SD4.D4_OP = '"+cProducao+"'"
   cQrySD4 += "   AND SD4.D_E_L_E_T_ <> '*'"
Return