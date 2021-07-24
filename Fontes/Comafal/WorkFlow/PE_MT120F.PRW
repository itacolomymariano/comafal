#INCLUDE "PROTHEUS.CH"
#include "TopConn.Ch"
#INCLUDE "TBICONN.CH "
#include "TbiCode.ch"
#INCLUDE "Ap5mail.ch"
#include "wfprocess.ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³PE_MT120F ºAutor  ³Five Solutions      º Data ³  09/04/2010 º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ PE disparado após inclusão do Pedido de Compras.           º±±
±±º          ³ Usado para disparar Workflow de Aprovação do Pedido.       º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ COMAFAL - PE,SP e RS                                      º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

//User Function MT120GOK(nOpWF,oProcess) //,cEmpWF,cFilWF)
//Execblock("MT120GOK",.F.,.F.,{cA120Num,l120Inclui,l120Altera,l120Deleta})
User Function MT120F(nOpWF,oProcess)

   
   //ExecBlock("MT120F",.f.,.f.,xFilial("SC7")+ca120Num)
   
   Default nOpWF        := 1
   //Private oProcess
   Private lPrcMaior := .F. //Five Solutions Consultoria - 06/06/2011
   
   Do Case 
      Case nOpWF == 1
           u_fAprvPC()
      Case nOpWF == 2
           u_RetAprvPC() //,cEmpWF,cFilWF)
           //u_testwf()
      Case nOpWF == 3
           u_fTOPC(cNumPC) //oProcess) //oProcess,cNumPC,cEmpWF,cFilWF)
   EndCase

Return   
/*
User Function u_testwf()
   //RPCSetEnv("91","01","","","","",{"SC7"})
   ConOut("Estou em testwf Empresa "+cEmpAnt+" Filial "+cFilAnt)
Return
  */
User Function fAprvPC(oProcess)   

   Private cNumPC       := cA120Num //SC7->C7_NUM
   Private cCdTarefa    := ""

   cQryValor := " SELECT SUM(C7_TOTAL) NVLRPC"
   cQryValor += "   FROM "+RetSQLName("SC7")
   cQryValor += "  WHERE C7_FILIAL = '"+xFilial("SC7")+"'"
   cQryValor += "    AND C7_NUM = '"+cNumPC+"'"
   cQryValor += "    AND D_E_L_E_T_ <> '*'"
   
   TCQuery cQryValor NEW ALIAS "TVLR"
   TCSetField("TVLR","NVLRPC","N",17,02)
   DbSelectArea("TVLR")
   nPCValor := TVLR->NVLRPC
   DbCloseArea()
   
   /* - Five Solutions Consultoria 31/05/2010
   *    Comentado pois todos os pedidos de compras serão submetidos a aprovação.
   If nPCValor < 3000
      Return
   EndIf
   */

   cQryTipo := " SELECT SC7.C7_PRODUTO,SB1.B1_TIPO"
   cQryTipo += "   FROM "+RetSQLName("SC7")+" SC7,"+RetSQLName("SB1")+" SB1"
   cQryTipo += "  WHERE SC7.C7_FILIAL = '"+xFilial("SC7")+"'"
   cQryTipo += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"
   cQryTipo += "    AND SC7.C7_PRODUTO = SB1.B1_COD"
   cQryTipo += "    AND SC7.C7_NUM = '"+cNumPC+"'"
   cQryTipo += "    AND SC7.D_E_L_E_T_ <> '*'"
   cQryTipo += "    AND SB1.D_E_L_E_T_ <> '*'"
   
   TCQuery cQryTipo NEW ALIAS "TVTP"
   DbSelectArea("TVTP")
   cTipo := TVTP->B1_TIPO
   DbCloseArea()
   
   //Five Solutions - 06/05/2011 - If cTipo == "MP"
   If Alltrim(cTipo) $ "MP|PI|PA" //Five Solutions - 06/05/2011
      Return
   EndIf
   
   //Bloqueia o Pedido
   _cExec := " UPDATE "+RetSqlName('SC7')+" SET C7_CONAPRO = 'B' "
   _cExec += "  WHERE C7_FILIAL = '"+xFilial("SC7")+"'"
   _cExec += "    AND C7_NUM = '"+cNumPC+"'"
   _cExec += " AND   D_E_L_E_T_ = ' ' "
			
   TcSqlExec( _cExec )
			
   _cExec := " COMMIT "
			
   TcSqlExec( _cExec )
      
   
   
   Private cCodProcesso := "000001" //Codigo do Cadastro de Processos
   Private cTexto       := "Solicitando Aprovacao do PC: "+cNumPC  // Crie um texto que identifique as etapas do processo que foi realizado
                                                                   // para futuras consultas na janela de rastreabilidade.     
   Private cCodStatus   := "100100" // Informe o código de status correspondente a essa etapa.
                                    // Código do cadastro de status de processo.
   
   Private cAssunto     := "Aprovacao do Pedido de Compras "+cNumPC
   Private cTitulo      :=  "Criando Aprovacao"
   //Private cHtmlAC      := "C:\Protheus10\Workflow\ApPedCom.htm"
   Private cHtmlAC      := "\Workflow\ApPedCom.htm"
   Private cSrvBrw      := Alltrim(GETMV("MV_WFBRWSR"))

   Private cForPC       := Posicione("SC7",1,xFilial("SC7")+cA120Num,"C7_FORNECE")
   Private cLjPC        := Posicione("SC7",1,xFilial("SC7")+cA120Num,"C7_LOJA")
   Private pcUsuario    := "" //"itacolomy@fivesolutions.com.br"
   Private cUsrProt     := cUserName //SubStr(cUsuario,7,15)   // Registre o nome do usuário corrente que está criando o processo
   /*
   Private cTo          := "itacolomy@fivesolutions.com.br"
   Private cCC          := "itacolomy@hotmail.com"
   Private cBCC         := "suporte_siga@yahoo.com.br"
   */
   Private cMotApRp     := "Informe aqui o motivo da Aprovação ou Reprovação"
   Private cCdUsEmp     := Alltrim(WFCodUser(cUsrProt))

   //pcAtachar := "\Workflow\ApPedCom.htm" //""

   
   
   oProcess	:= TWFProcess():New( "000001", "Aprovação de Compras" )

   If SC7->(FieldPos("C7_WFID")) > 0
      DbSelectArea("SC7")
      nGoSC7 := SC7->(RecNo())
      DbSetOrder(1) //C7_FILIAL+C7_NUM+C7_ITEM+C7_SEQUEN
      If DbSeek(xFilial("SC7")+cNumPC)
         While SC7->C7_NUM == cNumPC .And. SC7->(!Eof())
            If RecLock("SC7",.F.)
               SC7->C7_WFID := oProcess:fProcessID
               MsUnLock()
            EndIf
            DbSelectArea("SC7")
            SC7->(DbSkip())
         EndDo
      EndIf
   EndIf
   
   //Processa({ ||GHtmWFPC(cNumPC,cNumPC,cForPC,cLjPC)},"Gerando Html do Pedido "+cNumPC+" Aguarde ...")
   //Processa({ ||NovoPC(cNumPC,cNumPC,cForPC,cLjPC)},"Gerando Html do Pedido "+cNumPC+" Aguarde ...")
   //Five Solutions - 06/06/2011
   
   Processa({ ||AprovPC(cNumPC,cNumPC,cForPC,cLjPC,oProcess)},"Gerando Html do Pedido "+cNumPC+" Aguarde ...")
   
   
   //Processa({||GrCorpAp(cNumPC)},"Gerando Html do Corpo do Mail, Aguarde ...")
   //cMailID := oProcess:oWF //oProcess:fProcessID + oProcess:FTaskID //cCdTarefa
   //Processa({||LinkApPc(cNumPC)},"Gerando Html do Link do Mail, Aguarde ...")
   
   oProcess:cTo    := cCdUsEmp
   oProcess:csubject := cAssunto 
   oProcess:NewVersion(.T.)
   oProcess:nEncodeMime := 0
   oProcess:bReturn :=  "u_RetAprvPC()"
   //oProcess:btimeout := {{"U_Rcomw01(6,'"+ SC7->C7_NUM +"')",_nDia, _nHora, 0,}} // Tempo de TimeOutS 
                                              //   D  H  M - D=Dia, H=Hora, M=Minuto
   oProcess:btimeout := {{"u_fTOPC(cNumPC)",0, 2, 0,}}//,cEmpAnt,cFilAnt))" // Tempo de TimeOutS 
   
   // Repasse as informações para o método responsável pela rastreabilidade.
   oProcess:Track(cCodStatus, cTexto, cCdUsEmp) // Rastreabilidade

   //RastreiaWF(oProcess :fProcessID + '.' + oProcess:fTaskID,oProcess:fProcCode, '100100',"Solicitando Aprovacao do PC: "+cNumPC, cUserName)
   
   cMailID := oProcess:Start()
   
   RastreiaWF(oProcess :fProcessID + '.' + oProcess:fTaskID,oProcess:fProcCode, '100100',"Solicitando Aprovacao do PC: "+cNumPC, cUserName) 
   oProcess:Free() 
   oProcess:= Nil

   cWFHtmOrig := '\workflow\messenger\emp' + cEmpAnt + '\'+cCdUsEmp+'\' + cMailId + '.htm'
   cWFHtmDest := '\workflow\procs\' + cMailId + '.htm'
   COPY FILE &(cWFHtmOrig) TO &(cWFHtmDest)

   cWFValOrig := '\workflow\emp' + cEmpAnt + '\process\' + cMailId + '.val'
   cWFValDest := '\workflow\procs\' + cMailId + '.val'
   COPY FILE &(cWFValOrig) TO &(cWFValDest)


   pcAtachar := "" //'\workflow\messenger\emp' + cEmpAnt + '\'+cCdUsEmp+'\' + cMailId + '.htm'
   
   oProcess	:= TWFProcess():New( "000001", "Aprovação de Compras - Enviando Email com Link" ) 
   oProcess:NewVersion(.T.)
   oProcess:nEncodeMime := 0
   Processa({||LinkApPc(cNumPC)},"Gerando Html do Link do Mail, Aguarde ...")    

   
   oProcess:newtask('Link', '\workflow\WFLinkApPc.htm')  //Inicio uma nova Task com um HTML Simples

   /**********************
   * Five Solutions COnsultoria
   * ALÇADAS ENVIADA POR PATRÍCIA - 31/05/2010
   *

 	COMPRADOR	INDÚSTRIA	FINANCEIRO	DIRETOR
    COPE	R$0,00 à R$1.000,00 Miqueias Santino msantino@comafal.com.br
            R$1.000,01 à R$2.000,00 Flávio Trajano ftrajano@comafal.com.br
R$1.000,01 à R$10.000,00 Roberto Lira rlira@comafal.com.br
R$10.000,01 à R$........... Gilmar Filho gfilho@comafal.com.br

COSP	R$0,00 à R$1.000,00  Luciano Souza  lsouza@comafal.com.br
R$1.000,01 à R$2.000,00 Kleber Rodrigues krodrigues@comafal.com.br
R$1.000,01 à R$10.000,00 Roberto Lira rlira@comafal.com.br
R$10.000,01 à R$........... Gilmar Filho gfilho@comafal.com.br

CORS	R$0,00 à R$1.000,00  Cristiano Silva  cfsilva@comafal.com.br
R$1.000,01 à R$2.000,00  Elvio Almeida falmeida@comafal.com.br
R$1.000,01 à R$10.000,00 Roberto Lira rlira@comafal.com.br
R$10.000,01 à R$........... Gilmar Filho gfilho@comafal.com.br
     ********************************************************************************/
   //oProcess:cTo    := "itacolomy@fivesolutions.com.br"
   
   //If nPCValor <= 1000  //E-mail Enviados para os respectivos compradores de cada unidade COMAFAL.
   If nPCValor <= GetMv("FS_VAL1NV")  //Five - 16/12/2010
   
      If SM0->M0_CODFIL == "02" //COSP
         oProcess:cTo    := Alltrim(GetMV("FS_APN1SP")) //Five - 16/12/2010 - Antes, "lsouza@comafal.com.br"
      ElseIf SM0->M0_CODFIL == "03" //CORS
         oProcess:cTo    := Alltrim(GetMV("FS_APN1RS")) //Five - 16/12/2010 - Antes, "cfsilva@comafal.com.br"
      ElseIf SM0->M0_CODFIL == "06" //COPE
         oProcess:cTo    := Alltrim(GetMV("FS_APN1PE")) //Five - 16/12/2010 - Antes, "mandre@comafal.com.br" //"msantino@comafal.com.br"
      EndIf
   
   //ElseIf nPCValor >= 1001 .And. nPCValor <= 2000 //E-mail enviado para os respectivos gerentes indústrial de cada unidade e cópia para Roberto Lira(Gerente Financeiro)
   ElseIf nPCValor >= GetMv("FS_VLP2NV") .And. nPCValor <= GetMv("FS_VAL2NV") //Five - 16/12/2010
      
      If SM0->M0_CODFIL == "02"//COSP
         oProcess:cTo    := Alltrim(GetMV("FS_APN2SP")) //Five - 16/12/2010 - Antes, "krodrigues@comafal.com.br"
      ElseIf SM0->M0_CODFIL == "03"//CORS
         oProcess:cTo    := Alltrim(GetMV("FS_APN2RS")) //Five - 16/12/2010 - Antes, "falmeida@comafal.com.br"
      ElseIf SM0->M0_CODFIL == "06"//COPE
         oProcess:cTo    := Alltrim(GetMV("FS_APN2PE")) //Five - 16/12/2010 - Antes, "ftrajano@comafal.com.br"
      EndIf
      
      //oProcess:cCC  := "rlira@comafal.com.br" //Da faixa R$ 1.001,00 a R$ 2.000,00 serão enviados com cópia p/Roberto Lira(Gerente Financeiro)
      oProcess:cCC  := Alltrim(GetMV("FS_CCN2PC")) //Five - 16/12/2010
      
   //ElseIf nPCValor >= 2001 .And. nPCValor <= 10000 //Da faixa R$ 2.001,00 a R$ 10.000,00 serão enviados p/Roberto Lira(Gerente Financeiro)
   ElseIf nPCValor >= GetMv("FS_VLP3NV") .And. nPCValor <= GetMv("FS_VAL3NV") //Five - 16/12/2010
      oProcess:cTo    := Alltrim(GetMv("FS_APR3NV")) //Five - 16/12/2010 - Antes, "rlira@comafal.com.br"
      //If nPCValor >= 5000
      If nPCValor >= GetMv("FS_N3VMCC")  //Five - 16/12/2010
         oProcess:cCC  := Alltrim(GetMV("FS_CCN3PC"))//Five - 16/12/2010 - Antes, "gfilho@comafal.com.br"
      EndIf
   //ElseIf nPCValor > 10000
   ElseIf nPCValor > GetMv("FS_VLP4NV") //Five - 16/12/2010
      oProcess:cTo    := Alltrim(GetMv("FS_APR4NV")) //Five - 16/12/2010 - Antes, "gfilho@comafal.com.br"
   EndIf
   
   
   //oProcess:cCC  := NIL
   //oProcess:cBCC := NIL
	
   oProcess:csubject := cAssunto 
   //oProcess:NewVersion(.T.)
   //oProcess:AttachFile(Upper(pcAtachar))
   
   //oHtml   := oProcess:oHTML
   
   //oProcess:ohtml:valbyname('PCNUM', cNumPC )
   //oProcess:ohtml:ValByName('proc_link','http://ITACOLOMY-ACER/messenger/emp' + cEmpAnt + '/'+cCdUsEmp+'/' + cMailId + '.htm' ) //Defino o Link onde foi gravado o HTML pelo Workflow,abaixo do diretório do usuário definido em cTo do processo acima.
   //oProcess:ohtml:ValByName('proc_link','http://ITACOLOMY-ACER/procs/' + cMailId + '.htm' ) //Defino o Link onde foi gravado o HTML pelo Workflow,abaixo do diretório do usuário definido em cTo do processo acima.
   //Five 20/01/2011 - Comentado devido mudança de dominio - oProcess:ohtml:ValByName('proc_link','http://comafal.microsiga.com.br:8092/procs/' + cMailId + '.htm' ) //Defino o Link onde foi gravado o HTML pelo Workflow,abaixo do diretório do usuário definido em cTo do processo acima.
   oProcess:ohtml:ValByName('proc_link','http://comafal.totvs.com.br:8092/procs/' + cMailId + '.htm' ) //Defino o Link onde foi gravado o HTML pelo Workflow,abaixo do diretório do usuário definido em cTo do processo acima.
   
   oProcess:ohtml:ValByName('referente', 'Pedido de Compras Numero: '+cNumPC )
   oProcess:cBody := '\workflow\WFLinkApPc.htm' //"\workflow\CorpoApCp.htm"
   //oProcess:nEncodeMime := 0
   oProcess:Start()
   ConOut("O Html gerado foi cMailID: "+cMailID)

   cTexto := "Aguardando Retorno da Aprovacao do PC: "+cNumPC
   cCodStatus := "100200"
   oProcess:Track(cCodStatus, cTexto , cCdUsEmp) // Rastreabilidade
   RastreiaWF(oProcess :fProcessID + '.' + oProcess:fTaskID,oProcess:fProcCode, '100200',"Aguardando Retorno da Aprovacao do PC: "+cNumPC, cUserName)

Return

//Função TIME-OUT
User Function fTOPC(cNumPC,oProcess) //fTOPC(oProcess,cNumPC,xEmpWF,xFilWF)

   Private cCdTarefa    := ""
   Private oProcess

      
   Private cCodProcesso := "000001" //"PEDCOM"
   Private cAssunto     := "REENVIO - Aprovacao do Pedido de Compras "+cNumPC
   Private cTitulo      :=  "TimeOut do Pedido de Compras: "+cNumPC
   Private cHtmlAC      := "\Workflow\ApPedCom.htm"                         


   Private cForPC       := Posicione("SC7",1,xFilial("SC7")+cNumPC,"C7_FORNECE")
   Private cLjPC        := Posicione("SC7",1,xFilial("SC7")+cNumPC,"C7_LOJA")
   Private cUsrProt     := cUserName //SubStr(cUsuario,7,15)   // Registre o nome do usuário corrente que está criando o processo
   Private cTo          := "itacolomy@fivesolutions.com.br"
   Private cCC          := "itacolomy@hotmail.com"
   Private cBCC         := "suporte_siga@yahoo.com.br"
   Private cMotApRp     := "Informe aqui o motivo da Aprovação ou Reprovação"
   Private cCdUsEmp     := "000000" //Administador(Fixo p/Time-Out)//Alltrim(WFCodUser(cUsrProt)) //WFCodUser("BI")


   //pcAtachar := "\Workflow\ApPedCom.htm" //""

   cQryValor := " SELECT SUM(C7_TOTAL) NVLRPC"
   cQryValor += "   FROM "+RetSQLName("SC7")
   cQryValor += "  WHERE C7_FILIAL = '"+xFilial("SC7")+"'"
   cQryValor += "    AND C7_NUM = '"+cNumPC+"'"
   cQryValor += "    AND D_E_L_E_T_ <> '*'"
   
   TCQuery cQryValor NEW ALIAS "TVLR"
   TCSetField("TVLR","NVLRPC","N",17,02)
   DbSelectArea("TVLR")
   nPCValor := TVLR->NVLRPC
   DbCloseArea()

   cQryTipo := " SELECT SC7.C7_PRODUTO,SB1.B1_TIPO"
   cQryTipo += "   FROM "+RetSQLName("SC7")+" SC7,"+RetSQLName("SB1")+" SB1"
   cQryTipo += "  WHERE SC7.C7_FILIAL = '"+xFilial("SC7")+"'"
   cQryTipo += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"
   cQryTipo += "    AND SC7.C7_PRODUTO = SB1.B1_COD"
   cQryTipo += "    AND SC7.C7_NUM = '"+cNumPC+"'"
   cQryTipo += "    AND SC7.D_E_L_E_T_ <> '*'"
   cQryTipo += "    AND SB1.D_E_L_E_T_ <> '*'"
   
   TCQuery cQryTipo NEW ALIAS "TVTP"
   DbSelectArea("TVTP")
   cTipo := TVTP->B1_TIPO
   DbCloseArea()
   
   //Five Solutions - 06/05/2011 - If cTipo == "MP"
   If Alltrim(cTipo) $ "MP|PI|PA" //Five Solutions - 06/05/2011
      Return
   EndIf

   oProcess          	:= TWFProcess():New( "000001", "Time Out - Aprovação de Compras" )

   If SC7->(FieldPos("C7_WFID")) > 0
      DbSelectArea("SC7")
      nGoSC7 := SC7->(RecNo())
      DbSetOrder(1) //C7_FILIAL+C7_NUM+C7_ITEM+C7_SEQUEN
      If DbSeek(xFilial("SC7")+cNumPC)
         While SC7->C7_NUM == cNumPC .And. SC7->(!Eof())
            If RecLock("SC7",.F.)
               SC7->C7_WFID := oProcess:fProcessID
               MsUnLock()
            EndIf
            DbSelectArea("SC7")
            SC7->(DbSkip())
         EndDo
      EndIf
   EndIf
   
   //Processa({ ||GHtmWFPC(cNumPC,cNumPC,cForPC,cLjPC)},"Gerando Html do Pedido "+cNumPC+" Aguarde ...")
   //Processa({ ||NovoPC(cNumPC,cNumPC,cForPC,cLjPC)},"Gerando Html do Pedido "+cNumPC+" Aguarde ...")

   Processa({ ||AprovPC(cNumPC,cNumPC,cForPC,cLjPC,oProcess)},"Gerando Html do Pedido "+cNumPC+" Aguarde ...")
   
   
   //Processa({||GrCorpAp(cNumPC)},"Gerando Html do Corpo do Mail, Aguarde ...")
   //cMailID := oProcess:oWF //oProcess:fProcessID + oProcess:FTaskID //cCdTarefa
   //Processa({||LinkApPc(cNumPC)},"Gerando Html do Link do Mail, Aguarde ...")
   
   oProcess:cTo    := cCdUsEmp
   oProcess:csubject := cAssunto 
   oProcess:NewVersion(.T.)
   oProcess:nEncodeMime := 0
   oProcess:bReturn :=  "u_RetAprvPC()" //"u_MT120GOK(2,oProcess)" 
   //oProcess:btimeout := {{"U_Rcomw01(6,'"+ SC7->C7_NUM +"')",_nDia, _nHora, 0,}} // Tempo de TimeOutS 
                                              //   D  H  M - D=Dia, H=Hora, M=Minuto
   
   oProcess:btimeout := {{"u_fTOPC(cNumPC)",0, 2, 0,}} //{{"u_MT120GOK(3,oProcess)",0, 2, 0,}}  // Tempo de TimeOutS 
   
   cMailID := oProcess:Start()
   
   
   cWFHtmOrig := '\workflow\messenger\emp' + cEmpAnt + '\'+cCdUsEmp+'\' + cMailId + '.htm'
   cWFHtmDest := '\workflow\procs\' + cMailId + '.htm'
   COPY FILE &(cWFHtmOrig) TO &(cWFHtmDest)

   cWFValOrig := '\workflow\emp' + cEmpAnt + '\process\' + cMailId + '.val'
   cWFValDest := '\workflow\procs\' + cMailId + '.val'
   COPY FILE &(cWFValOrig) TO &(cWFValDest)


   pcAtachar := "" //'\workflow\messenger\emp' + cEmpAnt + '\'+cCdUsEmp+'\' + cMailId + '.htm'
   
   Processa({||LinkApPc(cNumPC)},"Gerando Html do Link do Mail, Aguarde ...")    
   oProcess:newtask('Link', '\workflow\WFLinkApPc.htm')  //Inicio uma nova Task com um HTML Simples
   //oProcess:cTo    := "itacolomy@fivesolutions.com.br" //;cpdteste@comafal.com.br;hsilva@comafal.com.br" //;itacolomy@hotmail.com;cecilia@fivesolutions.com.br;suporte_siga@yahoo.com.br"
   If nPCValor <= GetMv("FS_VAL1NV")  //Five - 16/12/2010 - E-mail Enviados para os respectivos compradores de cada unidade COMAFAL.
   
      If SM0->M0_CODFIL == "02" //COSP
         oProcess:cTo    := Alltrim(GetMV("FS_APN1SP")) //Five - 16/12/2010 - Antes, "lsouza@comafal.com.br"
      ElseIf SM0->M0_CODFIL == "03" //CORS
         oProcess:cTo    := Alltrim(GetMV("FS_APN1RS")) //Five - 16/12/2010 - Antes, "cfsilva@comafal.com.br"
      ElseIf SM0->M0_CODFIL == "06" //COPE
         oProcess:cTo    := Alltrim(GetMV("FS_APN1PE")) //Five - 16/12/2010 - Antes, "msantino@comafal.com.br"
      EndIf
   
   ElseIf nPCValor >= GetMv("FS_VLP2NV") .And. nPCValor <= GetMv("FS_VAL2NV") //E-mail enviado para os respectivos gerentes indústrial de cada unidade e cópia para Roberto Lira(Gerente Financeiro)
      
      If SM0->M0_CODFIL == "02"//COSP
         oProcess:cTo    := Alltrim(GetMV("FS_APN2SP")) //Five - 16/12/2010 - Antes, "krodrigues@comafal.com.br"
      ElseIf SM0->M0_CODFIL == "03"//CORS
         oProcess:cTo    := Alltrim(GetMV("FS_APN2RS")) //Five - 16/12/2010 - Antes, "falmeida@comafal.com.br"
      ElseIf SM0->M0_CODFIL == "06"//COPE
         oProcess:cTo    := Alltrim(GetMV("FS_APN2PE")) //Five - 16/12/2010 - Antes, "ftrajano@comafal.com.br"
      EndIf
      
      oProcess:cCC  := Alltrim(GetMV("FS_CCN2PC")) //Five 16/12/2010 Antes - "rlira@comafal.com.br" //Da faixa R$ 1.001,00 a R$ 2.000,00 serão enviados com cópia p/Roberto Lira(Gerente Financeiro)
      
   ElseIf nPCValor >= GetMv("FS_VLP3NV") .And. nPCValor <= GetMv("FS_VAL3NV") //Da faixa R$ 2.001,00 a R$ 10.000,00 serão enviados p/Roberto Lira(Gerente Financeiro)
      oProcess:cTo    := Alltrim(GetMv("FS_APR3NV"))
      If nPCValor >= GetMv("FS_N3VMCC")
         oProcess:cCC  := Alltrim(GetMV("FS_CCN3PC"))
      EndIf
   ElseIf nPCValor > GetMv("FS_VLP4NV")
      oProcess:cTo    := Alltrim(GetMv("FS_APR4NV"))
   EndIf

   //oProcess:cCC  := NIL
   //oProcess:cBCC := NIL
	
   oProcess:csubject := cAssunto 
   //oProcess:NewVersion(.T.)
   //oProcess:AttachFile(Upper(pcAtachar))
   
   //oHtml   := oProcess:oHTML
   
   //oProcess:ohtml:valbyname('PCNUM', cNumPC )
   //oProcess:ohtml:ValByName('proc_link','http://ITACOLOMY-ACER/messenger/emp' + cEmpAnt + '/'+cCdUsEmp+'/' + cMailId + '.htm' ) //Defino o Link onde foi gravado o HTML pelo Workflow,abaixo do diretório do usuário definido em cTo do processo acima.
   //oProcess:ohtml:ValByName('proc_link','http://ITACOLOMY-ACER/procs/' + cMailId + '.htm' ) //Defino o Link onde foi gravado o HTML pelo Workflow,abaixo do diretório do usuário definido em cTo do processo acima.
   //Five 20/01/2011 - Comentado devido mudança de dominio - oProcess:ohtml:ValByName('proc_link','http://comafal.microsiga.com.br:8092/procs/' + cMailId + '.htm' ) //Defino o Link onde foi gravado o HTML pelo Workflow,abaixo do diretório do usuário definido em cTo do processo acima.
   oProcess:ohtml:ValByName('proc_link','http://comafal.totvs.com.br:8092/procs/' + cMailId + '.htm' ) //Defino o Link onde foi gravado o HTML pelo Workflow,abaixo do diretório do usuário definido em cTo do processo acima.
   
   oProcess:ohtml:ValByName('referente', 'Pedido de Compras Numero: '+cNumPC )
   oProcess:cBody := '\workflow\WFLinkApPc.htm' //"\workflow\CorpoApCp.htm"
   //oProcess:nEncodeMime := 0
   cMailID := oProcess:Start()
   ConOut("O Html gerado foi cMailID: "+cMailID)

   cTexto := "Pedido de Compras "+cNumPC+" Reenviado para Aprovacao "
   cCodStatus := "100600"
   oProcess:Track(cCodStatus, cTexto , cCdUsEmp) // Rastreabilidade
   RastreiaWF(oProcess :fProcessID + '.' + oProcess:fTaskID,oProcess:fProcCode, '100600',"Pedido de Compras "+cNumPC+" Reenviado para Aprovacao ", cUserName)
   
   
Return

//Gera html
Static Function GHtmWFPC(cNumPC,cNumPC,cForPC,cLjPC)

   For nPC := 1 To 2
   
      If nPC == 1
         cQryPC  := " SELECT COUNT(*) NQTDREG"
      Else   
         cQryPC  := " SELECT SC7.C7_NUM,SC7.C7_EMISSAO,SC7.C7_FORNECE,SC7.C7_LOJA,SA2.A2_NOME,SC7.C7_COND,"
         cQryPC  += "        SC7.C7_ITEM,SC7.C7_PRODUTO,SB1.B1_DESC,SC7.C7_QUANT,SC7.C7_UM,SC7.C7_PRECO,SC7.C7_TOTAL,"
         cQryPC  += "        SC7.C7_DATPRF,SA2.A2_COD,SA2.A2_LOJA,SA2.A2_INSCR,SA2.A2_CGC,SA2.A2_CONTATO,SA2.A2_TEL,"
         cQryPC  += "        SA2.A2_FAX,SA2.A2_END,SA2.A2_BAIRRO,SA2.A2_MUN,SA2.A2_CEP,SA2.A2_CGC,SC7.C7_FRETE,"
         cQryPC  += "        SC7.C7_DESCRI,SC7.C7_VALIPI,SC7.C7_CC"
         
      EndIf
   
      cQryPC  += "   FROM "+RetSQLName("SC7")+" SC7,"+RetSQLName("SA2")+" SA2,"+RetSQLName("SB1")+" SB1"
      cQryPC  += "  WHERE SC7.C7_FILIAL = '"+xFilial("SC7")+"'"
      cQryPC  += "    AND SA2.A2_FILIAL = '"+xFilial("SA2")+"'"
      cQryPC  += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"
   
      cQryPC  += "    AND SC7.C7_FORNECE = SA2.A2_COD"
      cQryPC  += "    AND SC7.C7_LOJA = SA2.A2_LOJA"
      cQryPC  += "    AND SC7.C7_PRODUTO = SB1.B1_COD"
   
      cQryPC  += "    AND SC7.C7_NUM = '"+cNumPC+"'"
      cQryPC  += "    AND SC7.C7_FORNECE = '"+cForPC+"'"
      cQryPC  += "    AND SC7.C7_LOJA = '"+cLjPC+"'"
   
      cQryPC  += "    AND SC7.D_E_L_E_T_ <> '*'"
      cQryPC  += "    AND SA2.D_E_L_E_T_ <> '*'"
      cQryPC  += "    AND SB1.D_E_L_E_T_ <> '*'"
      
      If nPC > 1
         cQryPC  += " ORDER BY SC7.C7_ITEM"
      EndIf
   
      MemoWrite("PedidodeCompras.SQL",cQryPC)
   
      TCQuery cQryPC NEW ALIAS "TXC7"
      
      If nPC == 1
         DbSelectArea("TXC7")
         nReg := TXC7->NQTDREG
         TXC7->(DbCloseArea())
      EndIf
   
   Next nPC
   
   TCSetField("TXC7","C7_QUANT","N",12,03)
   TCSetField("TXC7","C7_PRECO","N",14,02)
   TCSetField("TXC7","C7_TOTAL","N",14,02)
   TCSetField("TXC7","C7_FRETE","N",14,02)
   TCSetField("TXC7","C7_DATPRF","D",08,00)
   TCSetField("TXC7","C7_EMISSAO","D",08,00)
   
   DbSelectArea("TXC7")
   ProcRegua(nReg)
   nCont := 1
   
   cHtmlWF := '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
   cHtmlWF += '<html xmlns="http://www.w3.org/1999/xhtml">'
   cHtmlWF += '<head>'
   cHtmlWF += '<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />'
   cHtmlWF += '<title>Aprova&ccedil;&atilde;o Pedidos de Compras - COMAFAL</title><style type="text/css">'
    
   cHtmlWF += '<!--'
   cHtmlWF += 'body,td,th {'
   cHtmlWF += '	font-size: 10px;'
   cHtmlWF += '}'
   cHtmlWF += '.style1 {'
   cHtmlWF += '	font-family: "Times New Roman", Times, serif;'
   cHtmlWF += '	font-weight: bold;'
   cHtmlWF += '}'
   cHtmlWF += '.style5 {font-size: 14px}'
   cHtmlWF += '-->'
   cHtmlWF += '</style></head>'
   cHtmlWF += ''
   cHtmlWF += '<body>'

   //Five 04/05/2010 - cHtmlWF += '<form action=”mailto:%WFMailTo%” method=”post” name="FormApCp">' //Five 03/05/2010

   cHtmlWF += '<div align="center" class="style1"><span class="style5">Aprova&ccedil;&atilde;o do Pedido de Compras Numero %PCNUM%</span></div>'
   //cHtmlWF += '<div align="center" class="style1"><span class="style5">Aprova&ccedil;&atilde;o do Pedido de Compras Numero '+cNumPC+'</span></div>'
   //cHtmlWF += '<div align="center" class="style1"><span class="style5">Aprova&ccedil;&atilde;o do Pedido de Compras Numero '+cNumPC+'</span></div>'
   cHtmlWF += '<br />'
   cHtmlWF += 'Cabe&ccedil;alho do Pedido<br /> '
   cHtmlWF += '<br />'
   cHtmlWF += '<table width="100%" border="1">'
   cHtmlWF += '  <tr bgcolor="#CCCCCC">'
   cHtmlWF += '    <th width="10%" scope="col">Emiss&atilde;o</th>'
   cHtmlWF += '    <th width="11%" scope="col">Fornecedor</th>'
   cHtmlWF += '    <th width="55%" scope="col">Nome</th>'
   cHtmlWF += '    <th width="11%" scope="col">Cond.Pgto</th>'
   cHtmlWF += '    <th width="13%" scope="col">N&uacute;mero</th>'
   cHtmlWF += '  </tr>'
   cHtmlWF += '  <tr>'
   cHtmlWF += '    <td>'+DTOC(TXC7->C7_EMISSAO)+'</td>'                    //cHtmlWF += '    <td>!emissao!</td>'
   cHtmlWF += '    <td>'+TXC7->C7_FORNECE+'/'+TXC7->C7_LOJA+'</td>'        //cHtmlWF += '    <td>!fornecedor!</td>'
   cHtmlWF += '    <td>'+Alltrim(TXC7->A2_NOME)+'</td>'                    //cHtmlWF += '    <td>!lb_nome!</td>' 
   cHtmlWF += '    <td>'+TXC7->C7_COND+'</td>'                             //cHtmlWF += '    <td>!lb_cond!</td>'     
   cHtmlWF += '    <td>'+TXC7->C7_NUM+'</td>'                              //cHtmlWF += '    <td>!pedido!</td>' 
   cHtmlWF += '  </tr>'
   cHtmlWF += '</table>'
   cHtmlWF += '<br />'
   cHtmlWF += '<br />'
   cHtmlWF += 'Itens do Pedido<br />'
   cHtmlWF += '<br /> '
   cHtmlWF += '<table width="100%" border="1">'
   cHtmlWF += '  <tr bgcolor="#CCCCCC">'
   cHtmlWF += '    <th width="9%" scope="col">Item</th>'
   cHtmlWF += '    <th width="38%" scope="col">Descri&ccedil;&atilde;o</th>'
   cHtmlWF += '    <th width="11%" scope="col">Quantidade</th>'
   cHtmlWF += '    <th width="9%" scope="col">Unidade</th>'
   cHtmlWF += '    <th width="10%" scope="col">Pre&ccedil;o</th>'
   cHtmlWF += '    <th width="11%" scope="col">Total</th>'
   cHtmlWF += '    <th width="12%" scope="col">Dt.Entrega</th>'
   cHtmlWF += '    <th width="12%" scope="col">Ultimo Preco</th>' //Five Solutions - 06/06/2011
   cHtmlWF += '  </tr>'
   
   DbSelectArea("TXC7")
   nTtPdComp := 0
   While TXC7->(!Eof())
      
      IncProc("Gerando Itens p/Aprovação ... "+Alltrim(Str(nCont))+" / "+Alltrim(Str(nReg)))
      /*
      cHtmlWF += '  <tr>'
      cHtmlWF += '    <td>'+TXC7->C7_ITEM+'</td>'                                        //cHtmlWF += '    <td>!produto.item!</td>'
      cHtmlWF += '    <td>'+TXC7->C7_PRODUTO+'-'+Alltrim(TXC7->B1_DESC)+'</td>'          //cHtmlWF += '    <td>!produto.codigo!-!produto.descricao!</td>'
      cHtmlWF += '    <td><input type=hidden name="produto.quant" value="'+Transform(TXC7->C7_QUANT,"@E 999,999,999.999")+'"></td>' //cHtmlWF += '    <td>%produto.quant%</td>'
      cHtmlWF += '    <td>'+TXC7->C7_UM+'</td>'                                          //cHtmlWF += '    <td>!produto.unid!</td>'
      cHtmlWF += '    <td><input type=hidden name="produto.preco" value="'+Transform(TXC7->C7_PRECO," 999,999,999.99")+'"></td>'    //cHtmlWF += '    <td>%produto.preco%</td>'
      cHtmlWF += '    <td><input type=hidden name="produto.total" value="'+Transform(TXC7->C7_TOTAL," 999,999,999.99")+'"></td>'    //cHtmlWF += '    <td>%produto.total%</td>'
      cHtmlWF += '    <td><input type=hidden name="produto.entrega" value="'+DTOC(TXC7->C7_DATPRF)+'"></td>'                        //cHtmlWF += '    <td>%produto.entrega%</td>'
      cHtmlWF += '  </tr>'
      */
      //Five Solutions - 06/06/2011 - Implementado para apresentar último preço de compra do Produto.
      nUPrcComp := 0
      DbSelectArea("SBZ") //Indicadores de Produtos
      DbSetOrder(1) //BZ_FILIAL+BZ_COD
      If DbSeek(xFilial("SBZ")+TXC7->C7_PRODUTO)
         nUPrcComp := SBZ->BZ_UPRC
         If nUPrcComp == 0
            DbSelectArea("SB1") //Cadastro Genérico de Produtos
            DbSetOrder(1) //B1_FILIAL+B1_COD
            If DbSeek(xFilial("SB1")+TXC7->C7_PRODUTO)
               nUPrcComp := SB1->B1_UPRC
            EndIf
         EndIf
      EndIf

	  If TXC7->C7_PRECO > nUPrcComp .And. !lPrcMaior
	     lPrcMaior := .T.
	  EndIf
	        
      DbSelectArea("TXC7")
      cHtmlWF += '  <tr>'
      cHtmlWF += '    <td>'+TXC7->C7_ITEM+'</td>'                                        //cHtmlWF += '    <td>!produto.item!</td>'
      cHtmlWF += '    <td>'+TXC7->C7_PRODUTO+'-'+Alltrim(TXC7->B1_DESC)+'</td>'          //cHtmlWF += '    <td>!produto.codigo!-!produto.descricao!</td>'
      cHtmlWF += '    <td>'+Transform(TXC7->C7_QUANT,"@E 999,999,999.999")+'</td>' //cHtmlWF += '    <td>%produto.quant%</td>'
      cHtmlWF += '    <td>'+TXC7->C7_UM+'</td>'                                          //cHtmlWF += '    <td>!produto.unid!</td>'
      cHtmlWF += '    <td>'+Transform(TXC7->C7_PRECO,"@E 999,999,999.99")+'</td>'    //cHtmlWF += '    <td>%produto.preco%</td>'
      cHtmlWF += '    <td>'+Transform(TXC7->C7_TOTAL,"@E 999,999,999.99")+'</td>'    //cHtmlWF += '    <td>%produto.total%</td>'
      cHtmlWF += '    <td>'+DTOC(TXC7->C7_DATPRF)+'</td>'                        //cHtmlWF += '    <td>%produto.entrega%</td>'
      cHtmlWF += '    <td>'+Transform(nUPrcComp,"@E 999,999,999.99")+'</td>'    //Último Preço de Compra do Produto
      cHtmlWF += '  </tr>'
      
      nTtPdComp += TXC7->C7_TOTAL //Soma itens para totalizar Pedido de Compras
       
      DbSelectArea("TXC7")
      //TXC7->(DbSkip())
      DbSkip()
      nCont ++
      
   EndDo
   DbSelectArea("TXC7")
   TXC7->(DbCloseArea())
   cHtmlWF += '</table>'
   cHtmlWF += '<br />'
   cHtmlWF += '<br />'
   cHtmlWF += '<table width="30%" border="1">'
   cHtmlWF += '  <tr>'
   cHtmlWF += '    <th width="38%" bgcolor="#CCCCCC" scope="col">Total do Pedido </th>'
   //cHtmlWF += '    <th width="62%" scope="col"><input type=hidden name="lbvalor" value="'+Transform(nTtPdComp,"@E 999,999,999.99")+'"></th>' //cHtmlWF += '    <th width="62%" scope="col">%lbvalor%</th>'
   cHtmlWF += '    <th width="62%" scope="col">'+Transform(nTtPdComp,"@E 999,999,999.99")+'</th>' //cHtmlWF += '    <th width="62%" scope="col">%lbvalor%</th>'
   cHtmlWF += '  </tr>'
   cHtmlWF += '</table>'
   cHtmlWF += '<br />'
   cHtmlWF += '<br />'
   cHtmlWF += '<span class="style1">APROVA&Ccedil;&Atilde;O:</span>'
   cHtmlWF += '<br />'
   cHtmlWF += '<br /><input type="radio" checked name="%aprovacao%" value="Sim">Sim.'
   //cHtmlWF += '<br /><input type="radio" checked name="aprovacao" value="Sim">Sim.'
   cHtmlWF += '<br>'
   cHtmlWF += '<input type="radio" name="%aprovacao%" value="Nao">'
   //cHtmlWF += '<input type="radio" name="aprovacao" value="Nao">'
   cHtmlWF += 'Não(Favor especificar Motivo) <br />'
   cHtmlWF += '<br />'
   cHtmlWF += '	<br />'
   cHtmlWF += '<form id="form1" name="form1" method="post" action="">'
   cHtmlWF += '  <label>Motivo'
   cHtmlWF += '  '
   //cHtmlWF += '  <textarea name="textarea" cols="80" rows="5">%lbmotivo%</textarea> '
   //cHtmlWF += '  <textarea name="textarea" cols="80" rows="5"><input type=hidden name="lbmotivo" value = " "></textarea>'
   //cHtmlWF += '  <textarea name="textarea" cols="80" rows="5">%LMOTIVO%</textarea>'
   cHtmlWF += '  <textarea name="textarea" cols="80" rows="5">Informe aqui o motivo da aprovação ou reprovação</textarea>'
   cHtmlWF += '  </label>'
   cHtmlWF += '</form>'
   cHtmlWF += '<br />'
   
   //Form para Enviar o E-mail de volta ao Protheus - Five Solutions - 04/05/2010
   /* TRECHO ORIGINAL
   cHtmlWF += '<form id="form2" name="form2" method="post" action="">'
   cHtmlWF += '  <input name="Enviar" type="submit" id="Enviar" value="Enviar" />'
   cHtmlWF += '  <input name="Limpar" type="submit" id="Limpar" value="Limpar" />'
   cHtmlWF += '</form>'
   */

   //Form para Enviar o E-mail de volta ao Protheus - Five Solutions - 04/05/2010 
   // TRECHO MODIFICADO
   
   cHtmlWF += '<form id="form2" name="form2" method="post" action="mailto:%25WFMailTo%25">'
   cHtmlWF += '  <input name="Enviar" type="submit" id="Enviar" value="Enviar" />'
   //cHtmlWF += '<form id="form2" name="form2" method="post" action="WFHTTPRET.APL">'
   //cHtmlWF += '  <input type=hidden name="Enviar" value="Enviar">'
   
   cHtmlWF += '  <input name="Limpar" type="submit" id="Limpar" value="Limpar" />'
   cHtmlWF += '</form>'
   /*
   cHtmlWF += '<form action="mailto:%WFMailTo%" method="POST" name="form2">    '
   cHtmlWF += '<p>'
   cHtmlWF += '<input type="button" name="B1" value="Enviar Aprovação" onclick="ValidaObs(form2)"> '
   cHtmlWF += '<input type="reset" name="B2" value="Limpar"> '
   cHtmlWF += '</p> '
   cHtmlWF += '</form>'   
   */
   //cHtmlWF += '  <input onclick="ValidaObs(form2)" type="button" value="Enviar" name="B3">'   
   //cHtmlWF += '<p><a style="color: blue; text-decoration: underline; text-underline: single" href="' +Alltrim(_cHttp)+'/messenger/emp' +cEmpAnt  + '/' + cCdUsEmp + '/' + alltrim(cMailID) + '.htm">clique aqui</a></span></p>'
   //cHtmlWF += '<p><a style="color: blue; text-decoration: underline; text-underline: single" href="http://ITACOLOMY-ACER/messenger/emp' +cEmpAnt  + '/' + cCdUsEmp + '/' + alltrim(cMailID) + '.htm">clique aqui</a></span></p>'

/*
<form name="FrontPage_Form1" method="post" action="mailto:%25WFMailTo%25">
	<!--mstheme--></font>
	<h4><!--mstheme--><font face="trebuchet ms,arial,helvetica" color="#000000">
	<font face="Verdana" color="#0000ff"><strong>Avaliação:</strong></font><!--mstheme--></font></h4>
	<!--mstheme--><font face="trebuchet ms,arial,helvetica">
	<p>
	<input onclick="ValidaObs(FrontPage_Form1)" type="button" value="Enviar" name="B3">
	<input type="reset" value="Limpar" name="B4"> </p>
</form>
*/   
   
   cHtmlWF += '</body>'
   cHtmlWF += '</html>'
   
   //MemoWrite("C:\Protheus10\Workflow\ApPedCom.htm",cHtmlWF)
   //ConOut("Gerou html em C:\Protheus10\Workflow\ApPedCom.htm")

   MemoWrite("\Workflow\ApPedCom.htm",cHtmlWF)
   ConOut("Gerou NOVO - html em \Workflow\ApPedCom.htm")
Return

User Function RetAprvPC(oProcess) //(oProcess,xEmpWF,xFilWF) // RetAprvPC(,oProcess)
   /*
    Esta função é executada no retorno da mensagem enviada
    pelo destinatário. O Workflow recria o processo em que
    parou anteriormente na função MT120GOK e repassa a
    variável objeto oProcess por parâmetro.
   */
   
   RPCSetEnv("91","01","","","","",{"SC7"})
   
   ConOut("Processando Retorno da Liberação do Pedido de Compras - Função RetAprvPC")
   
   //If oProcess <> NIL
		
      //cFilAnt		:= alltrim(oProcess:oHtml:RetByName("CFILANT"))
      //xEmpresa      := alltrim(oProcess:oHtml:RetByName("CEMPANT"))
      
      //WFPrepEnv({xEmpresa,cFilAnt}) 
      
	  cChaveSCR	:= alltrim(oProcess:oHtml:RetByName("CHAVE"))
	  cOpc     	:= alltrim(oProcess:oHtml:RetByName("OPC")) //Obtem aResposta - Aprovado Sim ou Não partir do html
	  cObs     	:= alltrim(oProcess:oHtml:RetByName("OBS"))
	  //cWFID     	:= alltrim(oProcess:oHtml:RetByName("WFID"))
	  //cWF       	:= alltrim(oProcess:oHtml:RetByName("WF")) //cWF -> ' '-Nao Enviado / 3-rejeitado por saldo insuficiente no retorno
		
	  //oProcess:Finish() // FINALIZA O PROCESSO
   /*		
		U_CONSOLE("2 - Chave :" + cChaveSCR)
		U_CONSOLE("2 - Opc   :" + cOpc)
		ConOut('cObs: '+cObs)
   */		
   //EndIf
	
   //If cOpc	$ "S|N"  // Aprovacao S-Sim N-Nao
		
		// Posiciona na tabela de Alcadas
		/*
		DbSelectArea("SCR")
		DbSetOrder(2)
		DbSeek(cChaveSCR)
		
		If !Found()
			Return .t.
		EndIf
		
		While !Eof() .And. SCR->(CR_FILIAL+CR_TIPO+CR_NUM+CR_USER) == cChaveSCR .And.;
			SCR->CR_STATUS == '03'
			DbSkip()
		EndDo
		
		IF TRIM(SCR->CR_WFID) <> TRIM(cWFID)
			//"Este processo nao foi encontrado e portanto deve ser descartado
			// abre uma notificacao a pessoa que respondeu
			
			U_CONSOLE("2 - Primeira Critica: "+IIF(!FOUND(),"Sim","Nao") )
			U_CONSOLE("2 - Segunda Critica: "+IIF(TRIM(SCR->CR_WFID) <> TRIM(cWFID),"Sim","Nao") )
			Return .T.
		ENDIF
		
		If SCR->(CR_FILIAL+CR_TIPO+CR_NUM+CR_USER) <> cChaveSCR
			ConOut('2- Fim de Processo(Mesmo Aprovador)')
			Return // Fim de Processo
		EndIf
		
		Reclock("SCR",.F.)
		SCR->CR_WF		:= "2"			// Status 2 - respondido
		MSUnlock()
		
		If !Empty(SCR->CR_DATALIB) .And. SCR->CR_STATUS$"03#04#05"
			U_CONSOLE("2 - Semaforo Verde" )
			Return .T.
		EndIf
		
		// REPosiciona na tabela de Alcadas
		DbSelectArea("SCR")
		DbSetOrder(2)
		DbSeek(cChaveSCR)
		
		While !Eof() .And. SCR->(CR_FILIAL+CR_TIPO+CR_NUM+CR_USER) == cChaveSCR .And.;
			SCR->CR_STATUS == '03'
			DbSkip()
		EndDo
		
		lLiberou := U_MaAlcDoc({SCR->CR_NUM,"PC",nTotal,SCR->CR_APROV,,,,,,,cObs},dDataBase,If(cOpc=="S",4,6), cWF)
		
		_lProcesso := .T.
		
		If !lLiberou
			ConOut('Ainda Nao Liberou')
		Else
			ConOut('Pedido Liberado - FILIAL + PEDIDO: '+Left(cChaveSCR,2)+'/'+SubStr(cChaveSCR,5,6))
			// Flegar a liberacao para o Workflow enviar a notificacao de pedidos
			// aprovados:
		*/	
	    
	 If cOpc == "S"  // Aprovacao S-Sim N-Nao
			_cExec := " UPDATE "+RetSqlName('SC7')+" SET C7_CONAPRO = 'L' "
			_cExec += " WHERE C7_FILIAL = '"+Left(cChaveSCR,2)+"' "
			_cExec += " AND   C7_NUM = '"+SubStr(cChaveSCR,3,6)+"' "
			_cExec += " AND   D_E_L_E_T_ = ' ' "
			
			TcSqlExec( _cExec )
			
			_cExec := " COMMIT "
			
			TcSqlExec( _cExec )
        
        cTexto := "Pedido de Compras "+SubStr(cChaveSCR,3,6)+" Aprovado "
        cCodStatus := "100300"
        cCdUsEmp     := Alltrim(WFCodUser(cUserName))
        oProcess:Track(cCodStatus, cTexto , cCdUsEmp) // Rastreabilidade
        RastreiaWF(oProcess :fProcessID + '.' + oProcess:fTaskID,oProcess:fProcCode, '100300',"Pedido de Compras "+SubStr(cChaveSCR,3,6)+" Aprovado ", cUserName)
        
	 Else		
        cTexto := "Pedido de Compras "+SubStr(cChaveSCR,3,6)+" Reprovado "
        cCodStatus := "100400"
        cCdUsEmp     := Alltrim(WFCodUser(cUserName))
        oProcess:Track(cCodStatus, cTexto , cCdUsEmp) // Rastreabilidade
        RastreiaWF(oProcess :fProcessID + '.' + oProcess:fTaskID,oProcess:fProcCode, '100400',"Pedido de Compras "+SubStr(cChaveSCR,3,6)+" Reprovado ", cUserName)
     EndIf


		//EndIf
		
		//U_RastroWf('SC',SubStr(cChaveSCR,5,6), '' )
		
   //EndIf
	
   //	U_CONSOLE("2 - Semaforo Verde" )
	   

   // Adicione as informacões a serem incluídas na rastreabilidade.
   /*
   cTexto := "Resultado da Aprovacao do Pedido de Compras"
   cProcWFCOM := "000002" // oProcess:fProcCode  - Código do Processo WorkFlow
   cCodStatus := "100200" // Código do cadastro de status de processo.
   // Através do atributo :cRetFrom da classe twfprocess, é possível obter o endereço
   // eletrônico de quem respondeu a mensagem. Obtenha-o para adicioná-lo à rastreabilidade.
   oProcess:Track(cCodProces+cCodStatus, cTexto, oProcess:cRetFrom)

   // Adicione as informacões a serem incluídas na rastreabilidade.
   
   cTexto := "Finalizando o processo..."
   cProcWFCOM := "000003" //oProcess:fProcCode -  Código do Processo WorkFlow
   cCodStatus := "100300" // Código do cadastro de status de processo
   */
   cAprovador := oProcess:cRetFrom
   //oProcess:Track(cCodProces+cCodStatus, cTexto, oProcess:cRetFrom) // Rastreabilidade
   

   cTexto := "Processo do Pedido de Compras "+SubStr(cChaveSCR,3,6)+" FINALIZADO "
   cCodStatus := "100500"
   oProcess:Track(cCodStatus, cTexto , cCdUsEmp) // Rastreabilidade
   ConOut("Processo de Aprovacao realizado por "+cAprovador)
   RastreiaWF(oProcess :fProcessID + '.' + oProcess:fTaskID,oProcess:fProcCode, '100500',"Processo do Pedido de Compras "+SubStr(cChaveSCR,3,6)+" FINALIZADO ", cUserName) 
   oProcess:Finish() // FINALIZA O PROCESSO


Return

Static Function GrCorpAp(cNumPC)

   cHtmCrp := '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
   cHtmCrp += '<html xmlns="http://www.w3.org/1999/xhtml">'
   cHtmCrp += '<head>'
   cHtmCrp += '<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />'
   cHtmCrp += '<title>Untitled Document</title>'
   cHtmCrp += '</head>'
   cHtmCrp += ''
   cHtmCrp += '<body>'
   cHtmCrp += 'Favor abrir arquivo anexo e realizar aprova&ccedil;&atilde;o ou reprova&ccedil;&atilde;o do pedido de compras n&uacute;mero %PCNUM% ' //'+cNumPC+''
   cHtmCrp += '</body>'
   cHtmCrp += '</html>'
   
   MemoWrite("\workflow\CorpoApCp.htm",cHtmCrp)
   
Return

Static Function LinkApPc

   LkApPc := '<html>'
   LkApPc += '<head>'
   LkApPc += '<title>Untitled Document</title>'
   LkApPc += '<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">'
   LkApPc += '</head>'
   LkApPc += '<body bgcolor="#FFFFFF" bgproperties="fixed" background="C:/mp8/workflow/fundo8.JPG">'
   LkApPc += '<body>'
   LkApPc += '<form name="form1" method="post" action="mailto:%WFMailTo%">'
   LkApPc += '  <p>&nbsp;</p>'
   LkApPc += '  <p>&nbsp;</p>'
   //LkApPc += '  <p>Sr. !usuario! </p>'
   LkApPc += '  <p>Favor acessar o <a href="!proc_link!" title="!titulo!">processo</a>  de workflow referente ao !referente! </p>'
   //  e realizar aprova&ccedil;&atilde;o ou reprova&ccedil;&atilde;o do pedido de pompras n&uacute;mero  '+cNumPC+'</p>'
   //LkApPc += '  <p>Favor queira acessar o <a href="!proc_link!" title="!titulo!">processo</a> '
   //LkApPc += '    de workflow referente à !referente! </p>'
   LkApPc += '</form>'
   LkApPc += '</body>'
   LkApPc += '</html>'
   
   MemoWrite("\workflow\WFLinkApPc.htm",LkApPc)
   
Return

Static Function NovoPC(cNumPC,cNumPC,cForPC,cLjPC)
   
   cHtmNwPC := '<html>'
   cHtmNwPC += '<head>'
   cHtmNwPC += '<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">'
   cHtmNwPC += '<title>Aprovação de Pedido de Compras</title>'
   cHtmNwPC += '<style type="text/css"> '
   cHtmNwPC += 'a:hover {  text-decoration: underline}'
   cHtmNwPC += 'a:link {  text-decoration: none}'
   cHtmNwPC += 'a:active {  }'
   cHtmNwPC += '.texto {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 12px;'
   cHtmNwPC += '	color: #333333;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: normal;'
   cHtmNwPC += '}'
   cHtmNwPC += '.Scroll {'
   cHtmNwPC += 'SCROLLBAR-FACE-COLOR: #DEDEDE; '
   cHtmNwPC += 'SCROLLBAR-HIGHLIGHT-COLOR: #DEDEDE; '
   cHtmNwPC += 'SCROLLBAR-SHADOW-COLOR: #ffffff; '
   cHtmNwPC += 'SCROLLBAR-3DLIGHT-COLOR: #ffffff; '
   cHtmNwPC += 'SCROLLBAR-ARROW-COLOR:  #ffffff; '
   cHtmNwPC += 'SCROLLBAR-TRACK-COLOR: #ffffff; '
   cHtmNwPC += 'SCROLLBAR-DARKSHADOW-COLOR: #DEDEDE; '
   cHtmNwPC += '}'
   cHtmNwPC += '.combo {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	margin: 1px;'
   cHtmNwPC += '	padding: 1px;'
   cHtmNwPC += '	border: 1px solid #000000;'
   cHtmNwPC += '}'
   cHtmNwPC += '.comboselect {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	color: #666666;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	border: 1px #CCCCCC double;'
   cHtmNwPC += '}'
   cHtmNwPC += '.links {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	color: #CC0000;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.links-clientes {'
   cHtmNwPC += 'font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: #003366; text-decoration: none'
   cHtmNwPC += '}'
   cHtmNwPC += '.textobold {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	color: #003366;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.textoItalico {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	color: 7F7F7F;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-style: italic;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.titulo {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 16px;'
   cHtmNwPC += '	color: #19167D;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.links-detalhes {'
   cHtmNwPC += 'font-family: Arial, Helvetica, sans-serif; font-size: 10px; color: FF0000; text-decoration: none'
   cHtmNwPC += '}'
   cHtmNwPC += '.TituloMenor {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 12px;'
   cHtmNwPC += '	color: #19167D;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.Botoes {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 10px;'
   cHtmNwPC += '	font-weight: normal;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	margin: 2px;'
   cHtmNwPC += '	padding: 2px;'
   cHtmNwPC += '	cursor: hand;'
   cHtmNwPC += '	border: 1px outset #000000;'
   cHtmNwPC += '}'
   cHtmNwPC += '.TarjaTopoMenu {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	height: 6px;'
   cHtmNwPC += '	background-image: url(http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/apmenu-right.jpg);'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoMenu {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-image: url(http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/apmenu-right.jpg);'
   cHtmNwPC += '}'
   cHtmNwPC += '.FonteMenu {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 12px;'
   cHtmNwPC += '	color: #19167D;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FonteSubMenu {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 12px;'
   cHtmNwPC += '	color: #19167D;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoSubMenu {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-image: url(http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/apmenu-right.jpg);'
   cHtmNwPC += '}'
   cHtmNwPC += '.DivisoriaOpçõesMenu {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #6680A6;'
   cHtmNwPC += '	background-image: url(http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/apmenu-right.jpg);'
   cHtmNwPC += '}'
   cHtmNwPC += '.TarjaTopoLogin {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #426285;'
   cHtmNwPC += '	height: 6px;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoLogin {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #F7F7F7;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoClaro {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #fbfbfb;'
   cHtmNwPC += '}'
   cHtmNwPC += '.TituloDestaques {'
   cHtmNwPC += '	font-family: Verdana, Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	color: #19167D;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.TextoDestaques {'
   cHtmNwPC += '	font-family: Verdana, Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 10px;'
   cHtmNwPC += '	color: #777777;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: normal;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoDestaques {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #E5E5E5;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoPontilhado {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-image: url(http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/pontilhado.gif);'
   cHtmNwPC += '	height: 5px;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoPontilhadoVertical {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-image: url(http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/pontilhado_vertical.gif);'
   cHtmNwPC += '	height: 5px;'
   cHtmNwPC += '}'
   cHtmNwPC += '.TituloTabelas {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 12px;'
   cHtmNwPC += '	color: #FFFFFF;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoTituloTabela {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #495E73;'
   cHtmNwPC += '}'
   cHtmNwPC += '.TarjaTopoCor {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	height: 6px;'
   cHtmNwPC += '	background-image: url(http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/pontilhado.gif);'
   cHtmNwPC += '	background-color: #6699CC;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoTabelaDestaques {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #495E73;'
   cHtmNwPC += '}'
   cHtmNwPC += '.comboselect-pequeno {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	color: #666666;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	border: 1px solid #CCCCCC;'
   cHtmNwPC += '	width: 132px;'
   cHtmNwPC += '	clear: none;'
   cHtmNwPC += '	float: none;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	left: 1px;'
   cHtmNwPC += '	top: 1px;'
   cHtmNwPC += '	right: 1px;'
   cHtmNwPC += '	bottom: 1px;'
   cHtmNwPC += '	clip:  rect(1px 1px 1px 1px);
   cHtmNwPC += '}'
   cHtmNwPC += '.comboselect-grande {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	color: #666666;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	border: 1px #CCCCCC double;'
   cHtmNwPC += '	width: 415px;'
   cHtmNwPC += '}'
   cHtmNwPC += '.texto-layer {'
   cHtmNwPC += 'font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; color: #000000; text-decoration: none'
   cHtmNwPC += '}'
   cHtmNwPC += '.tituloAvaliacao {'
   cHtmNwPC += '	font-family: Verdana, Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 9px;'
   cHtmNwPC += '	color: #19167D;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '	vertical-align: middle;'
   cHtmNwPC += '	text-align: center;'
   cHtmNwPC += '	line-height: 12px;'
   cHtmNwPC += '}'
   cHtmNwPC += '</style>'
   cHtmNwPC += '<link href="estilo.css" rel="stylesheet" type="text/css">'
   cHtmNwPC += '<style type="text/css"> '
   cHtmNwPC += '<!--'
   cHtmNwPC += '.style2 {font-size: 18px}'
   cHtmNwPC += '.style14 {	color: #FF6600;'
   cHtmNwPC += '	font-size: 24px;'
   cHtmNwPC += '}'
   cHtmNwPC += '.style3 {	font-size: 36px;'
   cHtmNwPC += '	color: #666666;'
   cHtmNwPC += '}'
   cHtmNwPC += '.style16 {color: #FF6600; font-size: 24px; font-weight: bold; }'
   cHtmNwPC += '.style17 {'
   cHtmNwPC += '	color: #19167D;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.style18 {font-size: 12px}'
   cHtmNwPC += '.style19 {color: #333333}'
   cHtmNwPC += '-->'
   cHtmNwPC += '</style>'
   cHtmNwPC += '</head>'
   cHtmNwPC += '<body>'
   cHtmNwPC += '<table width="98%" border="0" cellpadding="0" cellspacing="0">'
   cHtmNwPC += '  <tr>'
   cHtmNwPC += '    <td height="89" background="http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/topo_04.gif"><div align="center"><img src="http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/topo_03.gif" width="558" height="89"89">  <img src="http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/Logo_Vitarella.jpg" width="104" height="63"></div></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '  <tr background="http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/pontilhado.gif">'
   cHtmNwPC += '    <td height="5" class="TarjaTopoCor"><img src="http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/transparente.gif" width="1" height="1"></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '</table>'
   cHtmNwPC += '<table width="96%" border="0" align="center" cellpadding="0" cellspacing="1">'
   cHtmNwPC += '  <tr>'
   cHtmNwPC += '    <td width="69%"><div align="justify"><span class="titulo"><span class="style3">%TIPOPED%</span> </span></div></td>'
   cHtmNwPC += '    <td width="31%"><div align="right"><span class="TituloMenor"><span class="textoItalico style2">No. %C7_NUM% - %C7_FILIAL%&nbsp;&nbsp;</span></span></div></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '</table>'
   cHtmNwPC += '<table width="96%" border="1" align="center" cellpadding="6" cellspacing="0" bordercolor="#E5E5E5" bgcolor="#F7F7F7" >'
   cHtmNwPC += '  <tr>'
   cHtmNwPC += '    <td width="100%" height="606" bordercolor="#FFFFFF" class="titulo"><table border="0" cellpadding="2" cellspacing="1" width="100%">'
   cHtmNwPC += '      <tr bordercolor="#000000" bgcolor="#CCCCCC">'
   cHtmNwPC += '        <td colspan="2" bgcolor="#EBEBEB" class="TituloMenor"><div align="center" class="style14"><span class="style16"><font size="3">&raquo; </font></span>Dados / COMAFAL </div></td>'
   cHtmNwPC += '        <td colspan="2" bgcolor="#EBEBEB" class="style10 texto"><div align="center"><span class="style16"><font size="3">&raquo; </font>Dados / Fornecedor</span></div></td>'
   cHtmNwPC += '      </tr>'
   cHtmNwPC += '      <tr bordercolor="#000000">'
   cHtmNwPC += '        <td colspan="2" class="TituloMenor">%M0_NOMECOM%</td>'
   cHtmNwPC += '        <td colspan="2" class="style10 texto"><strong>%A2_NOME% - %A2_CODLOJA% - %A2_IE%</strong></td>'
   cHtmNwPC += '      </tr>'
   cHtmNwPC += '      <tr bordercolor="#000000">'
   cHtmNwPC += '        <td colspan="2" class="TituloMenor">%M0_ENDCOB%</td>'
   cHtmNwPC += '        <td colspan="2" class="style10 texto"><strong>%A2_END% - %A2_BAIRRO%</strong></td>'
   cHtmNwPC += '      </tr>'
   cHtmNwPC += '      <tr bordercolor="#000000">'
   cHtmNwPC += '        <td colspan="2" class="TituloMenor">%M0_CEPCOB% - %M0_CIDENT% - %M0_ESTENT% </td>'
   cHtmNwPC += '        <td colspan="2" class="style10 texto"><strong>%A2_MUN% - %A2_EST% - %A2_CEP% - %A2_CGC% </strong></td>'
   cHtmNwPC += '      </tr>'
   cHtmNwPC += '      <tr bordercolor="#000000">'
   cHtmNwPC += '        <td colspan="2" class="TituloMenor">%M0_TEL% - %M0_FAX% </td>'
   cHtmNwPC += '        <td colspan="2" class="style10 texto"><strong>%A2_CONTATO% - %A2_TEL% - %A2_FAX% </strong></td>'
   cHtmNwPC += '      </tr>'
   cHtmNwPC += '      <tr bordercolor="#000000">'
   cHtmNwPC += '        <td height="22" colspan="2" class="TituloMenor"><b>%M0_CGC% - %M0_INSC% </b></td>'
   cHtmNwPC += '        <td colspan="2" class="style10 texto">&nbsp;</td>'
   cHtmNwPC += '      </tr>'
   cHtmNwPC += '        <tr bgcolor="#EBEBEB">'
   cHtmNwPC += '          <td colspan="4" class="TituloMenor"><div align="center" class="style14"><span class="style16"><font size="3">&raquo; </font></span>Dados do Pedido</div></td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td width="15%" class="TituloMenor">Data de Emiss&atilde;o</td>'
   cHtmNwPC += '          <td width="35%" bgcolor="#FFFFFF" class="style21 style18 style19">%C7_EMISSAO%</td>'
   cHtmNwPC += '          <td width="12%" class="style10 texto"><span class="style17">Total do IPI </span></td>'
   cHtmNwPC += '          <td width="38%" bgcolor="#FFFFFF" class="texto style10"><div align="center"><span class="style23">%TOTIPI%</span></div></td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td class="TituloMenor">Condi&ccedil;&atilde;o de Pgto. </td>'
   cHtmNwPC += '          <td bgcolor="#FFFFFF" class="style21 style18 style19">%E4_DESCRI%</td>'
   cHtmNwPC += '          <td class="style10 texto"><span class="style17">Total do ICMS </span></td>'
   cHtmNwPC += '          <td bgcolor="#FFFFFF" class="texto style10"><div align="center"><span class="style23">%TOTICMS%</span></div></td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td class="TituloMenor">Comprador</td>'
   cHtmNwPC += '          <td bgcolor="#FFFFFF" class="style21 style18 style19">%C7_USER%</td>'
   cHtmNwPC += '          <td class="texto style10"><span class="style17">Frete</span></td>'
   cHtmNwPC += '          <td bgcolor="#FFFFFF" class="texto style10"><div align="center"><span class="style23">%TOTFRETE%</span></div></td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td class="TituloMenor">Total das Mercadorias </td>'
   cHtmNwPC += '          <td bgcolor="#FFFFFF" class="style21 style18 style19"><div align="center">%TOTPED%</div></td>'
   cHtmNwPC += '          <td class="style10 texto"><span class="style17">Despesas</span></td>'
   cHtmNwPC += '          <td bgcolor="#FFFFFF" class="style10 texto"><div align="center"><span class="style23">%TOTDESP%</span></div></td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td class="TituloMenor"><span class="TituloMenor"><strong>Total com Impostos </strong></span></td>'
   cHtmNwPC += '          <td bgcolor="#FFFFFF" class="FonteSubMenu style18 style19"><div align="center">%TOTIMP%</div></td>'
   cHtmNwPC += '          <td class="style10 texto"><span class="style17">Seguro</span></td>'
   cHtmNwPC += '          <td bgcolor="#FFFFFF" class="style10 texto"><div align="center"><span class="style23">%TOTSEG%</span></div></td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td class="TituloMenor">Total dos Descontos </td>'
   cHtmNwPC += '          <td bgcolor="#FFFFFF"><div align="center"><span class="style18 style19">%TOTDESC%</span></div></td>'
   cHtmNwPC += '          <td class="TituloMenor">Total Geral </td>'
   cHtmNwPC += '          <td bgcolor="#FFFFFF" class="texto"><div align="center">%TOTGER%</div></td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '      </table>'
   cHtmNwPC += '      <table border="0" cellspacing="1" width="100%">'
   cHtmNwPC += '        <tr bgcolor="#E4E4E4" class="TituloMenor">'
   
   /*
   cHtmNwPC += '          <td width="5%"  align="center"><b>Item</b></td>'
   cHtmNwPC += '          <td width="10%" align="center"><b>C&oacute;digo</b></td>'
   cHtmNwPC += '          <td width="23%" align="center" bgcolor="#E4E4E4"><b>Descri&ccedil;&atilde;o</b></td>'
   cHtmNwPC += '          <td width="5%" align="center"><b>Um</b></td>'
   cHtmNwPC += '          <td width="9%" align="center"><b>Quant</b></td>'
   cHtmNwPC += '          <td width="8%" align="center"><b>Vl.Unit&aacute;rio</b></td>'
   cHtmNwPC += '          <td width="5%" align="center"><b>IPI</b></td>'
   cHtmNwPC += '          <td width="10%" align="center"><b>Vl.Total</b></td>'
   cHtmNwPC += '          <td width="8%" align="center">Entrega Em </td>'
   cHtmNwPC += '          <td width="9%" align="center">C.Custo</td>'
   cHtmNwPC += '          <td width="8%" align="center">Cota&ccedil;&atilde;o</td>'
   */
   
   cHtmNwPC += '          <td width="10%" align="center"><b>C&oacute;digo</b></td>'
   cHtmNwPC += '          <td width="23%" align="center" bgcolor="#E4E4E4"><b>Descri&ccedil;&atilde;o</b></td>'
   cHtmNwPC += '          <td width="5%" align="center"><b>Um</b></td>'
   cHtmNwPC += '          <td width="9%" align="center"><b>Quant</b></td>'
   cHtmNwPC += '          <td width="8%" align="center"><b>Vl.Unit&aacute;rio</b></td>'
   cHtmNwPC += '          <td width="10%" align="center"><b>Vl.Total</b></td>'
   cHtmNwPC += '          <td width="8%" align="center">Natureza</td>'
   cHtmNwPC += '          <td width="27%" align="center">Descrição</td>'
   
      
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tbody>'
   cHtmNwPC += '          <tr bgcolor="#FFFFFF" class="texto-layer">'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.1%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.2%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.3%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.4%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.5%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.6%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.7%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.8%</div></td>'
   
   //cHtmNwPC += '            <td>%t1.9%</td>'
   //cHtmNwPC += '            <td>%t1.10%</td>'
   //cHtmNwPC += '            <td>%t1.11%</td>'
   
   cHtmNwPC += '          </tr>'
   cHtmNwPC += '          <tr>'
   cHtmNwPC += '            <td colspan="11" class="texto-layer" background="http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/pontilhado.gif"><span class="TituloMenor"><img src="http://192.168.1.215:1080/workflow/desenvolvimento/Figuras/transparente.gif" width="10" height="3"></span></td>'
   cHtmNwPC += '          </tr>'
   cHtmNwPC += '        </tbody>'
   cHtmNwPC += '      </table>'
   cHtmNwPC += '      <table border="0" width="100%">'
   cHtmNwPC += '        <tr height="21">'
   cHtmNwPC += '          <td colspan="2">&nbsp;</td>'
   cHtmNwPC += '        </tr>'

   cHtmNwPC += '<form method="post" action="mailto:%WFMailTo%" name="FrontPage_Form1">        '
   //cHtmNwPC += '<form method="post" action="mailto:%25WFMailTo%25" name="FrontPage_Form1">        '
   cHtmNwPC += '	<tr height="21">'
   cHtmNwPC += '          <td width="14%"><p align="center" class="TituloMenor">Aceite:</p></td>'
   cHtmNwPC += '          <td width="86%">            <input type="radio" checked name="OPC" value="%OPC%S">'
   cHtmNwPC += '            <span class="texto">Sim.</span></td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tr height="21">'
   cHtmNwPC += '          <td width="14%"></td>'
   cHtmNwPC += '          <td width="86%">            <input type="radio" name="OPC" value="%OPC%N">'
   cHtmNwPC += '            <span class="texto">N&atilde;o. (<strong>Especificar Motivo</strong>) </span></td>'
   cHtmNwPC += '        </tr>'
/* TRECHO OBSERVAÇÃO ORIGINAL
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td width="14%" height="100"></td>'
   cHtmNwPC += '          <td width="86%" height="100">'
   cHtmNwPC += '          <textarea name="OBS" cols="61" rows="4" wrap="PHYSICAL" class="style5"'
   cHtmNwPC += '                          >cQryPC  += " ORDER BY SC7.C7_ITEM"</textarea>           </td>'
   cHtmNwPC += '        </tr>'
  */

   // TRECHO OBSERVAÇÃO MODIFICADO
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td width="14%" height="100"></td>'
   cHtmNwPC += '          <td width="86%" height="100">'
   //cHtmNwPC += '          <textarea name="OBS" cols="61" rows="4" wrap="PHYSICAL" class="style5"'
   cHtmNwPC += '          <textarea name="OBS" cols="100" rows="4" wrap="PHYSICAL" class="style5"'
   cHtmNwPC += '                          >%OBS%</textarea>           </td>'
   cHtmNwPC += '        </tr>'
  
  
   cHtmNwPC += '        <tr height="27">'
   cHtmNwPC += '          <td width="14%"></td>'
   cHtmNwPC += '          <td width="86%"> <input  name="B1" type="submit" class="Botoes" value="Enviar"> <input name="B2" type="reset" class="Botoes"  value="Limpar">  </td>'
   cHtmNwPC += '        </tr>'
   
   /*
   cHtmNwPC += '<p>'
   cHtmNwPC += '    <input type="submit" name="Submit" value="Enviar">'
   cHtmNwPC += '  </p>'
   */
   
   cHtmNwPC += '        <input type="hidden" name="CHAVE" 		value="%CHAVE%">'
   cHtmNwPC += '        <input type="hidden" name="CFILANT" 	value="%CFILANT%">'
   cHtmNwPC += '        <input type="hidden" name="WF" 		value="%WF%">'
   cHtmNwPC += '        <input type="hidden" name="WFID" 		value="%WFID%">'
   cHtmNwPC += '      </table></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '</table>'
   cHtmNwPC += '<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1">'
   cHtmNwPC += '  <tr>'
   cHtmNwPC += '    <td bgcolor="#EBEBEB" class="style16"><font size="3">&raquo; </font>Observa&ccedil;&otilde;es</td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '</table>'
   cHtmNwPC += '<table width="96%" border="1" align="center" cellpadding="6" cellspacing="0" bordercolor="#E5E5E5" bgcolor="#F7F7F7" >'
   cHtmNwPC += '  <tr>'
   cHtmNwPC += '    <td width="100%" bordercolor="#FFFFFF" class="titulo"><p>%OBSPED%</p>'
   cHtmNwPC += '    <p>%OBSFORM%</p>'
   cHtmNwPC += '    <p>Somente aceitaremos a mercadoria se na sua Nota Fiscal constar o n&uacute;mero do nosso Pedido de Compras. </p></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '</table>'
   cHtmNwPC += '<table width="96%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" bgcolor="#F7F7F7" >'
   cHtmNwPC += '  <tr>'
   cHtmNwPC += '    <td width="100%" bordercolor="#FFFFFF"><div align="right" class="texto-layer">WorkFlow @ Microsiga Protheus</div></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '</table>'
   cHtmNwPC += '<p>&nbsp;</p>'
   cHtmNwPC += '</body>'
   cHtmNwPC += '</html>'

   MemoWrite("\Workflow\PCCOMAFAL.htm",cHtmNwPC)
   
   ConOut("Gerou NOVO - html em \Workflow\PCCOMAFAL.htm")

   //Five Solutions Consultoria - 14/05/2010 - Após gerar o Html roda a macro-substituições dos campos.
   
   For nPC := 1 To 2
   
      If nPC == 1
         cQryPC  := " SELECT COUNT(*) NQTDREG,SUM(SC7.C7_VALIPI) VLTOTIPI,SUM(SC7.C7_VALICM) VLRTICM,"
         cQryPC  += "        SUM(SC7.C7_TOTAL) NTOTMERC,SUM(SC7.C7_VLDESC) NDESCTOT"
      Else   
         cQryPC  := " SELECT SC7.C7_NUM,SC7.C7_EMISSAO,SC7.C7_FORNECE,SC7.C7_LOJA,SA2.A2_NOME,SC7.C7_COND,SA2.A2_NREDUZ,"
         cQryPC  += "        SC7.C7_ITEM,SC7.C7_PRODUTO,SB1.B1_DESC,SC7.C7_QUANT,SC7.C7_UM,SC7.C7_PRECO,SC7.C7_TOTAL,"
         cQryPC  += "        SC7.C7_DATPRF,SA2.A2_COD,SA2.A2_LOJA,SA2.A2_INSCR,SA2.A2_CGC,SA2.A2_CONTATO,SA2.A2_TEL,"
         cQryPC  += "        SA2.A2_FAX,SA2.A2_END,SA2.A2_BAIRRO,SA2.A2_MUN,SA2.A2_CEP,SA2.A2_CGC,SC7.C7_FRETE,"         
         cQryPC  += "        SC7.C7_SEGURO,SC7.C7_DESPESA,SC7.C7_VALFRE,SC7.C7_FILIAL,SC7.C7_USER,SE4.E4_DESCRI,"
         cQryPC  += "        SC7.C7_OBS,SA2.A2_EST,SC7.C7_DESCRI,SC7.C7_IPI,SC7.C7_VALIPI,SC7.C7_CC,SC7.C7_NUMCOT,"
         cQryPC  += "        SC7.C7_VLDESC,SC7.C7_VALICM,SC7.C7_X_NATUR,SED.ED_DESCRIC"
      EndIf
   
      cQryPC  += "   FROM "+RetSQLName("SC7")+" SC7,"+RetSQLName("SA2")+" SA2,"+RetSQLName("SB1")+" SB1,"
      cQryPC  += RetSQLName("SE4")+" SE4,"+RetSQLName("SED")+" SED"
      cQryPC  += "  WHERE SC7.C7_FILIAL = '"+xFilial("SC7")+"'"
      cQryPC  += "    AND SA2.A2_FILIAL = '"+xFilial("SA2")+"'"
      cQryPC  += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"
      cQryPC  += "    AND SE4.E4_FILIAL = '"+xFilial("SE4")+"'"
      
      cQryPC  += "    AND SED.ED_FILIAL = '"+xFilial("SED")+"'"
      cQryPC  += "    AND SC7.C7_X_NATUR = SED.ED_CODIGO"
         
      cQryPC  += "    AND SC7.C7_FORNECE = SA2.A2_COD"
      cQryPC  += "    AND SC7.C7_LOJA = SA2.A2_LOJA"
      cQryPC  += "    AND SC7.C7_PRODUTO = SB1.B1_COD"
      
      cQryPC  += "    AND SC7.C7_COND = SE4.E4_CODIGO"
   
      cQryPC  += "    AND SC7.C7_NUM = '"+cNumPC+"'"
      cQryPC  += "    AND SC7.C7_FORNECE = '"+cForPC+"'"
      cQryPC  += "    AND SC7.C7_LOJA = '"+cLjPC+"'"
   
      cQryPC  += "    AND SC7.D_E_L_E_T_ <> '*'"
      cQryPC  += "    AND SA2.D_E_L_E_T_ <> '*'"
      cQryPC  += "    AND SB1.D_E_L_E_T_ <> '*'"
      cQryPC  += "    AND SE4.D_E_L_E_T_ <> '*'"
      
      cQryPC  += "    AND SED.D_E_L_E_T_ <> '*'"
   
      If nPC > 1
         cQryPC  += " ORDER BY SC7.C7_ITEM"
      EndIf
      
      MemoWrite("PedidodeCompras.SQL",cQryPC)
   
      TCQuery cQryPC NEW ALIAS "TXC7"
      
      If nPC == 1
         DbSelectArea("TXC7")
         TCSetField("TXC7","VLTOTIPI","N",17,02)
         TCSetField("TXC7","VLRTICM","N",17,02)
         TCSetField("TXC7","NTOTMERC","N",17,02)
         TCSetField("TXC7","NDESCTOT","N",17,02)

         nReg      := TXC7->NQTDREG
         nVlrTtIPI := TXC7->VLTOTIPI
         nICMTtVl  := TXC7->VLRTICM
         nVlrMercad:= TXC7->NTOTMERC
         TXC7->(DbCloseArea())
      EndIf
   
   Next nPC
   
   TCSetField("TXC7","C7_QUANT","N",12,03)
   TCSetField("TXC7","C7_PRECO","N",14,02)
   TCSetField("TXC7","C7_TOTAL","N",14,02)
   TCSetField("TXC7","C7_VALIPI","N",14,02)
   TCSetField("TXC7","C7_SEGURO","N",14,02)
   TCSetField("TXC7","C7_DESPESA","N",14,02)
   TCSetField("TXC7","C7_VALFRE","N",14,02)
   TCSetField("TXC7","C7_DATPRF","D",08,00)
   TCSetField("TXC7","C7_EMISSAO","D",08,00)
   
   DbSelectArea("TXC7")

   cCdTarefa := oProcess:NewTask("Gerando E-mail da Aprovacao de Compras", "\Workflow\PCCOMAFAL.htm" ) 
   //oHtml   := oProcess:oHTML
   oProcess:oHTML:ValByName( "TIPOPED"	, "Pedido de Compras" )
   oProcess:ohtml:valbyname('C7_NUM', cNumPC )   
   oProcess:ohtml:valbyname('C7_FILIAL', TXC7->C7_FILIAL )   

   //Cabeçalho - Dados do Fornecedor:
   oProcess:oHTML:ValByName( "A2_NOME"		, Trim(TXC7->A2_NOME) )
   oProcess:oHTML:ValByName( "A2_CODLOJA"   , TXC7->A2_COD+' / '+TXC7->A2_LOJA)
   oProcess:oHTML:ValByName( "A2_END"       , "Endereco: "+TXC7->A2_END)
   oProcess:oHTML:ValByName( "A2_BAIRRO"    , "Bairro: "+TXC7->A2_BAIRRO)
   oProcess:oHTML:ValByName( "A2_MUN"       , "Municipio: "+TXC7->A2_MUN)
   oProcess:oHTML:ValByName( "A2_EST"       , "UF: "+TXC7->A2_EST)
   oProcess:oHTML:ValByName( "A2_CEP"       , "CEP: "+TXC7->A2_CEP)
   oProcess:oHTML:ValByName( "A2_CGC"       , "CGC: "+TXC7->A2_CGC)
   oProcess:oHTML:ValByName( "A2_IE"        , "IE: "+TXC7->A2_INSCR)
   oProcess:oHTML:ValByName( "A2_CONTATO"   , "Contato: "+TXC7->A2_CONTATO)
   oProcess:oHTML:ValByName( "A2_TEL"       , "Tel: "+TXC7->A2_TEL)
   oProcess:oHTML:ValByName( "A2_FAX"       , "Fax: "+TXC7->A2_FAX)
   //Cabeçalho - Dados do Cliente:
   oProcess:oHTML:ValByName( "M0_NOMECOM"   , SM0->M0_NOMECOM )
   oProcess:oHTML:ValByName( "M0_ENDCOB"    , "Endereco: "+SM0->M0_ENDCOB )
   oProcess:oHTML:ValByName( "M0_CEPCOB"    , "CEP: "+SM0->M0_CEPCOB )
   oProcess:oHTML:ValByName( "M0_CIDENT"    , "Cidade: "+SM0->M0_CIDENT )
   oProcess:oHTML:ValByName( "M0_ESTENT"    , "UF: "+SM0->M0_ESTENT )
   oProcess:oHTML:ValByName( "M0_TEL"       , "Tel: "+SM0->M0_TEL )
   oProcess:oHTML:ValByName( "M0_FAX"       , "Fax: "+SM0->M0_FAX )
   oProcess:oHTML:ValByName( "M0_CGC"       , "CGC: "+SM0->M0_CGC )
   oProcess:oHTML:ValByName( "M0_INSC"      , "IE: "+SM0->M0_INSC )

   // Dados do Pedido:
   oProcess:oHTML:ValByName( "C7_EMISSAO"   , TXC7->C7_EMISSAO )
   oProcess:oHTML:ValByName( "E4_DESCRI"    , TXC7->E4_DESCRI )
   oProcess:oHTML:ValByName( "C7_USER"      , UsrFullName(TXC7->C7_USER) )
   
   _nTotPed   := 0
   _nTotImp   := 0
   _nTotIPI   := 0
   _nTotICMS  := 0
   _nTotFrete := 0
   _nTotDesp  := 0
   _nTotSeg   := 0
   _nTotDesc  := 0
   cObs       := ""
   lUmaVez    := .T.

   DbSelectArea("TXC7")
   ProcRegua(nReg)
   nCont := 1
   While TXC7->(!Eof())
      
      IncProc("Gerando Itens p/Aprovação ... "+Alltrim(Str(nCont))+" / "+Alltrim(Str(nReg)))


	  //Five Solutions - 06/06/2011
      nUPrcComp := 0
      DbSelectArea("SBZ") //Indicadores de Produtos
      DbSetOrder(1) //BZ_FILIAL+BZ_COD
      If DbSeek(xFilial("SBZ")+TXC7->C7_PRODUTO)
         nUPrcComp := SBZ->BZ_UPRC
         If nUPrcComp == 0
            DbSelectArea("SB1") //Cadastro Genérico de Produtos
            DbSetOrder(1) //B1_FILIAL+B1_COD
            If DbSeek(xFilial("SB1")+TXC7->C7_PRODUTO)
               nUPrcComp := SB1->B1_UPRC
            EndIf
         EndIf
      EndIf
      
	  If TXC7->C7_PRECO > nUPrcComp .And. !lPrcMaior
	     lPrcMaior := .T.
	  EndIf
	  
	  AAdd( (oProcess:oHTML:ValByName( "t1.1"  )), TXC7->C7_PRODUTO)
	  AAdd( (oProcess:oHTML:ValByName( "t1.2"  )), TXC7->C7_DESCRI)
	  AAdd( (oProcess:oHTML:ValByName( "t1.3"  )), TXC7->C7_UM)
	  AAdd( (oProcess:oHTML:ValByName( "t1.4"  )), TRANSFORM(TXC7->C7_QUANT,'@E 999,999.999'))
	  AAdd( (oProcess:oHTML:ValByName( "t1.5"  )), TRANSFORM(TXC7->C7_PRECO,'@E 999,999.999'))
	  AAdd( (oProcess:oHTML:ValByName( "t1.6"  )), TRANSFORM(TXC7->C7_QUANT * TXC7->C7_PRECO,'@E 999,999.99'))
	  AAdd( (oProcess:oHTML:ValByName( "t1.7"  )), TXC7->C7_X_NATUR)
	  AAdd( (oProcess:oHTML:ValByName( "t1.8"  )), TXC7->ED_DESCRIC)
	  AAdd( (oProcess:oHTML:ValByName( "t1.9"  )), TRANSFORM(nUPrcComp,'@E 999,999.999'))//Five Solutions - 06/06/2011
     

      /*
	  AAdd( (oProcess:oHTML:ValByName( "t1.1"  )), TXC7->C7_ITEM)
	  AAdd( (oProcess:oHTML:ValByName( "t1.2"  )), TXC7->C7_PRODUTO)
	  AAdd( (oProcess:oHTML:ValByName( "t1.3"  )), TXC7->C7_DESCRI)
	  AAdd( (oProcess:oHTML:ValByName( "t1.4"  )), TXC7->C7_UM)
	  AAdd( (oProcess:oHTML:ValByName( "t1.5"  )), TRANSFORM(TXC7->C7_QUANT,'@E 999,999.999'))
	  AAdd( (oProcess:oHTML:ValByName( "t1.6"  )), TRANSFORM(TXC7->C7_PRECO,'@E 999,999.999'))
	  AAdd( (oProcess:oHTML:ValByName( "t1.7"  )), TRANSFORM(TXC7->C7_VALIPI,'@E 9,999.99'))
	  AAdd( (oProcess:oHTML:ValByName( "t1.8"  )), TRANSFORM(TXC7->C7_QUANT * TXC7->C7_PRECO,'@E 999,999.99'))
	  AAdd( (oProcess:oHTML:ValByName( "t1.9"  )), TXC7->C7_DATPRF)
	  AAdd( (oProcess:oHTML:ValByName( "t1.10"  )), TXC7->C7_CC)
	  AAdd( (oProcess:oHTML:ValByName( "t1.11"  )), TXC7->C7_NUMCOT)
	  */

	  
	  _nTotPed   += TXC7->C7_TOTAL
	  _nTotDesc  += TXC7->C7_VLDESC
	  _nTotImp   += ( TXC7->C7_TOTAL+TXC7->C7_VALIPI+TXC7->C7_SEGURO+TXC7->C7_DESPESA+TXC7->C7_VALFRE ) - TXC7->C7_VLDESC
	  _nTotIPI   += TXC7->C7_VALIPI
	  _nTotICMS  += TXC7->C7_VALICM
	  _nTotFrete += TXC7->C7_VALFRE
	  _nTotDesp  += TXC7->C7_DESPESA
	  _nTotSeg   += TXC7->C7_SEGURO
	
	  cTmpObs := Alltrim(TXC7->C7_OBS)
	  If cTmpObs <> Alltrim(Substr(cObs,4,150))
	     If lUmaVez
	         lUmaVez := .F.
	         cObs := cTmpObs
	     Else
	        cObs += ' - '+cTmpObs
	     EndIf
      EndIf
      DbSelectArea("TXC7")
      DbSkip()
      nCont ++
      
   EndDo
   DbSelectArea("TXC7")
   TXC7->(DbCloseArea())

   oProcess:oHTML:ValByName( "TOTPED"       , TRANSFORM(_nTotPed,'@E 999,999.99') )
   oProcess:oHTML:ValByName( "TOTDESC"      , TRANSFORM(_nTotDesc,'@E 999,999.99') )
   oProcess:oHTML:ValByName( "TOTIMP"       , TRANSFORM(_nTotImp,'@E 999,999.99') )
   oProcess:oHTML:ValByName( "TOTIPI"       , TRANSFORM(_nTotIPI,'@E 999,999.99') )
   oProcess:oHTML:ValByName( "TOTICMS"      , TRANSFORM(_nTotICMS,'@E 999,999.99') )
   oProcess:oHTML:ValByName( "TOTFRETE"     , TRANSFORM(_nTotFrete,'@E 999,999.99') )
   oProcess:oHTML:ValByName( "TOTDESP"      , TRANSFORM(_nTotDesp,'@E 999,999.99') )
   oProcess:oHTML:ValByName( "TOTSEG"       , TRANSFORM(_nTotSeg,'@E 999,999.99') )
   oProcess:oHTML:ValByName( "TOTGER"       , TRANSFORM(_nTotImp,'@E 999,999.99') )
   oProcess:oHTML:ValByName( "OBSPED", cObs )
   //oProcess:oHTML:ValByName( "OBSFORM", Formula(cObs2) )
   oProcess:oHTML:ValByName( "OBSFORM", " " )
   
Return





*************************************************************************************************************
*************************************************************************************************************

Static Function AprovPC(cNumPC,cNumPC,cForPC,cLjPC,oProcess)
   
   cHtmNwPC := '<html>'
   cHtmNwPC += '<head>'
   cHtmNwPC += '<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">'
   cHtmNwPC += '<title>Aprovação de Pedido de Compras</title>'
   cHtmNwPC += '<style type="text/css"> '
   cHtmNwPC += 'a:hover {  text-decoration: underline}'
   cHtmNwPC += 'a:link {  text-decoration: none}'
   cHtmNwPC += 'a:active {  }'
   cHtmNwPC += '.texto {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 12px;'
   cHtmNwPC += '	color: #333333;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: normal;'
   cHtmNwPC += '}'
   cHtmNwPC += '.Scroll {'
   cHtmNwPC += 'SCROLLBAR-FACE-COLOR: #DEDEDE; '
   cHtmNwPC += 'SCROLLBAR-HIGHLIGHT-COLOR: #DEDEDE; '
   cHtmNwPC += 'SCROLLBAR-SHADOW-COLOR: #ffffff; '
   cHtmNwPC += 'SCROLLBAR-3DLIGHT-COLOR: #ffffff; '
   cHtmNwPC += 'SCROLLBAR-ARROW-COLOR:  #ffffff; '
   cHtmNwPC += 'SCROLLBAR-TRACK-COLOR: #ffffff; '
   cHtmNwPC += 'SCROLLBAR-DARKSHADOW-COLOR: #DEDEDE; '
   cHtmNwPC += '}'
   cHtmNwPC += '.combo {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	margin: 1px;'
   cHtmNwPC += '	padding: 1px;'
   cHtmNwPC += '	border: 1px solid #000000;'
   cHtmNwPC += '}'
   cHtmNwPC += '.comboselect {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	color: #666666;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	border: 1px #CCCCCC double;'
   cHtmNwPC += '}'
   cHtmNwPC += '.links {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	color: #CC0000;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.links-clientes {'
   cHtmNwPC += 'font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: #003366; text-decoration: none'
   cHtmNwPC += '}'
   cHtmNwPC += '.textobold {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	color: #003366;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.textoItalico {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	color: 7F7F7F;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-style: italic;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.titulo {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 16px;'
   cHtmNwPC += '	color: #19167D;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.links-detalhes {'
   cHtmNwPC += 'font-family: Arial, Helvetica, sans-serif; font-size: 10px; color: FF0000; text-decoration: none'
   cHtmNwPC += '}'
   cHtmNwPC += '.TituloMenor {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 12px;'
   cHtmNwPC += '	color: #19167D;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.Botoes {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 10px;'
   cHtmNwPC += '	font-weight: normal;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	margin: 2px;'
   cHtmNwPC += '	padding: 2px;'
   cHtmNwPC += '	cursor: hand;'
   cHtmNwPC += '	border: 1px outset #000000;'
   cHtmNwPC += '}'
   cHtmNwPC += '.TarjaTopoMenu {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	height: 6px;'
   cHtmNwPC += '	background-image: url(file://///Srv-jbgarq02/Publico/Figuras Workflow/apmenu-right.jpg);'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoMenu {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-image: url(file://///Srv-jbgarq02/Publico/Figuras Workflow/apmenu-right.jpg);'
   cHtmNwPC += '}'
   cHtmNwPC += '.FonteMenu {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 12px;'
   cHtmNwPC += '	color: #19167D;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FonteSubMenu {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 12px;'
   cHtmNwPC += '	color: #19167D;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoSubMenu {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-image: url(file://///Srv-jbgarq02/Publico/Figuras Workflow/apmenu-right.jpg);'
   cHtmNwPC += '}'
   cHtmNwPC += '.DivisoriaOpçõesMenu {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #6680A6;'
   cHtmNwPC += '	background-image: url(file://///Srv-jbgarq02/Publico/Figuras Workflow/apmenu-right.jpg);'
   cHtmNwPC += '}'
   cHtmNwPC += '.TarjaTopoLogin {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #426285;'
   cHtmNwPC += '	height: 6px;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoLogin {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #F7F7F7;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoClaro {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #fbfbfb;'
   cHtmNwPC += '}'
   cHtmNwPC += '.TituloDestaques {'
   cHtmNwPC += '	font-family: Verdana, Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	color: #19167D;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.TextoDestaques {'
   cHtmNwPC += '	font-family: Verdana, Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 10px;'
   cHtmNwPC += '	color: #777777;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: normal;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoDestaques {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #E5E5E5;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoPontilhado {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-image: url(file://///Srv-jbgarq02/Publico/Figuras Workflow/pontilhado.gif);'
   cHtmNwPC += '	height: 5px;'
   cHtmNwPC += ' '
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoPontilhadoVertical {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-image: url(file://///Srv-jbgarq02/Publico/Figuras Workflow/pontilhado_vertical.gif);'
   cHtmNwPC += '	height: 5px;'
   cHtmNwPC += '}'
   cHtmNwPC += '.TituloTabelas {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 12px;'
   cHtmNwPC += '	color: #FFFFFF;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoTituloTabela {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #495E73;'
   cHtmNwPC += '}'
   cHtmNwPC += '.TarjaTopoCor {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	height: 6px;'
   cHtmNwPC += '	background-image: url(file://///Srv-jbgarq02/Publico/Figuras Workflow/pontilhado.gif);'
   cHtmNwPC += '	background-color: #6699CC;'
   cHtmNwPC += '}'
   cHtmNwPC += '.FundoTabelaDestaques {'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	background-color: #495E73;'
   cHtmNwPC += '}'
   cHtmNwPC += '.comboselect-pequeno {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	color: #666666;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	border: 1px solid #CCCCCC;'
   cHtmNwPC += '	width: 132px;'
   cHtmNwPC += '	clear: none;'
   cHtmNwPC += '	float: none;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	left: 1px;'
   cHtmNwPC += '	top: 1px;'
   cHtmNwPC += '	right: 1px;'
   cHtmNwPC += '	bottom: 1px;'
   cHtmNwPC += '	clip:  rect(1px 1px 1px 1px);'
   cHtmNwPC += '}'
   cHtmNwPC += '.comboselect-grande {'
   cHtmNwPC += '	font-family: Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	color: #666666;'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	border: 1px #CCCCCC double;'
   cHtmNwPC += '	width: 415px;'
   cHtmNwPC += '}'
   cHtmNwPC += '.texto-layer {'
   cHtmNwPC += 'font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; color: #000000; text-decoration: none'
   cHtmNwPC += '}'
   cHtmNwPC += '.tituloAvaliacao {'
   cHtmNwPC += '	font-family: Verdana, Arial, Helvetica, sans-serif;'
   cHtmNwPC += '	font-size: 9px;'
   cHtmNwPC += '	color: #19167D;'
   cHtmNwPC += '	text-decoration: none;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '	vertical-align: middle;'
   cHtmNwPC += '	text-align: center;'
   cHtmNwPC += '	line-height: 12px;'
   cHtmNwPC += '}'
   cHtmNwPC += '</style>'
   cHtmNwPC += '<link href="estilo.css" rel="stylesheet" type="text/css">'
   cHtmNwPC += '<style type="text/css"> '
   cHtmNwPC += '<!--'
   cHtmNwPC += '.style1 {font-size: 24px}'
   cHtmNwPC += '.style2 {font-size: 18px}'
   cHtmNwPC += '.style7 {font-size: 8px; }'
   cHtmNwPC += '.style9 {font-size: 36px; color: #0000FF; }'
   cHtmNwPC += '.style10 {'
   cHtmNwPC += '	font-size: 11px;'
   cHtmNwPC += '	font-weight: bold;'
   cHtmNwPC += '}'
   cHtmNwPC += '.style11 {font-size: 11px}'
   cHtmNwPC += '-->'
   cHtmNwPC += '</style>'
   cHtmNwPC += '</head>'
   cHtmNwPC += '<body>'
   cHtmNwPC += '<table width="98%" border="0" cellpadding="0" cellspacing="0">'
   cHtmNwPC += '  <tr>'
   cHtmNwPC += '    <td height="89" background="file://///Srv-jbgarq02/Publico/Figuras Workflow/topo_04.gif"><div align="center"><img src="file://///Srv-jbgarq02/Publico/Figuras Workflow/Logomarca_fundo_branco_1.jpg" width="113" height="76"></div></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '  <tr background="file://///Srv-jbgarq02/Publico/Figuras Workflow/pontilhado.gif">'
   cHtmNwPC += '    <td height="5" class="TarjaTopoCor"><img src="file://///Srv-jbgarq02/Publico/Figuras Workflow/transparente.gif" width="1" height="1"></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '</table>'
   cHtmNwPC += '<br>'
   cHtmNwPC += '<table width="96%" border="0" align="center" cellpadding="0" cellspacing="1">'
   cHtmNwPC += '  <tr>'
   cHtmNwPC += '    <td colspan="3"><div align="center"><span class="titulo"><span class="style9">%TAIMAUT%</span></span></div></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '  <tr>'
   cHtmNwPC += '    <td width="34%"><div align="justify"> <span class="titulo"><font size="3">&raquo; </font>Aprova&ccedil;&atilde;o de <span class="style1"> </span><span class="style1"> </span></span></div></td>'
   cHtmNwPC += '    <td width="30%"><div align="center"><span class="titulo"><span class="style1">Pedido de Compras</span> </span></div></td>'
   cHtmNwPC += '    <td width="36%"><div align="right"><span class="TituloMenor"><span class="textoItalico style2">No. %C7_NUM% - %C7_FILIAL%&nbsp;&nbsp;</span></span></div></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '</table>'
   cHtmNwPC += '<table width="96%" border="1" align="center" cellpadding="6" cellspacing="0" bordercolor="#E5E5E5" bgcolor="#F7F7F7" >'
   cHtmNwPC += '  <tr>'
   cHtmNwPC += '    <td width="100%" bordercolor="#FFFFFF" class="titulo"><table border="0" cellpadding="2" cellspacing="1" width="100%">'
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td width="12%" class="TituloMenor"><b>Emiss&atilde;o</b>: </td>'
   cHtmNwPC += '          <td width="38%" class="texto">%C7_EMISSAO%</td>'
   cHtmNwPC += '          <td width="12%" class="TituloMenor"><b>Total Bruto </b>: </td>'
   cHtmNwPC += '          <td width="38%" class="texto">%CR_TOTAL%</td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td width="12%" class="TituloMenor"><b>Comprador</b>:&nbsp; </td>'
   cHtmNwPC += '          <td width="38%" class="texto">%C7_USER%</td>'
   cHtmNwPC += '          <td width="12%" class="TituloMenor"><b>Desconto</b>: </td>'
   cHtmNwPC += '          <td width="38%" class="texto">%CR_DESCONTO%</td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td width="12%" class="TituloMenor"><b>Fornecedor</b>:&nbsp; </td>'
   cHtmNwPC += '          <td width="38%" class="texto">%A2_NOME%</td>'
   cHtmNwPC += '          <td width="12%" class="TituloMenor"><b>Total L&iacute;quido </b>: </td>'
   cHtmNwPC += '          <td width="38%" class="texto">%CR_LIQUIDO%</td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td class="TituloMenor">Tipo de Frete : </td>'
   cHtmNwPC += '          <td class="texto">%FRETE%</td>'
   cHtmNwPC += '          <td class="TituloMenor"><b>Cond.Pagamento</b>: </td>'
   cHtmNwPC += '          <td class="texto">%E4_DESCRI%</td>'
   cHtmNwPC += '        </tr>'
   
   
   /*
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td class="TituloMenor"><b>Aprovador</b>: </td>'
   cHtmNwPC += '          <td colspan="3" class="texto">%CR_USER%        </td>'
   cHtmNwPC += '        </tr>'
   
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td class="TituloMenor">Aplica&ccedil;&atilde;o:</td>
   cHtmNwPC += '          <td colspan="3" class="texto"><select'
   cHtmNwPC += '                              size="1" name="D3">'
   cHtmNwPC += '            <option selected>%Aplicacao%</option>'
   cHtmNwPC += '            &nbsp;'
   cHtmNwPC += '          </select></td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td class="TituloMenor">Urgência: </td>'
   cHtmNwPC += '          <td colspan="3" class="texto"><select'
   cHtmNwPC += '                              size="1" name="D4">'
   cHtmNwPC += '            <option selected>%Urgencia%</option>'
   cHtmNwPC += '            &nbsp;'
   cHtmNwPC += '          </select></td>'
   cHtmNwPC += '        </tr>'
   */
   
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td class="TituloMenor">&nbsp;</td>'
   cHtmNwPC += '          <td class="texto">&nbsp;</td>'
   cHtmNwPC += '          <td class="TituloMenor">&nbsp;</td>'
   cHtmNwPC += '          <td class="texto">&nbsp;</td>'
   cHtmNwPC += '        </tr>'
   
   cHtmNwPC += '      </table>'
   
*****************************************************************
   //Five Solutions - 06/06/2011
   CheckPrc()
   If lPrcMaior
      /*
      cHtmNwPC += '        <tr>'
      cHtmNwPC += '          <td class="TituloMenor">Neste Pedido existe produto com preco maior que o realizado na ultima compra</td>'
      cHtmNwPC += '        </tr>'
      */
      
      cHtmNwPC += ' <style type="text/css"> '
      cHtmNwPC += ' <!-- '
      //cHtmNwPC += ' .style1 {color: #FF0000} '
      cHtmNwPC += ' .style22 {color: #FF0000} '
      cHtmNwPC += ' --> '
      cHtmNwPC += ' </style> '
      cHtmNwPC += ' </head> '
      cHtmNwPC += ' <body> '
      
      //cHtmNwPC += ' <span class="style1">ATENCAO!!! Neste pedido existe(m) produto(s) com preco maior que o realizado na ultima compra</span> '
      cHtmNwPC += '        <tr>'
      cHtmNwPC += '          <td class="TituloMenor"><span class="style22">ATENCAO!!! Neste pedido existe(m) produto(s) com preco maior que o realizado na ultima compra</span></td>'
      cHtmNwPC += '        </tr>'
      
      cHtmNwPC += ' </body> '
      cHtmNwPC += ' </html> '
      
   EndIf

********************************************************************

   cHtmNwPC += '      <table border="0" cellspacing="1" width="100%">'
   cHtmNwPC += '        <tr bgcolor="#E4E4E4" class="TituloMenor">'
/*
   cHtmNwPC += '          <td width="5%" height="47"  align="center"><span class="style10">Item</span></td>'
   cHtmNwPC += '          <td width="6%" align="center">Ct</td>'
   cHtmNwPC += '          <td width="9%" align="center"><span class="style10">C&oacute;digo</span></td>'
   cHtmNwPC += '          <td width="14%" align="center" bgcolor="#E4E4E4"><span class="style10">Descri&ccedil;&atilde;o</span></td>'
   cHtmNwPC += '          <td width="5%" align="center"><span class="style10">Um</span></td>'
   cHtmNwPC += '          <td width="5%" align="center"><span class="style10">Quant</span></td>'
   cHtmNwPC += '          <td width="6%" align="center"><span class="style10">Vl.Unit</span></td>'
   cHtmNwPC += '          <td width="5%" align="center"><span class="style10">IPI</span></td>'
   cHtmNwPC += '          <td width="6%" align="center"><span class="style10">Vl.Total</span></td>'
   cHtmNwPC += '          <td width="7%" align="center"><span class="style11">Entrega</span></td>'
   cHtmNwPC += '          <td width="7%" align="center"><span class="style11">Pre&ccedil;o &Uacute;lt. Compra </span></td>'
   cHtmNwPC += '          <td width="7%" align="center"><p class="style11">Data da &Uacute;ltima Compra </p>          </td>'
   cHtmNwPC += '          <td width="6%" align="center"><span class="style11">&Uacute;lt. Fornec.</span></td>'
   cHtmNwPC += '          <td width="6%" align="center"><span class="style11">Centro de Custo</span></td>'
   cHtmNwPC += '          <td width="6%" align="center"><span class="style11"><b>Nro S.C </b></span></td>'
   */

**************************************************************************************************

   cHtmNwPC += '          <td width="10%" align="center"><b>C&oacute;digo</b></td>'
   cHtmNwPC += '          <td width="23%" align="center" bgcolor="#E4E4E4"><b>Descri&ccedil;&atilde;o</b></td>'
   cHtmNwPC += '          <td width="5%" align="center"><b>Um</b></td>'
   cHtmNwPC += '          <td width="9%" align="center"><b>Quant</b></td>'
   cHtmNwPC += '          <td width="8%" align="center"><b>Vl.Unit&aacute;rio</b></td>'
   cHtmNwPC += '          <td width="10%" align="center"><b>Vl.Total</b></td>'
   cHtmNwPC += '          <td width="8%" align="center">Natureza</td>'
   cHtmNwPC += '          <td width="27%" align="center">Descrição</td>'
   cHtmNwPC += '          <td width="27%" align="center">Ultimo Preco</td>'//Five Solutions - 06/06/2011

************************************************************************************************   
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tbody>'
   cHtmNwPC += '          <tr bgcolor="#FFFFFF" class="texto-layer">'
/*
   cHtmNwPC += '            <td height="1"><div align="center" class="style7">%t1.1%</div></td>'
   cHtmNwPC += '            <td><div align="center"><span class="style7">%t1.15%</span></div></td>'
   cHtmNwPC += '            <td height="1"><div align="center" class="style7">%t1.2%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center" class="style7">%t1.3%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center" class="style7">%t1.4%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center" class="style7">%t1.5%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center" class="style7">%t1.6%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center" class="style7">%t1.7%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center" class="style7">%t1.8%</div></td>'
   cHtmNwPC += '            <td><div align="center" class="style7">%t1.9%</div></td>'
   cHtmNwPC += '            <td><div align="center" class="style7">%t1.10%</div></td>'
   cHtmNwPC += '            <td><div align="center" class="style7">%t1.11%</div></td>'
   cHtmNwPC += '            <td><div align="center"><span class="style7">%t1.12%</span></div></td>'
   cHtmNwPC += '            <td><div align="center" class="style7">%t1.13%</div>              </td>'
   cHtmNwPC += '            <td height="1"><div align="center" class="style7">%t1.14%</div></td>'
   */

***************************************************************************************************************   
   cHtmNwPC += '            <td height="1"><div align="center">%t1.1%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.2%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.3%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.4%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.5%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.6%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.7%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.8%</div></td>'
   cHtmNwPC += '            <td height="1"><div align="center">%t1.9%</div></td>' //Five Solutions - 06/06/2011
   //cHtmNwPC += '            <td>%t1.9%</td>'
   //cHtmNwPC += '            <td>%t1.10%</td>'
   //cHtmNwPC += '            <td>%t1.11%</td>'
***************************************************************************************************************
   
   
   cHtmNwPC += '          </tr>'
   cHtmNwPC += '          <tr>'
   cHtmNwPC += '            <td colspan="15" class="texto-layer" background="file://///Srv-jbgarq02/Publico/Figuras Workflow/pontilhado.gif"><span class="TituloMenor"><img src="file://///Srv-jbgarq02/Publico/Figuras Workflow/transparente.gif" width="10" height="3"></span></td>'
   cHtmNwPC += '          </tr>'
   cHtmNwPC += '        </tbody>'
   cHtmNwPC += '      </table>'
   
   cHtmNwPC += '      <table border="0" width="100%">'
   cHtmNwPC += '        <tr height="21">'
   cHtmNwPC += '          <td colspan="2">&nbsp;</td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '<form method="post" action="mailto:%WFMailTo%" name="FrontPage_Form1">        '
   cHtmNwPC += '	<tr height="21">'
   cHtmNwPC += '          <td width="14%"><p align="center" class="TituloMenor">APROVA&Ccedil;&Atilde;O:</p></td>'
   cHtmNwPC += '          <td width="86%">            <input type="radio" checked name="OPC" value="%OPC%S">'
   cHtmNwPC += '            <span class="texto">Sim.</span></td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <tr height="21">'
   cHtmNwPC += '          <td width="14%"></td>'
   cHtmNwPC += '          <td width="86%">            <input type="radio" name="OPC" value="%OPC%N">'
   cHtmNwPC += '            <span class="texto">N&atilde;o. (Especificar motivo) </span></td>'
   cHtmNwPC += '        </tr>'
   /*
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td width="14%" height="100"></td>'
   cHtmNwPC += '          <td width="86%" height="100">'
   cHtmNwPC += '          <textarea name="OBS" cols="33" rows="3" wrap="PHYSICAL"'
   cHtmNwPC += '                          >%OBS%</textarea>           </td>'
   cHtmNwPC += '        </tr>'
   */
   // TRECHO OBSERVAÇÃO MODIFICADO
   cHtmNwPC += '        <tr>'
   cHtmNwPC += '          <td width="14%" height="100"></td>'
   cHtmNwPC += '          <td width="86%" height="100">'
   //cHtmNwPC += '          <textarea name="OBS" cols="61" rows="4" wrap="PHYSICAL" class="style5"'
   cHtmNwPC += '          <textarea name="OBS" cols="100" rows="4" wrap="PHYSICAL" class="style5"'
   cHtmNwPC += '                          >%OBS%</textarea>           </td>'
   cHtmNwPC += '        </tr>'
   
   cHtmNwPC += '        <tr height="27">'
   cHtmNwPC += '          <td width="14%"></td>'
   cHtmNwPC += '          <td width="86%"> <input  name="B1" type="submit" class="Botoes" value="Enviar"> <input name="B2" type="reset" class="Botoes"  value="Limpar">  </td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '        <input type="hidden" name="CHAVE" 		value="%CHAVE%">'
   cHtmNwPC += '        <input type="hidden" name="CFILANT" 	value="%CFILANT%">'
   
   cHtmNwPC += '        <input type="hidden" name="CEMPANT" 	value="%CEMPANT%">'
   
   cHtmNwPC += '        <input type="hidden" name="WF" 		value="%WF%">'
   cHtmNwPC += '        <input type="hidden" name="WFID" 		value="%WFID%">'
   cHtmNwPC += '      </table>'
   cHtmNwPC += '    </td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '</table>'

   /*
   cHtmNwPC += '<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1">'
   cHtmNwPC += '  <tr>'
   cHtmNwPC += '    <td><span class="titulo"><font size="3">&raquo; </font>Cota&ccedil;&otilde;es Efetuadas para Este Pedido </span></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '</table>'
   cHtmNwPC += '<table width="933" height="11%" border="0" align="center" cellspacing="1">'
   cHtmNwPC += '  <tr bgcolor="#E4E4E4" class="TituloMenor">'
   cHtmNwPC += '    <td width="63" height="54%" align="center"><div align="center"><b>Cota&ccedil;&atilde;o</b></div></td>'
   cHtmNwPC += '    <td width="72" align="center"><div align="center"><b>Fornecedor</b></div></td>'
   cHtmNwPC += '    <td width="46" align="center"><div align="center"><b> Loja</b></div></td>'
   cHtmNwPC += '    <td width="288" align="center">Nome</td>'
   cHtmNwPC += '    <td width="330" align="center"><div align="center"><b>Produto</b></div></td>'
   cHtmNwPC += '    <td width="115" align="center"><div align="center"><b>Pre&ccedil;o</b></div></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '  <tbody>'
   cHtmNwPC += '    <tr bgcolor="#FFFFFF" class="texto-layer">'
   cHtmNwPC += '      <td height="46%"><div align="center">%t2.1%</div></td>'
   cHtmNwPC += '      <td><div align="center">%t2.2%</div></td>'
   cHtmNwPC += '      <td><div align="center">%t2.3%</div></td>'
   cHtmNwPC += '      <td><div align="center">%t2.4%</div></td>'
   cHtmNwPC += '      <td><div align="center">%t2.5%</div></td>'
   cHtmNwPC += '      <td><div align="center">%t2.6%</div></td>'
   cHtmNwPC += '    </tr>'
   cHtmNwPC += '  </tbody>'
   cHtmNwPC += '</table>'

   cHtmNwPC += '<div align="right"></div>'
   cHtmNwPC += '<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1">'
   cHtmNwPC += '  <tr>'
   cHtmNwPC += '    <td><span class="titulo"><font size="3">&raquo; </font>Processo de Aprova&ccedil;&atilde;o do Pedido de Compra </span></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '</table>'
   cHtmNwPC += '<table width="96%" border="1" align="center" cellpadding="6" cellspacing="0" bordercolor="#E5E5E5" bgcolor="#F7F7F7" >'
   cHtmNwPC += '  <tr>'
   cHtmNwPC += '    <td width="100%" bordercolor="#FFFFFF" class="titulo"><table width="100%" border="0" cellspacing="1" height="42">'
   cHtmNwPC += '      <tr bgcolor="#E4E4E4" class="TituloMenor">'
   cHtmNwPC += '        <td width="4%" height="17" align="center"><b>N&iacute;vel</b></td>'
   cHtmNwPC += '        <td width="12%" height="17" align="center"><b>Usu&aacute;rio</b></td>'
   cHtmNwPC += '        <td width="11%" height="17" align="center"><b>Situa&ccedil;&atilde;o</b></td>'
   cHtmNwPC += '        <td width="14%" height="17" align="center"><b>Aprovado por</b></td>'
   cHtmNwPC += '        <td width="7%" height="17" align="center"><b>Data da Libera&ccedil;&atilde;o</b></td>'
   cHtmNwPC += '        <td width="52%" height="17" align="center"><b>Observa&ccedil;&otilde;es</b></td>'
   cHtmNwPC += '      </tr>'
   cHtmNwPC += '      <tbody>'
   cHtmNwPC += '        <tr bgcolor="#FFFFFF" class="texto-layer">'
   cHtmNwPC += '          <td height="1"><div align="center">%t.1%</div></td>'
   cHtmNwPC += '          <td height="1"><div align="center">%t.2%</div></td>'
   cHtmNwPC += '          <td height="1"><div align="center">%t.3%</div></td>'
   cHtmNwPC += '          <td height="1"><div align="center">%t.4%</div></td>'
   cHtmNwPC += '          <td height="1"><div align="center">%t.5%</div></td>'
   cHtmNwPC += '          <td height="1"><div align="center">%t.6%</div></td>'
   cHtmNwPC += '        </tr>'
   cHtmNwPC += '      </tbody>'
   cHtmNwPC += '    </table></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '</table>'
   
   */
   
   cHtmNwPC += '<table width="96%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" bgcolor="#F7F7F7" >'
   cHtmNwPC += '   <tr>'
   cHtmNwPC += '    <td width="100%" bordercolor="#FFFFFF"><div align="right" class="texto-layer">WorkFlow @ Microsiga Protheus</div></td>'
   cHtmNwPC += '  </tr>'
   cHtmNwPC += '</table>'
   //cHtmNwPC += '<p><span class="TituloMenor"><b>Legenda: &quot;Ct&quot; </b> = ( S ) O produto possui contrato de fornecimento vigente - ( N ) N&atilde;o possui contrato de fornecimento vigente</span></p>'
   cHtmNwPC += '</body>'
   cHtmNwPC += '</html>'

   MemoWrite("\Workflow\PCCOMAFAL.htm",cHtmNwPC)
   
   ConOut("Gerou NOVO - html em \Workflow\PCCOMAFAL.htm")
   
   //Five Solutions Consultoria - 14/05/2010 - Após gerar o Html roda a macro-substituições dos campos.
   
   For nPC := 1 To 2
   
      If nPC == 1
         cQryPC  := " SELECT COUNT(*) NQTDREG,SUM(SC7.C7_VALIPI) VLTOTIPI,SUM(SC7.C7_VALICM) VLRTICM,"
         cQryPC  += "        SUM(SC7.C7_TOTAL) NTOTMERC,SUM(SC7.C7_VLDESC) NDESCTOT"
      Else   
         cQryPC  := " SELECT SC7.C7_NUM,SC7.C7_EMISSAO,SC7.C7_FORNECE,SC7.C7_LOJA,SA2.A2_NOME,SC7.C7_COND,SA2.A2_NREDUZ,"
         cQryPC  += "        SC7.C7_ITEM,SC7.C7_PRODUTO,SB1.B1_DESC,SC7.C7_QUANT,SC7.C7_UM,SC7.C7_PRECO,SC7.C7_TOTAL,"
         cQryPC  += "        SC7.C7_DATPRF,SA2.A2_COD,SA2.A2_LOJA,SA2.A2_INSCR,SA2.A2_CGC,SA2.A2_CONTATO,SA2.A2_TEL,"
         cQryPC  += "        SA2.A2_FAX,SA2.A2_END,SA2.A2_BAIRRO,SA2.A2_MUN,SA2.A2_CEP,SA2.A2_CGC,SC7.C7_FRETE,"         
         cQryPC  += "        SC7.C7_SEGURO,SC7.C7_DESPESA,SC7.C7_VALFRE,SC7.C7_FILIAL,SC7.C7_USER,SE4.E4_DESCRI,"
         cQryPC  += "        SC7.C7_OBS,SA2.A2_EST,SC7.C7_DESCRI,SC7.C7_IPI,SC7.C7_VALIPI,SC7.C7_CC,SC7.C7_NUMCOT,"
         cQryPC  += "        SC7.C7_VLDESC,SC7.C7_VALICM,SC7.C7_X_NATUR,SED.ED_DESCRIC"
      EndIf
   
      cQryPC  += "   FROM "+RetSQLName("SC7")+" SC7,"+RetSQLName("SA2")+" SA2,"+RetSQLName("SB1")+" SB1,"
      cQryPC  += RetSQLName("SE4")+" SE4,"+RetSQLName("SED")+" SED"
      cQryPC  += "  WHERE SC7.C7_FILIAL = '"+xFilial("SC7")+"'"
      cQryPC  += "    AND SA2.A2_FILIAL = '"+xFilial("SA2")+"'"
      cQryPC  += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"
      cQryPC  += "    AND SE4.E4_FILIAL = '"+xFilial("SE4")+"'"
      
      cQryPC  += "    AND SED.ED_FILIAL = '"+xFilial("SED")+"'"
      cQryPC  += "    AND SC7.C7_X_NATUR = SED.ED_CODIGO"
         
      cQryPC  += "    AND SC7.C7_FORNECE = SA2.A2_COD"
      cQryPC  += "    AND SC7.C7_LOJA = SA2.A2_LOJA"
      cQryPC  += "    AND SC7.C7_PRODUTO = SB1.B1_COD"
      
      cQryPC  += "    AND SC7.C7_COND = SE4.E4_CODIGO"
   
      cQryPC  += "    AND SC7.C7_NUM = '"+cNumPC+"'"
      cQryPC  += "    AND SC7.C7_FORNECE = '"+cForPC+"'"
      cQryPC  += "    AND SC7.C7_LOJA = '"+cLjPC+"'"
   
      cQryPC  += "    AND SC7.D_E_L_E_T_ <> '*'"
      cQryPC  += "    AND SA2.D_E_L_E_T_ <> '*'"
      cQryPC  += "    AND SB1.D_E_L_E_T_ <> '*'"
      cQryPC  += "    AND SE4.D_E_L_E_T_ <> '*'"
      
      cQryPC  += "    AND SED.D_E_L_E_T_ <> '*'"
   
      If nPC > 1
         cQryPC  += " ORDER BY SC7.C7_ITEM"
      EndIf
      
      MemoWrite("PedidodeCompras.SQL",cQryPC)
   
      TCQuery cQryPC NEW ALIAS "TXC7"
      
      If nPC == 1
         DbSelectArea("TXC7")
         TCSetField("TXC7","VLTOTIPI","N",17,02)
         TCSetField("TXC7","VLRTICM","N",17,02)
         TCSetField("TXC7","NTOTMERC","N",17,02)
         TCSetField("TXC7","NDESCTOT","N",17,02)

         nReg      := TXC7->NQTDREG
         nVlrTtIPI := TXC7->VLTOTIPI
         nICMTtVl  := TXC7->VLRTICM
         nVlrMercad:= TXC7->NTOTMERC
         TXC7->(DbCloseArea())
      EndIf
   
   Next nPC
   
   TCSetField("TXC7","C7_QUANT","N",12,03)
   TCSetField("TXC7","C7_PRECO","N",14,02)
   TCSetField("TXC7","C7_TOTAL","N",14,02)
   TCSetField("TXC7","C7_VALIPI","N",14,02)
   TCSetField("TXC7","C7_SEGURO","N",14,02)
   TCSetField("TXC7","C7_DESPESA","N",14,02)
   TCSetField("TXC7","C7_VALFRE","N",14,02)
   TCSetField("TXC7","C7_DATPRF","D",08,00)
   TCSetField("TXC7","C7_EMISSAO","D",08,00)
   
   DbSelectArea("TXC7")

   oProcess:NewTask("Gerando E-mail da Aprovacao de Compras", "\Workflow\PCCOMAFAL.htm" ) 

   //oHTML   := oProcess:oHTML
   //oHtml:ValByName( "TIPOPED"	, "Pedido de Compras" )
   oProcess:oHTML:ValByName( "TAIMAUT"	, Alltrim(SM0->M0_NOME)+"/"+Alltrim(SM0->M0_FILIAL) )

   _cChave  := TXC7->(C7_FILIAL+C7_NUM) //xFilial("SC7")+cNumPC //SC7->(C7_FILIAL+C7_NUM)
   _cWF     := SPACE(2)
	//			_cSoli   := SC7->C7_USER
	//			_nRegSC7 := Recno()
				
   // Hidden Fields
   oProcess:oHTML:ValByName( "CHAVE"	    , _cChave)
   oProcess:oHTML:ValByName( "CFILANT"	    , cFilAnt)
   
   oProcess:oHTML:ValByName( "CEMPANT"	    , cEmpAnt) 
   
   oProcess:oHTML:ValByName( "WF"   		, Trim(_cWF))
   oProcess:oHTML:ValByName( "WFID"			, oProcess:fProcessId)
   
   oProcess:ohtml:valbyname('C7_NUM', cNumPC )   
   oProcess:ohtml:valbyname('C7_FILIAL', TXC7->C7_FILIAL )   

   //Cabeçalho - Dados do Fornecedor:
   oProcess:oHTML:ValByName( "A2_NOME"		, Trim(TXC7->A2_NOME) )

   // Dados do Pedido:
   oProcess:oHTML:ValByName( "C7_EMISSAO"   , TXC7->C7_EMISSAO )
   oProcess:oHTML:ValByName( "E4_DESCRI"    , TXC7->E4_DESCRI )
   oProcess:oHTML:ValByName( "C7_USER"      , UsrFullName(TXC7->C7_USER) )
   
   _nTotPed   := 0
   _nTotImp   := 0
   _nTotIPI   := 0
   _nTotICMS  := 0
   _nTotFrete := 0
   _nTotDesp  := 0
   _nTotSeg   := 0
   _nTotDesc  := 0
   cObs       := ""
   lUmaVez    := .T.

   DbSelectArea("TXC7")
   ProcRegua(nReg)
   nCont := 1
   While TXC7->(!Eof())
      
      IncProc("Gerando Itens p/Aprovação ... "+Alltrim(Str(nCont))+" / "+Alltrim(Str(nReg)))
	 
      
      //Five Solutions - 06/06/2011 - Implementado para apresentar último preço de compra do Produto.
      nUPrcComp := 0
      DbSelectArea("SBZ") //Indicadores de Produtos
      DbSetOrder(1) //BZ_FILIAL+BZ_COD
      If DbSeek(xFilial("SBZ")+TXC7->C7_PRODUTO)
         nUPrcComp := SBZ->BZ_UPRC
         If nUPrcComp == 0
            DbSelectArea("SB1") //Cadastro Genérico de Produtos
            DbSetOrder(1) //B1_FILIAL+B1_COD
            If DbSeek(xFilial("SB1")+TXC7->C7_PRODUTO)
               nUPrcComp := SB1->B1_UPRC
            EndIf
         EndIf
      EndIf
      /*
	  If TXC7->C7_PRECO > nUPrcComp .And. !lPrcMaior
	     lPrcMaior := .T.
	  EndIf
	  */
      DbSelectArea("TXC7")
      
	  AAdd( (oProcess:oHTML:ValByName( "t1.1"  )), TXC7->C7_PRODUTO)
	  AAdd( (oProcess:oHTML:ValByName( "t1.2"  )), TXC7->C7_DESCRI)
	  AAdd( (oProcess:oHTML:ValByName( "t1.3"  )), TXC7->C7_UM)
	  AAdd( (oProcess:oHTML:ValByName( "t1.4"  )), TRANSFORM(TXC7->C7_QUANT,'@E 999,999.999'))
	  AAdd( (oProcess:oHTML:ValByName( "t1.5"  )), TRANSFORM(TXC7->C7_PRECO,'@E 999,999.999'))
	  AAdd( (oProcess:oHTML:ValByName( "t1.6"  )), TRANSFORM(TXC7->C7_QUANT * TXC7->C7_PRECO,'@E 999,999.99'))
	  AAdd( (oProcess:oHTML:ValByName( "t1.7"  )), TXC7->C7_X_NATUR)
	  AAdd( (oProcess:oHTML:ValByName( "t1.8"  )), TXC7->ED_DESCRIC)
	  AAdd( (oProcess:oHTML:ValByName( "t1.9"  )), TRANSFORM(nUPrcComp,'@E 999,999.999'))//Five Solutions - 06/06/2011

	  _nTotPed   += TXC7->C7_TOTAL
	  _nTotDesc  += TXC7->C7_VLDESC
	  _nTotImp   += ( TXC7->C7_TOTAL+TXC7->C7_VALIPI+TXC7->C7_SEGURO+TXC7->C7_DESPESA+TXC7->C7_VALFRE ) - TXC7->C7_VLDESC
	  _nTotIPI   += TXC7->C7_VALIPI
	  _nTotICMS  += TXC7->C7_VALICM
	  _nTotFrete += TXC7->C7_VALFRE
	  _nTotDesp  += TXC7->C7_DESPESA
	  _nTotSeg   += TXC7->C7_SEGURO
	
	  cTmpObs := Alltrim(TXC7->C7_OBS)
	  If cTmpObs <> Alltrim(Substr(cObs,4,150))
	     If lUmaVez
	         lUmaVez := .F.
	         cObs := cTmpObs
	     Else
	        cObs += ' - '+cTmpObs
	     EndIf
      EndIf

      DbSelectArea("TXC7")
      DbSkip()
      nCont ++
      
   EndDo
   DbSelectArea("TXC7")
   TXC7->(DbCloseArea())

   oProcess:oHTML:ValByName( "OBS"		   , cObs )
   oProcess:oHTML:ValByName( "CR_TOTAL"       , TRANSFORM(_nTotImp,'@E 999,999.99') )   
   oProcess:oHTML:ValByName( "CR_DESCONTO"      , TRANSFORM(_nTotDesc,'@E 999,999.99') )
   oProcess:oHTML:ValByName( "CR_LIQUIDO"      , TRANSFORM(_nTotPed,'@E 999,999.99') )
Return(oProcess)

Static Function CheckPrc()
   

      cQryPC  := " SELECT SC7.C7_NUM,SC7.C7_EMISSAO,SC7.C7_FORNECE,SC7.C7_LOJA,SA2.A2_NOME,SC7.C7_COND,SA2.A2_NREDUZ,"
      cQryPC  += "        SC7.C7_ITEM,SC7.C7_PRODUTO,SB1.B1_DESC,SC7.C7_QUANT,SC7.C7_UM,SC7.C7_PRECO,SC7.C7_TOTAL,"
      cQryPC  += "        SC7.C7_DATPRF,SA2.A2_COD,SA2.A2_LOJA,SA2.A2_INSCR,SA2.A2_CGC,SA2.A2_CONTATO,SA2.A2_TEL,"
      cQryPC  += "        SA2.A2_FAX,SA2.A2_END,SA2.A2_BAIRRO,SA2.A2_MUN,SA2.A2_CEP,SA2.A2_CGC,SC7.C7_FRETE,"         
      cQryPC  += "        SC7.C7_SEGURO,SC7.C7_DESPESA,SC7.C7_VALFRE,SC7.C7_FILIAL,SC7.C7_USER,SE4.E4_DESCRI,"
      cQryPC  += "        SC7.C7_OBS,SA2.A2_EST,SC7.C7_DESCRI,SC7.C7_IPI,SC7.C7_VALIPI,SC7.C7_CC,SC7.C7_NUMCOT,"
      cQryPC  += "        SC7.C7_VLDESC,SC7.C7_VALICM,SC7.C7_X_NATUR,SED.ED_DESCRIC"
   
      cQryPC  += "   FROM "+RetSQLName("SC7")+" SC7,"+RetSQLName("SA2")+" SA2,"+RetSQLName("SB1")+" SB1,"
      cQryPC  += RetSQLName("SE4")+" SE4,"+RetSQLName("SED")+" SED"
      cQryPC  += "  WHERE SC7.C7_FILIAL = '"+xFilial("SC7")+"'"
      cQryPC  += "    AND SA2.A2_FILIAL = '"+xFilial("SA2")+"'"
      cQryPC  += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"
      cQryPC  += "    AND SE4.E4_FILIAL = '"+xFilial("SE4")+"'"
      
      cQryPC  += "    AND SED.ED_FILIAL = '"+xFilial("SED")+"'"
      cQryPC  += "    AND SC7.C7_X_NATUR = SED.ED_CODIGO"
         
      cQryPC  += "    AND SC7.C7_FORNECE = SA2.A2_COD"
      cQryPC  += "    AND SC7.C7_LOJA = SA2.A2_LOJA"
      cQryPC  += "    AND SC7.C7_PRODUTO = SB1.B1_COD"
      
      cQryPC  += "    AND SC7.C7_COND = SE4.E4_CODIGO"
   
      cQryPC  += "    AND SC7.C7_NUM = '"+cNumPC+"'"
      cQryPC  += "    AND SC7.C7_FORNECE = '"+cForPC+"'"
      cQryPC  += "    AND SC7.C7_LOJA = '"+cLjPC+"'"
   
      cQryPC  += "    AND SC7.D_E_L_E_T_ <> '*'"
      cQryPC  += "    AND SA2.D_E_L_E_T_ <> '*'"
      cQryPC  += "    AND SB1.D_E_L_E_T_ <> '*'"
      cQryPC  += "    AND SE4.D_E_L_E_T_ <> '*'"
      
      cQryPC  += "    AND SED.D_E_L_E_T_ <> '*'"
   
      
      cQryPC  += " ORDER BY SC7.C7_ITEM"
      
      
      MemoWrite("PedidodeCompras.SQL",cQryPC)
   
      TCQuery cQryPC NEW ALIAS "TXC7"
      
 
   TCSetField("TXC7","C7_QUANT","N",12,03)
   TCSetField("TXC7","C7_PRECO","N",14,02)
   TCSetField("TXC7","C7_TOTAL","N",14,02)
   TCSetField("TXC7","C7_VALIPI","N",14,02)
   TCSetField("TXC7","C7_SEGURO","N",14,02)
   TCSetField("TXC7","C7_DESPESA","N",14,02)
   TCSetField("TXC7","C7_VALFRE","N",14,02)
   TCSetField("TXC7","C7_DATPRF","D",08,00)
   TCSetField("TXC7","C7_EMISSAO","D",08,00)
   
   DbSelectArea("TXC7")
   While TXC7->(!Eof())
      
      //Five Solutions - 06/06/2011 - Implementado para apresentar último preço de compra do Produto.
      nUlPrcCp := 0
      DbSelectArea("SBZ") //Indicadores de Produtos
      DbSetOrder(1) //BZ_FILIAL+BZ_COD
      If DbSeek(xFilial("SBZ")+TXC7->C7_PRODUTO)
         nUlPrcCp := SBZ->BZ_UPRC
         If nUlPrcCp == 0
            DbSelectArea("SB1") //Cadastro Genérico de Produtos
            DbSetOrder(1) //B1_FILIAL+B1_COD
            If DbSeek(xFilial("SB1")+TXC7->C7_PRODUTO)
               nUlPrcCp := SB1->B1_UPRC
            EndIf
         EndIf
      EndIf
      
	  If TXC7->C7_PRECO > nUlPrcCp .And. !lPrcMaior
	     lPrcMaior := .T.
	  EndIf
	  
      DbSelectArea("TXC7")
      DbSkip()
      
   EndDo   
   DbSelectArea("TXC7")
   TXC7->(DbCloseArea())
Return