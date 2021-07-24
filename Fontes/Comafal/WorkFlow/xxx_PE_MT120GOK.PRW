#INCLUDE "PROTHEUS.CH"
#include "TopConn.Ch"
#INCLUDE "TBICONN.CH "
#include "TbiCode.ch"
#INCLUDE "Ap5mail.ch"




/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³PE_MT120GOKºAutor  ³Five Solutions      º Data ³  09/04/2010 º±±
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

User Function MT120GOK
/*
pcAssunto := "TESTE E-MAIL WF"
pcUsuario := "itacolomy@fivesolutions.com.br"
pcTitulo := "titulo do e-mail DE TESTE WF"
pcDetalhe := "DETALHE DA MENSAGEM DE TESTE WF"
pcAtachar := ""

oProcess          	:= TWFProcess():New( "000001", "Notificacao Avulsa" )
//oProcess          	:NewTask( "Notificacao Avulsa", "\workflow\html\Notificacoes.htm" )
oProcess          	:NewTask( "Notificacao Avulsa", "\workflow\Notificacoes.htm" )
oProcess:cSubject 	:= pcAssunto
If '@' $ pcUsuario
	oProcess:cTo   	:= pcUsuario
Else
	oProcess:cTo    := UsrRetMail(pcUsuario)
EndIf
oProcess:UserSiga	:= '000000'  // Fixo Administrador
oProcess:NewVersion(.T.)
oHtml     			:= oProcess:oHTML
oHtml:ValByName( "REFERENTE", pcTitulo)
oHtml:ValByName( "DESCRICAO", pcDetalhe)
If pcAtachar # Nil    
	oProcess:AttachFile(Upper(pcAtachar))
EndIf
oProcess:Start()

Return
*/

/*   
   Local   nDias := 0, nHoras := 0, nMinutos := 10
   Private cUsuarioProtheus := SubStr(cUsuario,7,15)   // Registre o nome do usuário corrente que está criando o processo:
   Private cCodProcesso := "000001" //"PEDCOM"
   Private cAssunto     := "Aprovacao do Pedido de Compras" 
   Private cTitulo      :=  "Criando Aprovacao"
   //Private cHtmlAC      := "C:\Protheus10\Workflow\ApPedCom.htm"
   Private cHtmlAC      := "\Workflow\ApPedCom.htm"
   
   Private cNumPC       := SC7->C7_NUM
   Private cTo          := "itacolomy@fivesolutions.com.br"
   Private cCC          := "itacolomy@hotmail.com"
   Private cBCC         := "suporte_siga@yahoo.com.br"

   //oProcess := TWFProcess():New(cCodProcesso, cAssunto)
   Processa({ ||GHtmWFPC()},"Gerando Html do Pedido "+cNumPC+" Aguarde ...")
   oProcess := TWFProcess():New(cCodProcesso, cAssunto) // Inicializa a classe TWFProcess e assinale a variável objeto oProcess:
   oProcess:NewTask(cTitulo, cHtmlAC) // Cria uma tarefa.
   oProcess:cSubject := cAssunto   
   oProcess:cTo  // Informa o endereço eletrônico do destinatário.
   oProcess:cCC  // Informa o endereço eletrônico para cópia.
   oProcess:cBCC // Informa o endereço eletrônico para cópia oculta.
   oProcess:cBody := "Testando..."
   oProcess:oHtml:ValByName("aprovacao"     , "Sim"  )
   oHtml := oProcess:oHTML
   oProcess:UserSiga := WFCodUser("BI")
   oProcess:AttachFile("\Workflow\ApPedCom.htm")   
   oProcess:Start()
   WFSendMail({"99","01"})
   cTexto := "Enviado " + cAssunto + " Número: " + cNumPC
   cCodStatus := "100100" // Código do cadastro de status de processo.
   oProcess:Track(cCodStatus, cTexto, cUsuarioProtheus) // Rastreabilidade
Return
*/
      

   Local   nDias := 0, nHoras := 0, nMinutos := 10
   Private cUsuarioProtheus := SubStr(cUsuario,7,15)   // Registre o nome do usuário corrente que está criando o processo:
   Private cCodProcesso := "000001" //"PEDCOM"
   Private cAssunto     := "Aprovacao do Pedido de Compras" 
   Private cTitulo      :=  "Criando Aprovacao"
   //Private cHtmlAC      := "C:\Protheus10\Workflow\ApPedCom.htm"
   Private cHtmlAC      := "\Workflow\ApPedCom.htm"
   Private cNumPC       := SC7->C7_NUM
   Private cTo          := "itacolomy@fivesolutions.com.br"
   Private cCC          := "itacolomy@hotmail.com"
   Private cBCC         := "suporte_siga@yahoo.com.br"
Processa({ ||GHtmWFPC()},"Gerando Html do Pedido "+cNumPC+" Aguarde ...")
_cPara := "itacolomy@fivesolutions.com.br"
cCopia := "itacolomy@hotmail.com"
_cAssunto := "Teste Email Workflow por outra funcionalidade"
//_cMsg := "Este email destina-se a teste para funcionamento de acesso/uso da conta workflow@fivesolutions.com.br para transmissoes de email dos processos Workflow"
_cMsg := "\Workflow\ApPedCom.htm"
_cAnexo := ""
_cServer := "smtp.jvfi.com.br"
//_cServer := "smtp.fivesolutions.com.br" 
//_cConta := "workflow@fivesolutions.com.br"
//_cConta := "itacolomy@fivesolutions.com.br"
_cConta := "jvfi@jvfi.com.br"
//_cSenha := "workflow"
//_cSenha := "five_itacolomy"
_cSenha := "13112000"

   CONNECT SMTP SERVER _cServer ACCOUNT _cConta PASSWORD _cSenha RESULT lConectou

   If !lConectou
      ConOut("O programa nao conseguiu se conectar ao servidor de e-mail."+_cServer+" não será possível enviar mail") 
      Return(.F.)
   Endif

   lAutent := mailauth(_cConta, _cSenha)
   
   If !lAutent
      GET MAIL ERROR cMenErr
      ConOut("Não conseguiu autenticação")
      Return(.F.)
   Endif
	if lConectou
		// MailSend está dentro do AP5MAIL.CH
		cFrom := ""
		aCopia := {"itacolomy@hotmail.com"}
		aOculto:= {}
		cAssunto := _cAssunto
		cCorpo :=  "TESTE DE CORPO"
		aAnexos := {}
		data lMsgOut	//exibe msg no console do servidor?
		::lMsgOut	:= .t.	//exibe msg no console do servidor?
		
		lConectou := MailSend( if(empty(cFrom),::_cConta,cFrom), aTo , if(len(aCopia)>0,aCopia,"") , if(len(aOculto)>0,aOculto,"") , cAssunto, cCorpo, aAnexos, .t. )

		if ::lMsgOut
			if lConectou
				Conout( "Enviado com sucesso!!" )
			else
				Conout( "Ocorreu um erro no envio da mensagem." )  
				Conout(MailGetErr())
			end
		end
		DISCONNECT SMTP SERVER
    else
    	if ::lMsgOut
			Conout( "Ocorreu um erro na conexao com o servidor SMTP" )  
			Conout(MailGetErr())
		end
    end
   /*
   if Empty(_cAnexo)
      SEND MAIL FROM _cConta   ;
      TO _cPara  ; 
      CC cCopia  ;
      SUBJECT _cAssunto ;
      BODY _cMsg       ;
      RESULT lEnviado
   else
      SEND MAIL FROM _cConta   ;
      TO _cPara  ;
      CC cCopia  ;      
      SUBJECT _cAssunto ;
      BODY _cMsg       ;
      ATTACHMENT _cAnexo ;
      RESULT lEnviado
   endif

   If !lEnviado
      GET MAIL ERROR cMenErr
      Return(.F.)
   Endif

   DISCONNECT SMTP SERVER RESULT lDisconectou
   */
Return(.T.)   

Static Function tst11
*****************************************************************************************************************/
  
   /*************************************
   * Five Solutions Consultoria
   * 09/04/2010
   *   Classe TWFProcess
   *   A classe TWFProcess é responsável pela criação e gerenciamento do processo. 
   *   Como toda classe, a TWFProcess é dividida em métodos e propriedades. 
   *   Veremos alguns dos principais métodos e propriedades que iremos usar para criação de um processo.
   *   
   *   
   *   Métodos
   *   :New(<cCodProc>,<cDescr>,<cProcID>)
   *   O método New() é responsável pela criação e inicialização da classe WFProcess.   
   *   
   *   Parâmetros:
   *   1. cCodProc: este parâmetro recebe o código do processo usado em “Cadastro de Processos”.
   *   2. cDescr: este parâmetro recebe a descrição do processo que está sendo criado no momento. 
   *              Se não for informado, será usada a descrição contida no cadastro de processo, localizada 
   *              através do parâmetro anterior cCodProc.
   *   3. cProcID: este parâmetro recebe o ID do processo criado anteriormente. Normalmente, é utilizado 
   *               para reconstruir um processo anterior, dando seqüência a ele.
   *******************************************************************************************************/
   /*
   Local   nDias := 0, nHoras := 0, nMinutos := 10
   Private cUsuarioProtheus := SubStr(cUsuario,7,15)   // Registre o nome do usuário corrente que está criando o processo:
   Private cCodProcesso := "000001" //"PEDCOM"
   Private cAssunto     := "Aprovacao do Pedido de Compras" 
   Private cTitulo      :=  "Criando Aprovacao"
   Private cHtmlAC      := "C:\Protheus10\Workflow\ApPedCom.htm"
   Private cNumPC       := SC7->C7_NUM
   Private cTo          := "itacolomy@fivesolutions.com.br"
   Private cCC          := "itacolomy@hotmail.com"
   Private cBCC         := "suporte_siga@yahoo.com.br"

   Processa({ ||GHtmWFPC()},"Gerando Html do Pedido "+cNumPC+" Aguarde ...")
   
   //oProcess := TWFProcess():New(cCodProcesso, cAssunto)
   oProcess := TWFProcess():New(cCodProcesso, cAssunto) // Inicializa a classe TWFProcess e assinale a variável objeto oProcess:
   ConOut("Inicializa a classe TWFProcess")
   //oProcess:NewTask("Criando Aprovacao", "\Workflow\ApPedCom.htm" ) 
   oProcess:NewTask(cTitulo, cHtmlAC) // Cria uma tarefa.
   ConOut("Cria uma tarefa NewTask")

   /*************************************
   *  :NewTask(<cDescr>,<cArqHtml>,<lCopiar>)
   *  Este método é responsável por criar a seqüência de tarefas a serem executadas e identificar qual html 
   *  será utilizado pelo processo.
   *  Parâmetros:
   *  1. cDescr  : este parâmetro recebe a descrição da tarefa.
   *  2. cArqHtml: recebe o caminho e o nome do arquivo html que fará uso no processo.   
   *  3. Copiar  : este parâmetro, do tipo lógico, é responsável por copiar todos os campos utilizados em uma 
   *               tarefa anterior, ou seja, para a tarefa que está sendo criada no momento.
   *********************************************************************************************************/

   //:cSubject
   
   oProcess:cSubject := cAssunto   // Repassa o texto do assunto criado para a propriedade específica do processo.
   ConOut("Repassa o texto do assunto - cSubject")
   /*************************************
   *  :cSubject
   *  Esta propriedade define o assunto da mensagem.
   *  Exemplo:
   *  oProcess:cSubject := “Aprovado do pedido de compras no. 1028”
   ***********************************************************/
   
   //:cTo, :cCC e :cBCC
   
   oProcess:cTo  // Informa o endereço eletrônico do destinatário.
   oProcess:cCC  // Informa o endereço eletrônico para cópia.
   oProcess:cBCC // Informa o endereço eletrônico para cópia oculta.
   
   ConOut("Direciona Email para cTo "+cTo+" cCC "+cCC+" cBCC "+cBCC)
   
   //WFSendMail({"99","01"}) 
   //ConOut("Verificando Caixa de Saída 1 ... ")
   oProcess:ClientName(cUsuarioProtheus)
   oProcess:Start()
   ConOut("Inicia Processo - Start()")
   
   WFSendMail({"99","01"})
   ConOut("Verificando Caixa de Saída 2 ... ")
   
   // Repasse as informações para o método responsável pela rastreabilidade.
   //          RastreiaWF( ID do Processo, Codigo do Processo, Codigo do Status, Descricao Especifica, Usuario )

   //Function RastreiaWF( cProcessID, cProcCode, cStatusCode, cDescription, cUserName )
   /*
   ±±³Parametros³ cIDProcess = ID do Processo                                ³±±
   ±±³          ³ cProcesso  = Codigo do Processo                            ³±±
   ±±³          ³ cStatus    = Codigo do Status                              ³±±
   ±±³          ³ cDescr     = Descricao Especifica                          ³±±
   ±±³          ³ cUsuario   = Usuario                                       ³±±                                                           ³±±  
   ±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
   ±±³ Uso      ³ Generico                                                   ³±±
   ±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
   ±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
   ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß

   Function RastreiaWF( cProcessID, cProcCode, cStatusCode, cDescription, cUserName )   
   if !dbSeek( xFilial("WF2") + cProcCode + cStatusCode ) 
      Help("",1,"WFNOSTAT")
      Return .f.
   End
   */
   //oProcess:Track(cCodProces+cCodStatus, cTexto, cUsuarioProtheus) // Rastreabilidade
   //ConOut("Preparando para atualizar Rastreabilidade")
   
   //cCodProces := oProcess:fProcCode
   //ConOut("cProcessID: "+cProcessID+" cCodProcesso "+cCodProcesso+" cProcWFCOM: "+cProcWFCOM+" cCodStatus: "+cProcWFCOM+cCodStatus+" cTexto: "+cTexto+" cUsuarioProtheus: "+cUsuarioProtheus)
   //ConOut("cProcessID: "+cProcessID+" cProcCode: "+cProcCode+" cStatusCode: "+cStatusCode+" cDescription: "+cDescription+" cUserName: "+cUserName)
   //(oProcess:fProcessID+'.'+oProcess:fTaskID,oProcess:fProcCode,'100001','Processo do Pedido '+cNum+' iniciado!')
   //oProcess:Track(cProcessID,cCodProces,cCodStatus, cTexto, cUsuarioProtheus) // Rastreabilidade

   //oProcess:Track(cProcessID,cCodProcesso,cProcWFCOM+cCodStatus, cTexto, cUsuarioProtheus) // Rastreabilidade
   //oProcess:Track(cProcessID,cProcWFCOM,cCodStatus, cTexto, cUsuarioProtheus) // Rastreabilidade


   //oProcess:Track(cProcessID,cProcWFCOM,cCodStatus, cTexto, cUsuarioProtheus) // Rastreabilidade

   //oProcess:Track(cProcessID,cCodProcesso,cCodStatus, cTexto, cUsuarioProtheus) // Rastreabilidade

   // Crie um texto que identifique as etapas do processo que foi realizado
   // para futuras consultas na janela de rastreabilidade.
   cTexto := "Enviado " + cAssunto + " Número: " + cNumPC

   // Informe o código de status correspondente a essa etapa.
   cProcWFCOM := "000001" //oProcess:fProcCode // Código do Processo WorkFlow
   cCodStatus := "100100" // Código do cadastro de status de processo.

   cProcessID := oProcess:fProcessID+'.'+oProcess:fTaskID 
   oProcess:Track(cCodStatus, cTexto, cUsuarioProtheus) // Rastreabilidade
   
   ConOut("Rastreabilidade do Processo - Track")
   //:Track(<cCodStatus>,<cDescr>,<cUsuario>)

   /*************************************
   *  :Track(<cCodStatus>,<cDescr>,<cUsuario>)
   *  Este método é responsável por incluir as descrições dos passos seguidos pelo fluxo do processo e 
   *  apresentá-los na consulta da rastreabilidade.
   *  Parâmetros:
   *  1. cCodStatus: código do status do processo.
   *  2. cDescr: descrição do passo ocorrido.
   *  3. cUsuario: nome do usuário a que se destinou a tarefa.
   *  Exemplo:
   *  oProcess:Track(“100200”, “Enviando o pedido para aprovacao”, “AprovadorA”)   
   *****************************************************************************************************/

   /*************************************
   *  Propriedades
   *  :cTo, :cCC e :cBCC
   *  Estas propriedades definem o endereço dos destinatários. Poderá ser informado mais que um destinatário 
   *  ao mesmo tempo. Basta incluir um ponto-e-vírgula “;” entre eles. Se for informada uma palavra qualquer 
   *  que não seja um endereço de e-mail válido, o Workflow irá considerar que se trata de um diretório em 
   *  que o html gerado deverá ser gravado. Pode-se mesclar os tipos.
   *  Exemplo:
   *  oProcess:cTo := “aluno1@microsiga.com.br;aluno2@microsiga.com.br”
   *  oProcess:cCC := “aluno3@microsiga.com.br;Aluno”
   *  oProcess:cBCC := “aluno4@microsiga.com.br”
   **********************************************************************************************************/


   //:AttachFile(<cArquivo>)

   /*************************************
   *  :AttachFile(<cArquivo>)
   *  Este método é responsável pela inclusão de arquivos anexos à mensagem. 
   *  Esses arquivos deverão estar abaixo do root path do Protheus.
   *  Parâmetros:
   *  cArquivo: caminho e nome do arquivo a ser anexo à mensagem.
   *  Exemplo:
   *  oProcess:AttachFile(“\Workflow\teste.txt”)   
   *******************************************************************************/

   //:Start(<cHtmlCopiarPara>) -> cProcessKey

   /*************************************
   *  :Start(<cHtmlCopiarPara>) -> cProcessKey
   *  Este método é responsável por construir todo o processo, gravar os registros nas tabelas do Workflow 
   *  e enviar a mensagem para os destinatários. O valor de retorno é uma chave composta pela codificação 
   *  ProcessID + TaskID e números aleatórios, no total de 17 algarismos em hexadecimal.
   *  Parâmetros:
   *  1. cHtmlCopiarPara: Caminho em que o Workflow deverá realizar uma cópia do html final.
   *  Exemplo:
   *  cID := oProcess:Start(“\Workflow\copia”)
   *  if file( “\Workflow\copia\” + cID)
   *  conout(“Arquivo copiado com sucesso.”)
   *  endif
   **********************************************************************************************************/

   //oProcess:Finish()

   /*************************************
   *  :Finish()
   *  Este método é responsável por finalizar o processo. Após a finalização, ele não estará mais disponível 
   *  para execuções do tipo retorno e timeout.
   *  Exemplo:
   *  oProcess:Finish()   
   ***********************************************************************************************************/

   //:cBody
   
   //oProcess:cBody := cHtmlAC
   
   /*************************************
   *  :cBody
   *  Esta propriedade armazenará um texto que permanecerá no corpo da mensagem. Caso seja utilizado, 
   *  o html irá como anexo da mensagem.
   *  Exemplo:
   *  oProcess:cBody := “Testando...”   
   ***************************************************************************************************/

   //:bReturn

   // Função de retorno a ser executada quando a mensagem de
   // respostas retornar ao Workflow:
   
   oProcess:bReturn := "u_RetAprvPC(oProcess)"
   
   //ConOut('oProcess:bReturn := u_RetAprvPC(oProcess)')
   /*************************************
   *  :bReturn
   *  Esta propriedade contém o nome da função que será executada no momento em que o Workflow receber a 
   *  mensagem de resposta de um dos destinatários via e-mail ou serviço http.
   *  Exemplo:
   *  oProcess:bReturn := “U_Retorno”   
   ****************************************************************************************************/

   //:bTimeOut

   // Função do tipo timeout que será executada se houver um timeout
   // ocorrido para esse processo. Neste exemplo, ela será executada cinco minutos após o envio
   // do e-mail para o destinatário. Caso queira-se aumentar ou diminuir o tempo, altere
   // os valores das variáveis: nDias, nHoras e nMinutos.
   
   //oProcess:bTimeOut := {"u_TOutApPC", nDias, nHoras, nMinutos}
   
   /*************************************
   *  :bTimeOut
   *  Esta propriedade recebe um array de timeouts contendo nomes das funções e tempo de espera. 
   *  Caso o tempo seja alcançado, serão executadas as funções mencionadas no 1º item do array. 
   *  Poderão ser especificados mais do que um array de timeouts.
   *  { { <cFuncao>, <nDias>, <nHoras>, <nMinutos> }, { ... } }
   *  Exemplo:
   *  oProcess:bTimeOut := { { “TimeOut1”, 0, 5, 30 } }
   *  ou
   *  oProcess:bTimeOut := { { “TimeOut1”, 0, 5, 30 }, { “TimeOut2”, 1, 10, 0 } }
   ************************************************************************************************/

   //:fProcessID

   // Utilize a funcao WFCodUser para obter o código do usuário Protheus.
   //oProcess:UserSiga := WFCodUser("BI")
   //ConOut("oProcess:UserSiga := WFCodUser('BI')")
   // Antes de assinalar o ID do processo no campo, é verificado se realmente o
   // campo existe na tabela SC7-Pedido de Compras.
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
   //ConOut("Registrou oProcess:fProcessID no SC7->C7_WFID")
   /*************************************
   *  :fProcessID
   *  Esta propriedade fornece o número ID do processo.
   *  Exemplo:
   *  cProcID := oProcess:fProcessID   
   *******************************************************/

   //:fTaskID
   
   /*************************************
   *  :fTaskID
   *  Esta propriedade fornece o número ID da tarefa criada para um determinado processo, 
   *  através do método :NewTask().
   *  Exemplo:
   *  oProcess:NewTask( "100100", "\Workflow\WFW120p.htm" )
   *  cTaskID := oProcess:fTaskID   
   *****************************************************************************************/

   //:oHTML
   // Assinale novos valores às macros existentes no html:
   /*
   dEmiPC   := SC7->C7_EMISSAO
   cForcdPC := SC7->C7_FORNECE + " / " + SC7->C7_LOJA
   cCondPgto:= SC7->C7_COND
   cForNome := Alltrim(Posicione("SA2",1,xFilial("SA2")+SC7->(C7_FORNECE+C7_LOJA),"A2_NOME"))
   DbSelectArea("SC7")
   
   oProcess:oHtml:ValByName("emissao"    , dEmiPC   )
   oProcess:oHtml:ValByName("fornecedor" , cForcdPC )
   oProcess:oHtml:ValByName("lb_nome"    , cForNome )
   oProcess:oHtml:ValByName("lb_cond"    , cCondPgto)
   oProcess:oHtml:ValByName("pedido"     , cNumPC   )
   */

   /*************************************
   *  :oHTML
   *  Esta propriedade é responsável pelo tratamento das palavras-chaves no html mencionado no método :NewTask(). 
   *  Esse objeto é uma referência da classe TWFHtml() e disponibiliza dois dos métodos mais importantes
   *  :RetByName() e ValByName().
   *
   *  :oHTML:RetByName( <cNome da macro> )
   *  
   *  Esta propriedade tem como objetivo obter o conteúdo da “macro” quando as respostas retornarem para o Workflow.
   *  O método somente deverá ser usado na função de retorno.
   *  Exemplo:
   *  cNome := oProcess:oHtml:RetByName( “Nome” )
   *  Em uma tabela:
   *  aCodigo := oProcess:oHtml:RetByName( “produto.Codigo” )
   *  :oHTML:ValByName(<cMacro>,<uConteudo>)
   *  Este método tem como objetivo atribuir ou obter um valor a uma “macro” existente no html. 
   *  Deverá ser usado somente no momento em que estiver assinalando valores ao html ou na função de timeouts, 
   *  em que o uso é necessário por motivo de não haver recebido resposta.
   *  Parâmetros:
   *  cMacro: nome da macro (palavra-chave) encontrada no html e identificada entre os símbolos “%” e “!”.
   *  uConteudo: valor a ser atribuído à macro.
   *  Exemplo:
   *  Assinalando um valor:
   *  oProcess:oHtml:ValByName( “Nome”, “Aluno1” )
   *  Usando na função de timeout:
   *  cNome := oProcess:oHtml:ValByName( “Nome” )
   *  Atribuindo um valor a uma tabela:
   *  AAdd( oProcess:oHtml:ValByName( “produto.codigo” ), SB1->SB1_COD )
   ***********************************************************************************************************/

   //WFReadMail()

   //WFReturn({"99","01"}) 

   // O método Start() gera todo o processo e envia a mensagem
   // ao destinatário.

   //oProcess:Finish()   
   //ConOut("oProcess:Finish()")
   /*
   WFReturn({"99","01"})
   ConOut("Executei WFReturn") 
   WFSendMail({"99","01"})
   ConOut("Executei WFSendMail") */


   oProcess:Finish()
Return

/***************************************************************************
*
*                        G E R A N D O   O   H T M L
*
*****************************************************************************/

Static Function GHtmWFPC

   For nPC := 1 To 2
   
      If nPC == 1
         cQryPC  := " SELECT COUNT(*) NQTDREG"
      Else   
         cQryPC  := " SELECT SC7.C7_NUM,SC7.C7_EMISSAO,SC7.C7_FORNECE,SC7.C7_LOJA,SA2.A2_NOME,SC7.C7_COND,"
         cQryPC  += "        SC7.C7_ITEM,SC7.C7_PRODUTO,SB1.B1_DESC,SC7.C7_QUANT,SC7.C7_UM,SC7.C7_PRECO,SC7.C7_TOTAL,"
         cQryPC  += "        SC7.C7_DATPRF"
      EndIf
   
      cQryPC  += "   FROM "+RetSQLName("SC7")+" SC7,"+RetSQLName("SA2")+" SA2,"+RetSQLName("SB1")+" SB1"
      cQryPC  += "  WHERE SC7.C7_FILIAL = '"+xFilial("SC7")+"'"
      cQryPC  += "    AND SA2.A2_FILIAL = '"+xFilial("SA2")+"'"
      cQryPC  += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"
   
      cQryPC  += "    AND SC7.C7_FORNECE = SA2.A2_COD"
      cQryPC  += "    AND SC7.C7_LOJA = SA2.A2_LOJA"
      cQryPC  += "    AND SC7.C7_PRODUTO = SB1.B1_COD"
   
      cQryPC  += "    AND SC7.C7_NUM = '"+SC7->C7_NUM+"'"
      cQryPC  += "    AND SC7.C7_FORNECE = '"+SC7->C7_FORNECE+"'"
      cQryPC  += "    AND SC7.C7_LOJA = '"+SC7->C7_LOJA+"'"
   
      cQryPC  += "    AND SC7.D_E_L_E_T_ <> '*'"
      cQryPC  += "    AND SA2.D_E_L_E_T_ <> '*'"
      cQryPC  += "    AND SB1.D_E_L_E_T_ <> '*'"
   
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
   cHtmlWF += '<div align="center" class="style1"><span class="style5">Aprova&ccedil;&atilde;o do Pedido de Compras</span></div>'
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
      cHtmlWF += '  <tr>'
      cHtmlWF += '    <td>'+TXC7->C7_ITEM+'</td>'                                        //cHtmlWF += '    <td>!produto.item!</td>'
      cHtmlWF += '    <td>'+TXC7->C7_PRODUTO+'-'+Alltrim(TXC7->B1_DESC)+'</td>'          //cHtmlWF += '    <td>!produto.codigo!-!produto.descricao!</td>'
      cHtmlWF += '    <td>'+Transform(TXC7->C7_QUANT,"@E 999,999,999.999")+'</td>' //cHtmlWF += '    <td>%produto.quant%</td>'
      cHtmlWF += '    <td>'+TXC7->C7_UM+'</td>'                                          //cHtmlWF += '    <td>!produto.unid!</td>'
      cHtmlWF += '    <td>'+Transform(TXC7->C7_PRECO,"@E 999,999,999.99")+'</td>'    //cHtmlWF += '    <td>%produto.preco%</td>'
      cHtmlWF += '    <td>'+Transform(TXC7->C7_TOTAL,"@E 999,999,999.99")+'</td>'    //cHtmlWF += '    <td>%produto.total%</td>'
      cHtmlWF += '    <td>'+DTOC(TXC7->C7_DATPRF)+'</td>'                        //cHtmlWF += '    <td>%produto.entrega%</td>'
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
   cHtmlWF += '<br>'
   cHtmlWF += '<input type="radio" name="%aprovacao%" value="Nao">'
   cHtmlWF += 'Não(Favor especificar Motivo) <br />'
   cHtmlWF += '<br />'
   cHtmlWF += '	<br />'
   cHtmlWF += '<form id="form1" name="form1" method="post" action="">'
   cHtmlWF += '  <label>Motivo'
   cHtmlWF += '  '
   //cHtmlWF += '  <textarea name="textarea" cols="80" rows="5">%lbmotivo%</textarea> '
   //cHtmlWF += '  <textarea name="textarea" cols="80" rows="5"><input type=hidden name="lbmotivo" value = " "></textarea>'
   cHtmlWF += '  <textarea name="textarea" cols="80" rows="5">%lbmotivo%</textarea>'
   cHtmlWF += '  </label>'
   cHtmlWF += '</form>'
   cHtmlWF += '<br />'
   cHtmlWF += '<form id="form2" name="form2" method="post" action="">'
   cHtmlWF += '  <input name="Enviar" type="submit" id="Enviar" value="Enviar" />'
   cHtmlWF += '  <input name="Limpar" type="submit" id="Limpar" value="Limpar" />'
   cHtmlWF += '</form>'
   cHtmlWF += '</body>'
   cHtmlWF += '</html>'
   
   //MemoWrite("C:\Protheus10\Workflow\ApPedCom.htm",cHtmlWF)
   //ConOut("Gerou html em C:\Protheus10\Workflow\ApPedCom.htm")

   MemoWrite("\Workflow\ApPedCom.htm",cHtmlWF)
   ConOut("Gerou NOVO - html em \Workflow\ApPedCom.htm")
Return

User Function RetAprvPC(oProcess)
   /*
    APVRetorno - Esta função é executada no retorno da mensagem enviada
    pelo destinatário. O Workflow recria o processo em que
    parou anteriormente na função APVInicio e repassa a
    variável objeto oProcess por parâmetro.
   */
   ConOut("Estou no Retorno RetAprvPC")
   // Obtem aResposta - Aprovado Sim ou Não partir do html
   cRespApv := oProcess:oHtml:RetByName("aprovacao")

   // Adicione as informacões a serem incluídas na rastreabilidade.
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
   cAprovador := oProcess:cRetFrom
   oProcess:Track(cCodProces+cCodStatus, cTexto, oProcess:cRetFrom) // Rastreabilidade
   
   cMsgRet := "o APROVADOR FOI "+cAprovador+" A Resposta do Aprovador foi: "+cRespApv
   
   MemoWrite("Resposta_Aprovacao_de_Pedido_de_Compras.TXT",cMsgRet)

Return

User Function TOutApPC
Return



/**************************************************************************************************************




Utilize o ponto de Entrada e crie a estrutura para enviar o E-mail, segue exemplo abaixo:

cLista   := "teste@teste.com.br"
oProcess := TWFProcess():New( "000001", "TESTE" )
oProcess:NewTask( "TESTE", "\WORKFLOW\EVENTO\TESTE.HTML")
oHtml    := oProcess:oHTML

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÕ¿
//³Inicia o preenchimento da HTML³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÕÙ

oHtml := oProcess:oHTML
oProcess:cSubject := "TESTE "
oProcess:cTo      := cLista 
oProcess:Start()
RastreiaWF(oProcess:fProcessID+'.'+oProcess:fTaskID,WF1->WF1_COD,WF2->WF2_STATUS,"Email´s Enviado´s Para os Usuario´s Envolvido´s")
WFSendMail()
oProcess:Finish()
WFSendMail()


****************************************************************************************************************/