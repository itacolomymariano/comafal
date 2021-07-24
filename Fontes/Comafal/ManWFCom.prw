#include "Protheus.ch"
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³MANWFCOM  ºAutor  ³Five Solutions      º Data ³  15/12/2010 º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Tela de Manutenção do Workflow para Alçada de Compras      º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ COMAFAL - Workflows-Compras - PE, SP e RS.                º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function ManWFCom

   Private o1VlrNv
   Private o1ApCOSP
   Private o1ApCORS
   Private o1ApCOPE

   Private o2VlPNv
   Private o2VlrNv
   Private o2ApCOSP
   Private o2ApCORS
   Private o2ApCOPE
   Private o2CCopia

   Private o3VlPNv
   Private o3VlrNv
   Private o3ApUnk
   Private o3VlpCc
   Private o3CCopia
   /*
   Private o3ApCOSP
   Private o3ApCORS
   Private o3ApCOPE
   */
   
   Private o4VlPNv
   Private o4ApUnk
      
   Private n1VlrNv := GetMv("FS_VAL1NV")            //Valor Máximo para aprovação de pedido de compras Nível I
   Private c1ApCOSP:= GetMV("FS_APN1SP")+Space(150) //Aprovador Nível I de COSP - COMAFAL-SP
   Private c1ApCORS:= GetMV("FS_APN1RS")+Space(150) //Aprovador Nível I de CORS - COMAFAL-RS
   Private c1ApCOPE:= GetMV("FS_APN1PE")+Space(150) //Aprovador Nível I de COPE - COMAFAL-PE

   Private n2VlPNv := GetMv("FS_VLP2NV")            //Valor Mínimo para aprovação de pedido de compras Nível II
   Private n2VlrNv := GetMv("FS_VAL2NV")            //Valor Máximo para aprovação de pedido de compras Nível II
   Private c2ApCOSP:= GetMV("FS_APN2SP")+Space(150) //Aprovador Nível II de COSP - COMAFAL-SP
   Private c2ApCORS:= GetMV("FS_APN2RS")+Space(150) //Aprovador Nível II de CORS - COMAFAL-RS
   Private c2ApCOPE:= GetMV("FS_APN2PE")+Space(150) //Aprovador Nível II de COPE - COMAFAL-PE
   Private c2CCopia:= GetMV("FS_CCN2PC")+Space(150) //Gestor que receberá cópia dos pedidos de compras para aprovação Nivel II.

   Private n3VlPNv := GetMv("FS_VLP3NV")            //Valor Mínimo para aprovação de pedido de compras Nível III 
   Private n3VlrNv := GetMv("FS_VAL3NV")            //Valor Máximo para aprovação de pedido de compras Nível III
   Private c3ApUnk := GetMv("FS_APR3NV")+Space(150) //Único Aprovador Nível III
   Private n3VlpCc := GetMv("FS_N3VMCC")            //Valor Minimo para que o pedido nível III seja enviado com cópia para o Gestor.
   Private c3CCopia:= GetMV("FS_CCN3PC")+Space(150) //Gestor que receberá cópia dos pedidos de compras para aprovação Nivel III quando o valor atingir o mínimo exigido para cópia.
   
   /*
   Private c3ApCOSP:= GetMV("FS_APN3SP") //Aprovador Nível III de COSP - COMAFAL-SP
   Private c3ApCORS:= GetMV("FS_APN3RS") //Aprovador Nível III de CORS - COMAFAL-RS
   Private c3ApCOPE:= GetMV("FS_APN3PE") //Aprovador Nível III de COPE - COMAFAL-PE
   */
   
   Private n4VlPNv := GetMv("FS_VLP4NV")            //Valor Mínimo para aprovação de pedido de compras Nível IV 
   Private c4ApUnk := GetMv("FS_APR4NV")+Space(150) //Único Aprovador Nível IV
   
   //DEFINE MSDIALOG oDlgWFCom FROM 000,000 TO 480,800 TITLE "Manutenção Workflow - Alçadas de Compras" Of oMainWnd PIXEL
   DEFINE MSDIALOG oDlgWFCom FROM 000,000 TO 680,800 TITLE "Manutenção Workflow - Alçadas de Compras" Of oMainWnd PIXEL
   
   //@ 000,000 BITMAP oBmp RESNAME "PROJETOAP" oF oDlgWFCom SIZE 60,150 NOBORDER WHEN .F. PIXEL
   @ 015,015 BITMAP oBmp RESNAME "LOGIN" oF oDlgWFCom SIZE 55,1000 NOBORDER WHEN .F. PIXEL
   
   DEFINE FONT oBold NAME "Arial" SIZE 0, -13 BOLD
   
   @ 005,065 TO 085,370 LABEL "Alçada Nível I             " Of oDlgWFCom PIXEL
   
   @ 017,070 Say "Valor máximo para aprovação" SIZE 80,10 Of oDlgWFCom PIXEL
   @ 015,160 MsGet o1VlrNv VAR n1VlrNv SIZE 50,10  Picture "@E 999,999,999.99" Of oDlgWFCom PIXEL VALID(fVn1VlrNv()) WHEN .T.
   
   @ 037,070 Say "Aprovador COSP " SIZE 60,10 Of oDlgWFCom PIXEL
   @ 035,140 MsGet o1ApCOSP VAR c1ApCOSP SIZE 220,10 Picture "@S" Of oDlgWFCom PIXEL WHEN If(n1VlrNv > 0,.T.,.F.)
   @ 052,070 Say "Aprovador CORS " SIZE 60,10 Of oDlgWFCom PIXEL
   @ 050,140 MsGet o1ApCORS VAR c1ApCORS SIZE 220,10 Picture "@S" Of oDlgWFCom PIXEL WHEN If(n1VlrNv > 0,.T.,.F.)
   @ 067,070 Say "Aprovador COPE " SIZE 60,10 Of oDlgWFCom PIXEL
   @ 065,140 MsGet o1ApCOPE VAR c1ApCOPE SIZE 220,10 Picture "@S" Of oDlgWFCom PIXEL WHEN If(n1VlrNv > 0,.T.,.F.)
   
   @ 090,065 TO 190,370 LABEL "Alçada Nível II            " Of oDlgWFCom PIXEL

   @ 102,070 Say "Valor mínimo para aprovação" SIZE 80,10 Of oDlgWFCom PIXEL
   @ 100,160 MsGet o2VlPNv VAR n2VlPNv SIZE 50,10  Picture "@E 999,999,999.99" Of oDlgWFCom PIXEL VALID(fVn2VlPNv()) WHEN .T.
   
   @ 102,220 Say "Valor máximo para aprovação" SIZE 80,10 Of oDlgWFCom PIXEL
   @ 100,310 MsGet o2VlrNv VAR n2VlrNv SIZE 50,10  Picture "@E 999,999,999.99" Of oDlgWFCom VALID(fVn2VlrNv()) PIXEL WHEN .T.
   
   @ 122,070 Say "Aprovador COSP " SIZE 60,10 Of oDlgWFCom PIXEL
   @ 120,140 MsGet o2ApCOSP VAR c2ApCOSP SIZE 220,10 Picture "@S" Of oDlgWFCom PIXEL WHEN If(n2VlrNv > 0,.T.,.F.)
   @ 137,070 Say "Aprovador CORS " SIZE 60,10 Of oDlgWFCom PIXEL
   @ 135,140 MsGet o2ApCORS VAR c2ApCORS SIZE 220,10 Picture "@S" Of oDlgWFCom PIXEL WHEN If(n2VlrNv > 0,.T.,.F.)
   @ 152,070 Say "Aprovador COPE " SIZE 60,10 Of oDlgWFCom PIXEL
   @ 150,140 MsGet o2ApCOPE VAR c2ApCOPE SIZE 220,10 Picture "@S" Of oDlgWFCom PIXEL WHEN If(n2VlrNv > 0,.T.,.F.)
   @ 167,070 Say "Com Copia para " SIZE 60,10 Of oDlgWFCom PIXEL
   @ 165,140 MsGet o2CCopia VAR c2CCopia SIZE 220,10 Picture "@S" Of oDlgWFCom PIXEL WHEN If((Alltrim(cUserName)=="RLIRA" .Or. Alltrim(cUserName)=="Administrador"),.T.,.F.)   
   
   @ 195,065 TO 275,370 LABEL "Alçada Nível III           " Of oDlgWFCom PIXEL

   @ 207,070 Say "Valor mínimo para aprovação" SIZE 80,10 Of oDlgWFCom PIXEL
   @ 205,160 MsGet o3VlPNv VAR n3VlPNv SIZE 50,10  Picture "@E 999,999,999.99" Of oDlgWFCom PIXEL VALID(fVn3VlPNv()) WHEN .T.
   
   @ 207,220 Say "Valor máximo para aprovação" SIZE 80,10 Of oDlgWFCom PIXEL
   @ 205,310 MsGet o3VlrNv VAR n3VlrNv SIZE 50,10  Picture "@E 999,999,999.99" Of oDlgWFCom PIXEL VALID(fVn3VlrNv()) WHEN .T.
   
   @ 227,070 Say "Aprovador Nível III " SIZE 60,10 Of oDlgWFCom PIXEL
   @ 225,140 MsGet o3ApUnk VAR c3ApUnk SIZE 220,10 Picture "@S" Of oDlgWFCom PIXEL WHEN If(n3VlrNv > 0,.T.,.F.)
   
   @ 242,070 Say "Valor p/Copia " SIZE 60,10 Of oDlgWFCom PIXEL
   @ 240,140 MsGet o3VlpCc VAR n3VlpCc SIZE 50,10 Picture "@E 999,999,999.99" Of oDlgWFCom PIXEL VALID(fVn3VlpCc()) WHEN If(n3VlrNv > 0,.T.,.F.)
   
   @ 257,070 Say "Com Copia para " SIZE 60,10 Of oDlgWFCom PIXEL
   @ 255,140 MsGet o3CCopia VAR c3CCopia SIZE 220,10 Picture "@S" Of oDlgWFCom PIXEL WHEN If((Alltrim(cUserName)=="Administrador"),.T.,.F.)   
   
   @ 280,065 TO 325,370 LABEL "Alçada Nível IV            " Of oDlgWFCom PIXEL

   @ 292,070 Say "Aprova valores superior a  " SIZE 80,10 Of oDlgWFCom PIXEL
   @ 290,160 MsGet o4VlPNv VAR n4VlPNv SIZE 50,10  Picture "@E 999,999,999.99" Of oDlgWFCom PIXEL VALID(fVn4VlPNv()) WHEN If((Alltrim(cUserName)=="Administrador"),.T.,.F.)   

   @ 307,070 Say "Aprovador Nível IV  " SIZE 60,10 Of oDlgWFCom PIXEL
   @ 305,140 MsGet o4ApUnk VAR c4ApUnk SIZE 220,10 Picture "@S" Of oDlgWFCom PIXEL WHEN If(n3VlrNv > 0,.T.,.F.) .And. If((Alltrim(cUserName)=="Administrador"),.T.,.F.)   

   @ 330,250 BUTTON oSaveMC PROMPT "Salvar" SIZE 50,12 ACTION(fGMWFC()) Of oDlgWFCom PIXEL
   @ 330,310 BUTTON oCancMC PROMPT "Sair" SIZE 50,12 Of oDlgWFCom PIXEL ACTION oDlgWFCom:End() //(fFecMan())
   
   ACTIVATE MSDIALOG oDlgWFCom CENTER
   
Return

/***********************************************************************************************
*
*        V A L I D A Ç Ã O   D O S   V A L O R E S   M Í N I M O S   D E   A P R O V A Ç Ã O
*
************************************************************************************************/

Static Function fVn1VlrNv() //Validação do Valor Máximo de Aprovação do Nível I
   
   If n1VlrNv > n2VlPNv
      Alert("O valor máximo de aprovação do Nível I R$ "+Transform(n1VlrNv,"@E 999,999,999.99")+" não poderá ser superior ao valor mínimo para aprovação do Nível II R$ "+Transform(n2VlPNv,"@E 999,999,999.99"))
      n1VlrNv := 0
      o1VlrNv:Refresh()
      Return(.F.)
   EndIf
   
   If n1VlrNv > n3VlPNv
      Alert("O valor máximo de aprovação do Nível I R$ "+Transform(n1VlrNv,"@E 999,999,999.99")+" não poderá ser superior ao valor mínimo para aprovação do Nível III R$ "+Transform(n3VlPNv,"@E 999,999,999.99"))
      n1VlrNv := 0
      o1VlrNv:Refresh()
      Return(.F.)
   EndIf

   If n1VlrNv > n4VlPNv
      Alert("O valor máximo de aprovação do Nível I R$ "+Transform(n1VlrNv,"@E 999,999,999.99")+" não poderá ser superior ao valor mínimo para aprovação do Nível IV R$ "+Transform(n4VlPNv,"@E 999,999,999.99"))
      n1VlrNv := 0
      o1VlrNv:Refresh()
      Return(.F.)
   EndIf
      
Return(.T.)

Static Function fVn2VlPNv() //Validação do Valor Mínimo de Aprovação do Nível II
   
   If n2VlPNv < n1VlrNv
      Alert("O Valor Mínimo para Aprovação do Nível II R$ "+Transform(n2VlPNv,"@E 999,999,999.99")+" não poderá ser inferior ao ao valor máximo de aprovação do Nível I R$ "+Transform(n1VlrNv,"@E 999,999,999.99"))
      n2VlPNv := 0 
      o2VlPNv:Refresh()
      Return(.F.)
   EndIf
   If n2VlPNv > n2VlrNv
      Alert("O Valor Mínimo para Aprovação do Nível II R$ "+Transform(n2VlPNv,"@E 999,999,999.99")+" não poderá ser superior ao valor máximo de aprovação do Nível II R$ "+Transform(n2VlrNv,"@E 999,999,999.99"))
      n2VlPNv := 0 
      o2VlPNv:Refresh()
      Return(.F.)
   EndIf

   If n2VlPNv > n3VlPNv
      Alert("O Valor Mínimo para Aprovação do Nível II R$ "+Transform(n2VlPNv,"@E 999,999,999.99")+" não poderá ser superior ao valor mínimo de aprovação do Nível III R$ "+Transform(n3VlPNv,"@E 999,999,999.99"))
      n2VlPNv := 0 
      o2VlPNv:Refresh()
      Return(.F.)
   EndIf
   
Return(.T.)

Static Function fVn3VlPNv() //Validação do Valor Mínimo de Aprovação do Nível III
   
   If n3VlPNv < n2VlrNv
      Alert("O Valor Mínimo para Aprovação do Nível III R$ "+Transform(n3VlPNv,"@E 999,999,999.99")+" não poderá ser inferior ao ao valor máximo de aprovação do Nível II R$ "+Transform(n2VlrNv,"@E 999,999,999.99"))
      n3VlPNv := 0 
      o3VlPNv:Refresh()
      Return(.F.)
   EndIf
   
   If n3VlPNv > n3VlrNv
      Alert("O Valor Mínimo para Aprovação do Nível III R$ "+Transform(n3VlPNv,"@E 999,999,999.99")+" não poderá ser superior ao valor máximo de aprovação do Nível III R$ "+Transform(n3VlrNv,"@E 999,999,999.99"))
      n3VlPNv := 0 
      o3VlPNv:Refresh()
      Return(.F.)
   EndIf
   
   If n3VlPNv > n4VlPNv
      Alert("O Valor Mínimo para Aprovação do Nível III R$ "+Transform(n3VlPNv,"@E 999,999,999.99")+" não poderá ser superior ao valor mínimo de aprovação do Nível IV R$ "+Transform(n4VlPNv,"@E 999,999,999.99"))
      n3VlPNv := 0 
      o3VlPNv:Refresh()
      Return(.F.)
   EndIf
   
Return(.T.)

/***********************************************************************************************
*
*        V A L I D A Ç Ã O   D O S   V A L O R E S   M Á X I M O S   D E   A P R O V A Ç Ã O
*
************************************************************************************************/

Static Function fVn2VlrNv() //Validação do Valor Máximo de Aprovação do Nível II
   If n2VlrNv < n2VlPNv
      Alert("O Valor Máximo para Aprovação do Nível II R$ "+Transform(n2VlrNv,"@E 999,999,999.99")+" não poderá ser inferior ao valor mínimo de aprovação do Nível II R$ "+Transform(n2VlPNv,"@E 999,999,999.99"))
      n2VlrNv := 0 
      o2VlrNv:Refresh()
      Return(.F.)
   EndIf
   If n2VlrNv < n1VlrNv
      Alert("O Valor Máximo para Aprovação do Nível II R$ "+Transform(n2VlrNv,"@E 999,999,999.99")+" não poderá ser inferior ao valor máximo de aprovação do Nível I R$ "+Transform(n1VlrNv,"@E 999,999,999.99"))
      n2VlrNv := 0 
      o2VlrNv:Refresh()
      Return(.F.)
   EndIf
   If n2VlrNv > n3VlPNv
      Alert("O Valor Máximo para Aprovação do Nível II R$ "+Transform(n2VlrNv,"@E 999,999,999.99")+" não poderá ser supeior ao valor mínimo de aprovação do Nível III R$ "+Transform(n3VlPNv,"@E 999,999,999.99"))
      n2VlrNv := 0 
      o2VlrNv:Refresh()
      Return(.F.)
   EndIf
Return

Static Function fVn3VlrNv() //Validação do Valor Máximo de Aprovação do Nível III
   If n3VlrNv < n3VlPNv
      Alert("O Valor Máximo para Aprovação do Nível III R$ "+Transform(n3VlrNv,"@E 999,999,999.99")+" não poderá ser inferior ao valor mínimo de aprovação do Nível III R$ "+Transform(n3VlPNv,"@E 999,999,999.99"))
      n3VlrNv := 0 
      o3VlrNv:Refresh()
      Return(.F.)
   EndIf
   If n3VlrNv < n2VlrNv
      Alert("O Valor Máximo para Aprovação do Nível III R$ "+Transform(n3VlrNv,"@E 999,999,999.99")+" não poderá ser inferior ao valor máximo de aprovação do Nível II R$ "+Transform(n2VlrNv,"@E 999,999,999.99"))
      n3VlrNv := 0 
      o3VlrNv:Refresh()
      Return(.F.)
   EndIf
   If n3VlrNv > n4VlPNv
      Alert("O Valor Máximo para Aprovação do Nível III R$ "+Transform(n3VlrNv,"@E 999,999,999.99")+" não poderá ser supeior ao valor mínimo de aprovação do Nível IV R$ "+Transform(n4VlPNv,"@E 999,999,999.99"))
      n3VlrNv := 0 
      o3VlrNv:Refresh()
      Return(.F.)
   EndIf
Return

Static Function fVn3VlpCc()
   If n3VlpCc < n3VlPNv
      Alert("O Valor para cópiar a Aprovação do Nível III R$ "+Transform(n3VlpCc,"@E 999,999,999.99")+" não poderá ser inferior ao valor mínimo de aprovação do Nível III R$ "+Transform(n3VlPNv,"@E 999,999,999.99"))
      n3VlpCc := 0 
      o3VlpCc:Refresh()
      Return(.F.)
   EndIf
   If n3VlpCc > n3VlrNv
      Alert("O Valor para copiar a Aprovação do Nível III R$ "+Transform(n3VlpCc,"@E 999,999,999.99")+" não poderá ser superior ao valor máximo de aprovação do Nível III R$ "+Transform(n3VlrNv,"@E 999,999,999.99"))
      n3VlpCc := 0 
      o3VlpCc:Refresh()
      Return(.F.)
   EndIf
Return(.T.)
/***********************************************************************************************
*
*                            G R A V A Ç Ã O     D O S     P A R Â M E T R O S                                     
*
************************************************************************************************/

Static Function fVn4VlPNv() //Validação do Valor Mínimo de Aprovação do Nível IV

   If n4VlPNv < n3VlrNv
      Alert("O Valor Mínimo para Aprovação do Nível IV R$ "+Transform(n4VlPNv,"@E 999,999,999.99")+" não poderá ser inferior ao valor máximo de aprovação do Nível III R$ "+Transform(n3VlrNv,"@E 999,999,999.99"))
      n4VlPNv := 0 
      o4VlPNv:Refresh()
      Return(.F.)
   EndIf

Return(.T.)

Static Function fGMWFC() //Grava os Parâmetros atualizados durante a manutenção.
   
   PutMv("FS_VAL1NV",n1VlrNv) //Valor Máximo para aprovação de pedido de compras Nível I
   PutMv("FS_APN1SP",c1ApCOSP) //Aprovador Nível I de COSP - COMAFAL-SP
   PutMv("FS_APN1RS",c1ApCORS) //Aprovador Nível I de CORS - COMAFAL-RS
   PutMv("FS_APN1PE",c1ApCOPE) //Aprovador Nível I de COPE - COMAFAL-PE

   PutMv("FS_VLP2NV",n2VlPNv) //Valor Mínimo para aprovação de pedido de compras Nível II
   PutMv("FS_VAL2NV",n2VlrNv) //Valor Máximo para aprovação de pedido de compras Nível II
   PutMv("FS_APN2SP",c2ApCOSP) //Aprovador Nível II de COSP - COMAFAL-SP
   PutMv("FS_APN2RS",c2ApCORS) //Aprovador Nível II de CORS - COMAFAL-RS
   PutMv("FS_APN2PE",c2ApCOPE) //Aprovador Nível II de COPE - COMAFAL-PE
   PutMv("FS_CCN2PC",c2CCopia) //Gestor que receberá cópia dos pedidos de compras para aprovação Nivel II.

   PutMv("FS_VLP3NV",n3VlPNv) //Valor Mínimo para aprovação de pedido de compras Nível III 
   PutMv("FS_VAL3NV",n3VlrNv) //Valor Máximo para aprovação de pedido de compras Nível III
   PutMv("FS_APR3NV",c3ApUnk) //Único Aprovador Nível III
   PutMv("FS_N3VMCC",n3VlpCc) //Valor Minimo para que o pedido nível III seja enviado com cópia para o Gestor.
   PutMv("FS_CCN3PC",c3CCopia) //Gestor que receberá cópia dos pedidos de compras para aprovação Nivel III quando o valor atingir o mínimo exigido para cópia.
   
   PutMv("FS_VLP4NV",n4VlPNv) //Valor Mínimo para aprovação de pedido de compras Nível IV 
   PutMv("FS_APR4NV",c4ApUnk) //Único Aprovador Nível IV
   
   o1VlrNv:Refresh()
   o1ApCOSP:Refresh()
   o1ApCORS:Refresh()
   o1ApCOPE:Refresh()

   o2VlPNv:Refresh()
   o2VlrNv:Refresh()
   o2ApCOSP:Refresh()
   o2ApCORS:Refresh()
   o2ApCOPE:Refresh()
   o2CCopia:Refresh()

   o3VlPNv:Refresh()
   o3VlrNv:Refresh()
   o3ApUnk:Refresh()
   o3VlpCc:Refresh()
   o3CCopia:Refresh()
  
   o4VlPNv:Refresh()
   o4ApUnk:Refresh()

   MsgInfo("Parâmetros Atualizados com Sucesso!")
   oDlgWFCom:End()
   
Return