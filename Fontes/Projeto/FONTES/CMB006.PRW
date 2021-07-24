#include "rwmake.ch"     

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Programa  ³ CMb006   ³ Autor ³ Valmir Belchior       ³ Data ³ 14/05/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Este Rdmake compoe as Rotinas de geracao do SISPAG, arquivo³±±
±±³          ³ 341REM.PAG e 341RET.PAG                                    ³±±
±±³          ³ Calcula o layout para o Valor no Codigo de Barras          ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Adaptado por Jeniffer Larrosa para a versao 7.10                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±*/
User Function CMb006()    

// Variaveis Privadas
PRIVATE cValCamp  := ""

If Len(Alltrim(SE2->E2_X_CODBA)) == 47
	cValCamp := Substr(SE2->E2_X_CODBA,34,14)
ElseIf Len(Alltrim(SE2->E2_X_CODBA)) == 44	       // Alterado de 44 p/ 47. Comtempla o Fator de Vencimento.
	cValCamp := Substr(SE2->E2_X_CODBA,6,14)
ElseIf Len(Alltrim(SE2->E2_X_CODBA)) >= 36 .and. Len(Alltrim(SE2->E2_X_CODBA)) <= 43
	cValCamp := Alltrim(Substr(SE2->E2_X_CODBA,34,14))
Else
	cValCamp := "00000000000000"                 
EndIf	

cValCamp := Strzero(Val(cValCamp),14)

Return cValCamp