#include "rwmake.ch"        

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Programa  ³ CMb004   ³ Autor ³ Valmir Belchior       ³ Data ³ 14/05/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Este Rdmake compoe as Rotinas de geracao do SISPAG, arquivo³±±
±±³          ³ 341REM.PAG e 341RET.PAG                                    ³±±
±±³          ³ Calcula o layout para o DV do Codigo de Barras             ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Adaptado por Jeniffer Larrosa para a versao 7.10                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±*/
User Function CMb004()        

// Variaveis Privadas
PRIVATE cDigCamp := 0

If	Len(Alltrim(SE2->E2_X_CODBA)) == 47
	cDigCamp	:=	Substr(SE2->E2_X_CODBA,33,1)
ElseIf Len(Alltrim(SE2->E2_X_CODBA)) == 44		  // Alterado de 44 p/ 47. Comtempla o Fator de Vencimento.
	cDigCamp	:=	Substr(SE2->E2_X_CODBA,5,1)
ElseIf Len(Alltrim(SE2->E2_X_CODBA)) >= 36 .and. Len(Alltrim(SE2->E2_X_CODBA)) <= 43
	cDigCamp	:=	Substr(SE2->E2_X_CODBA,33,1)
Else
	cDigCamp	:=	"0"
Endif

Return cDigCamp