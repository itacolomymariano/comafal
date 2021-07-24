#include "rwmake.ch"        

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Programa  ³ CMB005   ³ Autor ³ Valmir Belchior       ³ Data ³ 14/05/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Este Rdmake compoe as Rotinas de geracao do SISPAG, arquivo³±±
±±³          ³ 341REM.PAG e 341RET.PAG                                    ³±±
±±³          ³ Calcula o layout para o Campo Livre de Dados do Codigo de  ³±±
±±³          ³ Barras.                                                    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Adaptado por Jeniffer Larrosa para a versao 7.10                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±*/
User Function CMB005()     

// Variaveis Private
PRIVATE cCampFree

If	Len(Alltrim(SE2->E2_X_CODBA)) == 44
	cCampFree := Substr(SE2->E2_X_CODBA,20,25)
ElseIF Len(Alltrim(SE2->E2_X_CODBA)) == 47
	cCampFree := Substr(SE2->E2_X_CODBA,5,5) + Substr(SE2->E2_X_CODBA,11,10) + Substr(SE2->E2_X_CODBA,22,10)
ElseIf Len(Alltrim(SE2->E2_X_CODBA)) >= 36 .AND. Len(Alltrim(SE2->E2_X_CODBA)) <= 40
	cCampFree := Substr(SE2->E2_X_CODBA,5,5) + Substr(SE2->E2_X_CODBA,11,10) + Substr(SE2->E2_X_CODBA,22,10)
Else
	cCampFree := Replicate("0",25)
EndIf

Return cCampFree