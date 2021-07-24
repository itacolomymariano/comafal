#Include "FiveWin.Ch"
#Include "TopConn.Ch"

/*���������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
��� Programa � SD3250E  � Autor � Larson Zordan         � Data �10.02.2004���
�������������������������������������������������������������������������Ĵ��
���Descricao � Ponto de Entrada Apos o Esorno   do SD3 (Mata250)          ���
�������������������������������������������������������������������������Ĵ��
���Uso       � COMAFAL                                                    ���
�������������������������������������������������������������������������Ĵ�� 
��� ATUALIZACOES SOFRIDAS DESDE A CONSTRUCAO INICIAL.                     ���
�������������������������������������������������������������������������Ĵ��
��� PROGRAMADOR  � DATA   �HLPDSK�  MOTIVO DA ALTERACAO                   ���
�������������������������������������������������������������������������Ĵ��
��� Larson       �10/02/04�      � Desenvolvimento incial do programa.    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
����������������������������������������������������������������������������*/
User Function SD3250E()      
Local aAreaAnt := GetArea()
Local cNumOC   := SC2->C2_NUMOC

If !l250Auto .And. !Empty(cNumOC) .And. AllTrim(Posicione("SB1",1,xFilial("SB1")+SC2->C2_PRODUTO,"B1_GRUPO")) == "20"
	SZ3->(dbSetOrder(1))
	If SZ3->(dbSeek(xFilial("SZ3")+cNumOC)) .And. SZ3->Z3_OPSBXA > 0
		RecLock("SZ3",.F.)
		Replace SZ3->Z3_OPSBXA With SZ3->Z3_OPSBXA - 1
		MsUnLock()
	EndIf
EndIf

RestArea(aAreaAnt)
Return