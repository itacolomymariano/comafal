User Function VldCPagto

   Local lRVldCP := .T.
   
   //Restrito a Condi??es iniciadas com a letra "V"
   If UPPER(Substr(M->E4_CODIGO,1,1)) == "V"
   
      Private nAutoriz  := GetMV("MV_PRZMED")
      Private cCondicao := Alltrim(M->E4_COND)
      Private cCondClc  := STRTRAN(cCondicao,",","+")
      Private nQtdPc    := RunPc(cCondicao)
      Private nPrzTot   := &cCondClc
      Private nMedPrazo := (nPrzTot / nQtdPc)

      If nMedPrazo > nAutoriz
         lRVldCP := .F.
         Alert("O Prazo M?dio desta Condi??o de Pagamento, "+Alltrim(Str(nMedPrazo))+" dias,"+;
               " ultrapassa o Prazo M?dio autorizado pela Diretoria que ? de "+Alltrim(Str(nAutoriz))+" dias,"+;
               "informe uma Condi??o de Pagamento com Prazo M?dio igual ou inferior ao autorizado.")
      EndIf
   
   EndIf

Return(lRVldCP)

User Function MT360VLD

   //Restrito a Condi??es iniciadas com a letra "V"
   If UPPER(Substr(M->E4_CODIGO,1,1)) == "V"
   
      Private nAutoriz  := GetMV("MV_PRZMED")
      Private cCondicao := Alltrim(M->E4_COND)
      Private cCondClc  := STRTRAN(cCondicao,",","+")
      Private nQtdPc    := RunPc(cCondicao)
      Private nPrzTot   := &cCondClc
      Private nMedPrazo := (nPrzTot / nQtdPc)

      If nMedPrazo > nAutoriz
         lRVldCP := .F.
         Alert("O Prazo M?dio desta Condi??o de Pagamento, "+Alltrim(Str(nMedPrazo))+" dias,"+;
               " ultrapassa o Prazo M?dio autorizado pela Diretoria que ? de "+Alltrim(Str(nAutoriz))+" dias,"+;
               "informe uma Condi??o de Pagamento com Prazo M?dio igual ou inferior ao autorizado.")
      EndIf
   
   EndIf

Return(lRVldCP)

Static Function RunPc(cCondicao)
   
   nQuantPC := 1//J? iniciamos considerando o m?nimo de uma parcela, ent?o se existir a primeira v?rgula
                //isto indica que temos duas parcelas, e se for encontrada uma pr?xima v?rgula temos tr?s
                // e assim por diante contaremos o n?mero de parcelas de acordo com o n?mero de v?rgulas
                // encontradas.
   nIncio  := 1
   For nPc := 1 To Len(cCondicao)
      cTexto := Substr(IIF(nPC==1,cCondicao,cTexto),nIncio,Len(cCondicao)) 
      //Alert("cTexto "+cTexto+" nIncio "+Alltrim(Str(nIncio)))
      nPosVirg := At(",",cTexto)
      //Alert("nPosVirg "+Alltrim(Str(nPosVirg))+" nPC "+Alltrim(Str(nPC))+" nQuantPC "+Alltrim(Str(nQuantPC)))
      nIncio  := nPosVirg + 1
      If nPosVirg > 0
         nQuantPC ++
      EndIf
   Next nPC
   
Return(nQuantPC)