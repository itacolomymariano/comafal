User Function cDelArq
   Processa({fLimpWF()},"Limpando as Pastas Workflow, Aguarde ...")
   /*
   Delete File &("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\emp02\mail\workflow\archive\*.eml")
   Delete File &("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\emp02\mail\workflow\InBox\*.eml")   
   Delete File &("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\emp02\mail\workflow\Sent\*.wfm")
   Delete File &("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\process\*.val")
   Delete File &("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\temp\*.htm")
   Delete File &("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\*.log")
   */
Return
Static Function fLimpWF()
   /*
   aArqDel := Directory("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\emp02\mail\workflow\archive\*.eml")
   nTtDel  := 0
   nReg := Len(aArqDel)
   For nDel := 1 To Len(aArqDel)
      IncProc("Excluindo Arquivos da Pasta \archive ... "+Alltrim(Str(nDel))+" / "+Alltrim(Str(nReg)))
      Delete File &("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\emp02\mail\workflow\archive\"+aArqDel[nDel])
      nTtDel ++
   Next nDel

   aArqDel := Directory("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\emp02\mail\workflow\InBox\*.eml")
   nReg := Len(aArqDel)
   For nDel := 1 To Len(aArqDel)
      IncProc("Excluindo Arquivos da Pasta \InBox ... "+Alltrim(Str(nDel))+" / "+Alltrim(Str(nReg)))
      Delete File &("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\emp02\mail\workflow\InBox\"+aArqDel[nDel])
      nTtDel ++
   Next nDel

   aArqDel := Directory("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\emp02\mail\workflow\Sent\*.wfm")
   nReg := Len(aArqDel)
   For nDel := 1 To Len(aArqDel)
      IncProc("Excluindo Arquivos da Pasta \Sent ... "+Alltrim(Str(nDel))+" / "+Alltrim(Str(nReg)))
      Delete File &("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\emp02\mail\workflow\Sent\"+aArqDel[nDel])
      nTtDel ++
   Next nDel

   aArqDel := Directory("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\emp02\process\*.val")
   nReg := Len(aArqDel)
   For nDel := 1 To Len(aArqDel)
      IncProc("Excluindo Arquivos da Pasta \process ... "+Alltrim(Str(nDel))+" / "+Alltrim(Str(nReg)))
      Delete File &("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\emp02\process\"+aArqDel[nDel])
      nTtDel ++
   Next nDel

   aArqDel := Directory("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\mail\temp\*.htm")
   nReg := Len(aArqDel)
   For nDel := 1 To Len(aArqDel)
      IncProc("Excluindo Arquivos da Pasta \temp ... "+Alltrim(Str(nDel))+" / "+Alltrim(Str(nReg)))
      Delete File &("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\mail\temp\"+aArqDel[nDel])
      nTtDel ++
   Next nDel
   
   aArqDel := Directory("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\emp02\mail\workflow\*.log")
   nReg := Len(aArqDel)
   For nDel := 1 To Len(aArqDel)
      IncProc("Excluindo Arquivos da Pasta \workflow ... "+Alltrim(Str(nDel))+" / "+Alltrim(Str(nReg)))
      Delete File &("D:\Outsourcing\Clientes\COMAFAL\P10_TESTE\Protheus_Data_Padrao2\workflow\emp02\mail\workflow\"+aArqDel[nDel])
      nTtDel ++
   Next nDel
   */

   aArqDel := Directory("\workflow\emp02\mail\workflow\archive\*.eml","D")
   nTtDel  := 0
   nReg := Len(aArqDel)
   ProcRegua(nReg)
   For nDel := 1 To Len(aArqDel)
      IncProc("Excluindo Arquivos da Pasta \archive ... "+Alltrim(Str(nDel))+" / "+Alltrim(Str(nReg)))
      Delete File &("\workflow\emp02\mail\workflow\archive\"+aArqDel[nDel,1])
      nTtDel ++
   Next nDel

   aArqDel := Directory("\workflow\emp02\mail\workflow\InBox\*.eml","D")
   nReg := Len(aArqDel)
   ProcRegua(nReg)
   For nDel := 1 To Len(aArqDel)
      IncProc("Excluindo Arquivos da Pasta \InBox ... "+Alltrim(Str(nDel))+" / "+Alltrim(Str(nReg)))
      Delete File &("\workflow\emp02\mail\workflow\InBox\"+aArqDel[nDel,1])
      nTtDel ++
   Next nDel

   aArqDel := Directory("\workflow\emp02\mail\workflow\Sent\*.wfm","D")
   nReg := Len(aArqDel)
   ProcRegua(nReg)
   For nDel := 1 To Len(aArqDel)
      IncProc("Excluindo Arquivos da Pasta \Sent ... "+Alltrim(Str(nDel))+" / "+Alltrim(Str(nReg)))
      Delete File &("\workflow\emp02\mail\workflow\Sent\"+aArqDel[nDel,1])
      nTtDel ++
   Next nDel

   aArqDel := Directory("\workflow\emp02\process\*.val","D")
   nReg := Len(aArqDel)
   ProcRegua(nReg)
   For nDel := 1 To Len(aArqDel)
      IncProc("Excluindo Arquivos da Pasta \process ... "+Alltrim(Str(nDel))+" / "+Alltrim(Str(nReg)))
      Delete File &("\workflow\emp02\process\"+aArqDel[nDel,1])
      nTtDel ++
   Next nDel

   aArqDel := Directory("\workflow\mail\temp\*.html","D")
   nReg := Len(aArqDel)
   ProcRegua(nReg)
   For nDel := 1 To Len(aArqDel)
      IncProc("Excluindo Arquivos da Pasta \temp ... "+Alltrim(Str(nDel))+" / "+Alltrim(Str(nReg)))
      Delete File &("\workflow\mail\temp\"+aArqDel[nDel,1])
      nTtDel ++
   Next nDel
   
   aArqDel := Directory("\workflow\emp02\mail\workflow\*.log","D")
   nReg := Len(aArqDel)
   ProcRegua(nReg)
   For nDel := 1 To Len(aArqDel)
      IncProc("Excluindo Arquivos da Pasta \workflow ... "+Alltrim(Str(nDel))+" / "+Alltrim(Str(nReg)))
      Delete File &("\workflow\emp02\mail\workflow\"+aArqDel[nDel,1])
      nTtDel ++
   Next nDel
   
   aArqDel := Directory("\workflow\procs\*.*","D")
   nReg := Len(aArqDel)
   ProcRegua(nReg)
   For nDel := 1 To Len(aArqDel)
      IncProc("Excluindo Arquivos da Pasta \process ... "+Alltrim(Str(nDel))+" / "+Alltrim(Str(nReg)))
      Delete File &("\workflow\procs\"+aArqDel[nDel,1])
      nTtDel ++
   Next nDel

   MsgInfo("Processo Concluido com Sucesso! "+Alltrim(Str(nTtDel))+" arquivos excluidos")         
   
Return(.T.)