my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            33120A-1
DATE:                  2019-03-29 08:43:36
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       69
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON
#------------------CONFIG PLANILHA---------------
  1.001  DISP         CERTIFIQUE-SE QUE A PLANILHA DE DADOS
  1.001  DISP         FOI DEVIDADEMENTE CRIADA COM OS PONTOS E PADRÕES
  1.001  DISP         UTILIZADOS NESTA CALIBRAÇÃO E ESTÁ SALVA COM
  1.001  DISP         O NOME "33120A" NO ENDEREÇO:
  1.001  DISP         "Z:/Software/PLANILHAS/33120A.xlt"
#---------------------CONFIG EXCEL-------------------
  1.002  MATH         xlFile = "Z:/Software/PLANILHAS/33120A.xlt"
  1.003  LIB          COM xlApp = "Excel.Application";
  1.004  LIB          xlApp.Visible = True;
  1.005  LIB          COM xlWB = xlApp.Workbooks;
  1.006  LIB          xlWB.Open(xlFile);
#------------------CONFIG WORKSHEET-------------------
  1.007  LIB          COM xlWS = xlApp.Worksheets["FREQ"];
  1.008  LIB          xlWS.Select();
#-----------------------CONFIG TEST FREQUENCY------------------------
  1.009  OPBR         DESEJA CALIBRAR FREQUÊNCIA?
  1.010  JMPT         1.012
  1.011  JMP          1.013
  1.012  CALL         33120A-2
#-----------------------CONFIG TESTE LEVEL------------------------
  1.013  OPBR         DESEJA CALIBRAR NÍVEL?
  1.014  JMPT         1.016
  1.015  JMP          1.017
  1.016  CALL         33120A-3
#-----------------------CONFIG TESTE VOLTAGE------------------------
  1.017  OPBR         DESEJA CALIBRAR TENSÃO?
  1.018  JMPT         1.020
  1.019  JMP          1.021
  1.020  CALL         33120A-5
#-----------------------CONFIG TEST REFERENCE------------------------
  1.021  OPBR         DESEJA CALIBRAR FREQUÊNCIA DE REFERÊNCIA?
  1.022  JMPT         1.024
  1.023  JMP          1.025
  1.024  CALL         33120A-4
#-----------------------END CAL------------------------
  1.025  DISP         FIM

  #1.026  MATH         L = "Filename:="
  #1.027  MATH         M = ITOC(34)&"C:/TESTE/yyyyyyyyyy.xls"&ITOC(34)

  #1.028  MATH         N = ", FileFormat:=xlExcel8"
  #1.029  MATH         O = ", Password:="&ITOC(34)&ITOC(34)
  #1.030  MATH         P = ", WriteResPassword:="&ITOC(34)&ITOC(34)
  #1.031  MATH         Q = ", ReadOnlyRecommended:=False, CreateBackup:=False"

  #1.032  MATH         SS = L&M&N&O&P&Q
  #1.033  MATH         T = "C:/TESTE/yyyyyyyyyy.xls"

  #1.034  LIB          xlWB.SaveAs(T);

  #1.036  LIB          xlApp.DisplayAlerts = False;
  #1.037  LIB          xlApp.Save();
  #1.038  LIB          xlApp.Quit();

#ActiveWorkbook.SaveAs Filename:="C:\Users\AGE9\Desktop\yyyyyyyyyy.xls"
#LIB          workbook.SaveCopyAs(@savefile);
#LIB          COM workbook = excel.Workbooks
