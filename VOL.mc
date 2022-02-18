my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            33120A-5
DATE:                  2019-03-29 09:54:24
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       129
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON
#-------------------CLEAR-----------------------
  1.001  ASK-   R D   N B            P J S U       M C X Z        A  L  T  W
#-------------------VARIÁVEIS-----------------------
  1.002  MATH         P = 0
  1.003  MATH         LP = 2
  1.004  MATH         CP = 1
  1.005  MATH         T  = 0
  1.006  MATH         L = 0
  1.007  MATH         LEV =0
  1.008  MATH         EXP = 0
  1.009  MATH         LINHA = 2
  1.010  MATH         COLUNA = 5
  1.011  MATH         TEMPO = 20
  1.012  MATH         EX = 0
  1.013  MATH         Y = 0
#-------------------CONFIG EXCEL-----------------------
  1.014  LIB          COM xlWS = xlApp.Worksheets["VOL"];
  1.015  LIB          xlWS.Select();
#-------------------CLEAR GENERATOR----------------------
  1.016  IEEE         [@13]*RST
  1.017  IEEE         *CLS
#-------------------CAL MULTIMETER-----------------------
  1.018  OPBR         DESEJA REALIZAR O AUTOCAL?
  1.019  JMPT         1.022
  1.020  JMPF         1.027
  1.021  JMP          1.094
  1.022  IEEE         [@23]RESET
  1.023  IEEE         ACAL ALL
  1.024  WAIT         -t 15:00 AUTOCAL RUN
  1.025  MATH         Y = Y + 1
  1.026  JMP          1.091
#-------------------SETUP-----------------------
  1.027  IF           Y <= 1
  1.028  IEEE         [@23]RESET
  1.029  JMP          1.091
  1.030  ENDIF
  1.031  IEEE         [@23][TERM CR]
  1.032  IEEE         FUNC ACV
  1.033  IEEE         SETACV SYNC
  1.034  IEEE         LFILTER ON
  1.035  IEEE         NDIG 8
  1.036  IEEE         RANGE AUTO
  1.037  IEEE         TARM AUTO
  1.038  IEEE         RES 0.00001
#-------------------CONFIG  Nº MEAS----------------
  1.039  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.040  MATH         A = MEM
#-----------------CONFIG POINT------------------
  1.041  DO
  1.042  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.043  LIB          PONTO = P1.Value2;
  1.044  IF           PONTO == 0
  1.045  JMP          1.094
  1.046  ENDIF
  1.047  MATH         CP = CP + 1
  1.048  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.049  LIB          TEX = T1.Value2;
  1.050  MATH         P = PONTO&TEX
 #----------------LEVEL-------------------
  1.051  MATH         CP = CP + 1
  1.052  LIB          COM L1 = xlApp.Cells[LP,CP];
  1.053  LIB          LEV = L1.Value2;
  1.054  MATH         CP = CP + 1
  1.055  LIB          COM T2 = xlApp.Cells[LP,CP];
  1.056  LIB          EXP = T2.Value2;
  1.057  MATH         L = LEV&EXP
  1.058  MATH         Z1 = CMP  (EXP,"mVrms")
  1.059  MATH         Z2 = CMP  (EXP,"Vrms")
#----------------------END-------------------------------
  1.060  IF           P == 00
  1.061  JMP          1.094
  1.062  ENDIF
#----------------------------CONFIG OUT GENERATOR--------------
  1.063  DO
  1.064  IEEE         [@13]:FREQ [V P]
  1.065  IEEE         :VOLT:UNIT VRMS
  1.066  IEEE         :VOLT:LEV [V L]
  #1.065  IEEE         OUTP ON
  1.067  WAIT         [D2000]
#------------CONFIG MULT----------------
  1.068  WAIT         -t [V TEMPO] Please Standby
  1.069  IEEE         [@23]ACV[I]
  1.070  IF           Z1 == 1
  1.071  MATH         EX = 1E-3
  1.072  ENDIF
  1.073  IF           Z2 == 1
  1.074  MATH         EX = 1E0
  1.075  ENDIF
  1.076  MATH         MEM = MEM / EX
#------------------SAVE DATE----------------
  1.077  LIB          COM selectedCell = xlApp.Cells[LINHA,COLUNA];
  1.078  LIB          selectedCell.Select();
  1.079  LIB          selectedCell.FormulaR1C1 = [MEM];
  1.080  MATH         T = T + 1
  1.081  MATH         COLUNA = COLUNA + 1
  1.082  MATH         CP = CP + 1
  1.083  UNTIL        T == A
  1.084  MATH         T  = 0
  1.085  MATH         COLUNA = 5
  1.086  MATH         LINHA = LINHA + 1
  1.087  MATH         CP = 1
  1.088  MATH         LP = LP + 1
  1.089  UNTIL        PONTO == 0
  1.090  JMP          1.094
#-------------------SETUP-----------------------
  1.091  DISP         Connect the generator to the UUT as follows:
  1.091  DISP
  1.091  DISP         [32]   Generator         to         Multimeter
  1.091  DISP         [32]
  1.091  DISP         [32]     OUTPUT -------------------> INPUT (2 WIRE)
  1.091  DISP         [32]
  1.091  DISP         [32]     GPIB MULTIMETRO 3458A = 23
  1.091  DISP         [32]     GPIB GERADOR 33120A = 13
  1.092  PIC          SETUP4
  1.093  JMP          1.031
#------------------RESET------------------
  1.094  IEEE         [@23]*RST
  1.095  IEEE         [@13]*RST
