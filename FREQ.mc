my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            33120A-2
DATE:                  2018-06-14 10:09:24
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       2
NUMBER OF LINES:       142
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON
  1.001  DISP         Connect the generator to the UUT as follows:
  1.001  DISP
  1.001  DISP         [32]   Generator         to         Counter
  1.001  DISP         [32]     OUTPUT -------------------> CHANNEL 1
  1.001  DISP         [32]
  1.001  DISP         [32]     GPIB CONTADOR 53132A = 3
  1.001  DISP         [32]     GPIB GERADOR 33120A = 13
  1.002  PIC          SETUP1
  1.003  LABEL        FREQUENCY-1
  1.004  RSLT         =
# ================= Select COUNT OUTPUT for frequency. =====================
  1.005  HEAD         {FREQUENCY CHANNEL 1}
  1.006  ASK-   R D   N B            P J S U       M C X Z        A  L  T  W
#------------CONFIG EXCEL---------------
  1.007  LIB          COM xlWS = xlApp.Worksheets["FREQ"];
  1.008  LIB          xlWS.Select();
#------------------CONFIG GENERATOR---------
  1.009  TARGET       -p
  1.010  RSLT         =
  1.011  IEEE         [@13]*RST
  1.012  IEEE         *CLS
 # IEEE         :VOLTage:UNIT DBM
  1.013  IEEE         :VOLTage:UNIT VRMS
  1.014  IEEE         :VOLT:LEV 1
#-----------------CONFIG COUNT----------------
  1.015  IEEE         [@3]*RST
  1.016  IEEE         :FUNC 'FREQ 1'
  1.017  IEEE         INIT:CONT OFF
  1.018  IEEE         INP1:COUP DC
  1.019  IEEE         INP1:IMP 50
  1.020  IEEE         EVEN1:LEV 0.5
  1.021  TARGET       -m
#-------------------CONFIG  Nº MEAS----------------
  1.022  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.023  MATH         A = MEM
#-----------------CONFIG POINT------------------
  1.024  MATH         P = 0
  1.025  MATH         LP = 2
  1.026  MATH         CP = 1
  1.027  MATH         T  = 0
  1.028  MATH         LINHA = 2
  1.029  MATH         COLUNA = 3
  1.030  DO
  1.031  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.032  LIB          PONTO = P1.Value2;
  1.033  IF           PONTO == 0
  1.034  JMP          2.019
  1.035  ENDIF
  1.036  MATH         CP = CP + 1
  1.037  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.038  LIB          TEX = T1.Value2;
  1.039  MATH         P = PONTO&TEX
  1.040  MATH         EX = 0
  1.041  MATH         Z1 = CMP  (TEX,"MHz")
  1.042  MATH         Z2 = CMP  (TEX,"kHz")
  1.043  MATH         Z3 = CMP  (TEX,"Hz")
#-------------------TRIGGER---------------------------
  1.044  IF           PONTO > 10 && Z3 == 1
  1.045  JMP          2.015
  1.046  ENDIF
#-------------------FILTER---------------------------
  1.047  IF           PONTO > 100 && Z2 == 1 || Z1 == 1
  1.048  JMP          2.017
  1.049  ELSE
  1.050  IEEE         INP1:FILT ON
  1.051  ENDIF
#----------------------END-------------------------------
  1.052  IF           P == 00
  1.053  JMP          2.019
  1.054  ENDIF
#----------------------GATE-------------------------------
  1.055  IF           PONTO < 1 && Z3 == 1
  1.056  MATH         GATE = 100
  1.057  ELSE
  1.058  MATH         GATE = 10
  1.059  ENDIF
#----------------------------CONFIG OUT GENERATOR--------------
  1.060  IEEE         [@13]:FREQ [V P]
  1.061  WAIT         [D2000]
#------------CONFIG IN COUNT----------------
  1.062  MATH         TEMPO = GATE + (GATE / 2)
  1.063  IEEE         [@3]FREQ:ARM:STOP:TIM [V GATE]
  1.064  TARGET       -m
  1.065  IEEE         INIT:CONT ON
  1.066  DO
  1.067  WAIT         -t [V TEMPO] Please Standby
  1.068  IEEE         READ:FREQ?[I]
  1.069  IF           Z1 == 1
  1.070  MATH         EX = 1E6
  1.071  ENDIF
  1.072  IF           Z2 == 1
  1.073  MATH         EX = 1E3
  1.074  ENDIF
  1.075  IF           Z3 == 1
  1.076  MATH         EX = 1E0
  1.077  ENDIF
  1.078  IF           Z1 == 1 && PONTO == 1
  1.079  MATH         EX = 1E6
  1.080  ENDIF
  1.081  IF           Z2 == 1 && PONTO == 1
  1.082  MATH         EX = 1E3
  1.083  ENDIF
  1.084  IF           Z3 == 1 && PONTO == 1
  1.085  MATH         EX = 1E0
  1.086  ENDIF
  1.087  MATH         MEM = MEM / EX
  1.088  MEMCX        0              TOL
  2.001  LIB          COM selectedCell = xlApp.Cells[LINHA,COLUNA];
  2.002  LIB          selectedCell.Select();
  2.003  LIB          selectedCell.FormulaR1C1 = [MEM];
  2.004  MATH         T = T + 1
  2.005  MATH         COLUNA = COLUNA + 1
  2.006  MATH         CP = CP + 1
  2.007  UNTIL        T == A
  2.008  MATH         T  = 0
  2.009  MATH         COLUNA = 3
  2.010  MATH         LINHA = LINHA + 1
  2.011  MATH         CP = 1
  2.012  MATH         LP = LP + 1
  2.013  UNTIL        PONTO == 0
  2.014  JMP          2.019
#----------------CONFIG TRIGGER-------------------
  2.015  IEEE         EVEN1:LEV:AUTO ON
  2.016  JMP          1.052
#----------------CONFIG FILTER 100k-------------------
  2.017  IEEE         INP1:FILT OFF
  2.018  JMP          1.051
#------------------RESET------------------
  2.019  IEEE         [@3]*RST
  2.020  IEEE         [@13]*RST
