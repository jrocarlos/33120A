my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            33120A-4
DATE:                  2018-06-14 10:35:04
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       2
NUMBER OF LINES:       112
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON
  1.001  DISP
  1.001  DISP         [32]   Generator         to         Counter
  1.001  DISP         [32]   REFERENCE -------------------> CHANNEL 1
  1.001  DISP         [32]
  1.001  DISP         [32]     GPIB CONTADOR 53132A = 3
  1.002  PIC          SETUP3
  1.003  LABEL        FREQUENCY-REF
  1.004  RSLT         =
# ================= Select COUNT OUTPUT for frequency. =====================
  1.005  HEAD         {FREQUENCY CHANNEL 1}
  1.006  ASK-   R D   N B            P J S U       M C X Z        A  L  T  W
#-------------------CONFIG EXCEL-----------------------
  1.007  LIB          COM xlWS = xlApp.Worksheets["REF"];
  1.008  LIB          xlWS.Select();
#-----------------CONFIG COUNT----------------
  1.009  IEEE         [@3]*RST
  1.010  IEEE         :FUNC 'FREQ 1'
  1.011  IEEE         INIT:CONT OFF
  1.012  IEEE         INP1:COUP DC
  1.013  IEEE         INP1:IMP 50
  1.014  IEEE         EVEN1:LEV:AUTO ON
  1.015  IEEE         INP1:FILT OFF
  1.016  TARGET       -m
#-------------------CONFIG  Nº MEAS----------------
  1.017  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.018  MATH         A = MEM
#-----------------CONFIG POINT------------------
  1.019  MATH         P = 0
  1.020  MATH         LP = 2
  1.021  MATH         CP = 1
  1.022  MATH         T  = 0
  1.023  MATH         LINHA = 2
  1.024  MATH         COLUNA = 3
  1.025  DO
  1.026  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.027  LIB          PONTO = P1.Value2;
  1.028  IF           PONTO == 0
  1.029  JMP          2.014
  1.030  ENDIF
  1.031  MATH         CP = CP + 1
  1.032  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.033  LIB          TEX = T1.Value2;
  1.034  MATH         P = PONTO&TEX
  1.035  MATH         EX = 0
  1.036  MATH         Z1 = CMP  (TEX,"MHz")
  1.037  MATH         Z2 = CMP  (TEX,"kHz")
  1.038  MATH         Z3 = CMP  (TEX,"Hz")
#----------------------END-------------------------------
  1.039  IF           P == 00
  1.040  JMP          2.014
  1.041  ENDIF
#----------------------END-------------------------------
  1.042  IF           PONTO < 1 && Z3 == 1
  1.043  MATH         GATE = 100
  1.044  ELSE
  1.045  MATH         GATE = 10
  1.046  ENDIF
#------------CONFIG IN COUNT----------------
  1.047  MATH         TEMPO = GATE + (GATE / 2)
  1.048  IEEE         [@3]FREQ:ARM:STOP:TIM [V GATE]
  1.049  TARGET       -m
  1.050  IEEE         INIT:CONT ON
  1.051  DO
  1.052  WAIT         -t [V TEMPO] Please Standby
  1.053  IEEE         READ:FREQ?[I]
  1.054  IF           Z1 == 1
  1.055  MATH         EX = 1E6
  1.056  ENDIF
  1.057  IF           Z2 == 1
  1.058  MATH         EX = 1E3
  1.059  ENDIF
  1.060  IF           Z3 == 1
  1.061  MATH         EX = 1E0
  1.062  ENDIF
  1.063  IF           Z1 == 1 && PONTO == 1
  1.064  MATH         EX = 1E6
  1.065  ENDIF
  1.066  IF           Z2 == 1 && PONTO == 1
  1.067  MATH         EX = 1E3
  1.068  ENDIF
  1.069  IF           Z3 == 1 && PONTO == 1
  1.070  MATH         EX = 1E0
  1.071  ENDIF
  1.072  MATH         MEM = MEM / EX
  1.073  MEMCX        0              TOL
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
#------------------RESET------------------
  2.014  IEEE         [@3]*RST
  2.015  IEEE         [@13]*RST
