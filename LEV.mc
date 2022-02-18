my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            33120A-3
DATE:                  2018-06-18 08:57:35
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       124
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON
  1.001  LABEL        FREQUENCY-1
  1.002  RSLT         =
# ================= Select COUNT OUTPUT for frequency. =====================
  1.003  HEAD         {FREQUENCY CHANNEL 1}
  1.004  ASK-   R D   N B            P J S U       M C X Z        A  L  T  W
#------------CONFIG EXCEL---------------
  1.005  LIB          COM xlWS = xlApp.Worksheets["LEV"];
  1.006  LIB          xlWS.Select();
#---------------------------PERDA POR INSERÇÃO----------------------
  1.007  MEMI         INSIRA O VALOR DE PERDA DO CONECTOR
  1.008  MATH         MEM2 = MEM
#----------------------------CONFIG GENERATOR----------------------
  1.009  RSLT         =
  1.010  IEEE         [@13]*RST
  1.011  IEEE         *CLS
  1.012  IEEE         :VOLTage:UNIT DBM
#---------------------------CONFIG METER-----------------------
  1.013  IEEE         [@21]*RST
  1.014  IEEE         SYST:PRES
  #1.017  IEEE         CAL:ZERO:AUTO ONCE
  #1.018  IEEE         CAL:ALL:ZERO:FAST:AUTO
  #1.019  IEEE         SYST:ERR?
  #1.020  IEEE         SENS:AVER:COUN:AUTO ON
  #1.021  IEEE         SENS:AVER:COUN 16
  #1.022  IEEE         OUTP:ROSC ON
#-----------------ZERO METER----------------
  1.015  DISP         CONECTE O SENSOR NA PORTA "POWER REF"
  1.015  DISP
  1.015  DISP         [32]   POWER METER         to         SENSOR
  1.015  DISP         [32]
  1.015  DISP         [32]
  1.015  DISP         [32]   POWER REF -------------------> SENSOR
  1.015  DISP         [32]
  1.015  DISP         [32]     GPIB POWER METER NRP2 = 21
  1.015  DISP         [32]     GPIB GERADOR 33120A = 13
  1.015  DISP         [32]
  1.016  PIC          SETUP2-1
  1.017  IEEE         SENS1:FREQ:CW 50 MHZ
  1.018  IEEE         INIT:CONT ON
  1.019  IEEE         CAL1:ZERO:AUTO ONCE
  1.020  WAIT         [D7000]
  1.021  IEEE         OUTP:ROSC ON
  1.022  WAIT         [D8000]
  1.023  IEEE         OUTP:ROSC OFF
  1.024  DISP         Connect the generator to the UUT as follows:
  1.024  DISP
  1.024  DISP         [32]   Generator         to         Meter
  1.024  DISP         [32]   OUTPUT -------------------> POWER SENSOR
  1.024  DISP         [32]
  1.025  PIC          SETUP2-2
#-------------------CONFIG  Nº MEAS----------------
  1.026  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.027  MATH         A = MEM
#-----------------CONFIG POINT------------------
  1.028  MATH         P = 0
  1.029  MATH         LP = 2
  1.030  MATH         CP = 1
  1.031  MATH         T  = 0
  1.032  MATH         L = 0
  1.033  MATH         LINHA = 2
  1.034  MATH         COLUNA = 5
  1.035  MATH         CPERDA = COLUNA + A
  1.036  DO
  1.037  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.038  LIB          PONTO = P1.Value2;
  1.039  IF           PONTO == 0
  1.040  JMP          1.078
  1.041  ENDIF
  1.042  MATH         CP = CP + 1
  1.043  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.044  LIB          TEX = T1.Value2;
  1.045  MATH         P = PONTO&TEX
#----------------LEVEL-------------------
  1.046  MATH         CP = CP + 1
  1.047  LIB          COM L1 = xlApp.Cells[LP,CP];
  1.048  LIB          L = L1.Value2;
#----------------------END-------------------------------
  1.049  IF           P == 00
  1.050  JMP          1.078
  1.051  ENDIF
#----------------------------CONFIG OUT GENERATOR--------------
  1.052  IEEE         [@13]:FREQ [V P]
  1.053  IEEE         :VOLT:LEV [V L]
  1.054  WAIT         [D2000]
#------------CONFIG IN COUNT----------------
  1.055  MATH         TEMPO = 5
  1.056  IEEE         [@21]SENS1:FREQ:CW  [V P]
  1.057  IEEE         INIT:CONT ON
  1.058  DO
  1.059  WAIT         -t [V TEMPO] Please Standby
  1.060  IEEE         READ?[I]
#------------------SAVE DATE----------------
  1.061  LIB          COM selectedCell = xlApp.Cells[LINHA,COLUNA];
  1.062  LIB          selectedCell.Select();
  1.063  LIB          selectedCell.FormulaR1C1 = [MEM];
  1.064  MATH         T = T + 1
  1.065  MATH         COLUNA = COLUNA + 1
  1.066  MATH         CP = CP + 1
  1.067  UNTIL        T == A
#---------------------SAVE LOSS---------------------
  1.068  LIB          COM selectedCell2 = xlApp.Cells[LINHA,CPERDA];
  1.069  LIB          selectedCell2.Select();
  1.070  LIB          selectedCell2.Value2 = [MEM2];
  1.071  MATH         T  = 0
  1.072  MATH         COLUNA = 5
  1.073  MATH         LINHA = LINHA + 1
  1.074  MATH         CP = 1
  1.075  MATH         LP = LP + 1
  1.076  UNTIL        PONTO == 0
  1.077  JMP          1.078
#------------------RESET------------------
  1.078  IEEE         [@21]*RST
  1.079  IEEE         [@13]*RST
