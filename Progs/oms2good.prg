PROCEDURE oms2good
 IF !fso.FileExists(tc_basedir+'\Talon.dbf')
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ TALON.DBF!',0+16,'')
  RETURN 
 ENDIF 
 IF OpenFile(tc_basedir+'\Talon', 'Talon', 'shar')>0
  IF USED('talon')
   USE IN talon 
  ENDIF 
  RETURN 
 ENDIF 
 
 WAIT "РАСЧЕТ..." WINDOW NOWAIT 
 m.accsum = 0
 SELECT talon 
 SUM FOR ok s_all TO m.accsum
 USE 
 WAIT CLEAR 
 
 MESSAGEBOX(TRANSFORM(m.accsum, '99 999 999.99'),0+64,'Доступная сумма')
 RELEASE m.accsum

RETURN 