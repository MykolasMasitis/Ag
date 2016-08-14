PROCEDURE RestoreUIDInTapDiags
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ ВОССТАНОВИТЬ UID'+;
  CHR(13)+CHR(10)+'В ТАБЛИЦЕ T_TAPDIAGNOSIS_31571?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(tc_comdir+'\tapdiagsuidet.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ФАЙЛ TAPDIAGSUIDET.DBF НЕ СОЗДАВАЛСЯ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(tc_comdir+'\lastuidet.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ФАЙЛ LASTUIDET.DBF НЕ СОЗДАВАЛСЯ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 IF OpenFile(tc_comdir+'\tapdiagsuidet', 'tapdiagsuidet', 'shar')>0
  IF USED('tapdiagsuidet')
   USE IN tapdiagsuidet
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(tc_comdir+'\lastuidet', 'lastuidet', 'shar')>0
  IF USED('tapdiagsuidet')
   USE IN tapdiagsuidet
  ENDIF 
  IF USED('lastuidet')
   USE IN lastuidet
  ENDIF 
  RETURN 
 ENDIF 

 WAIT "ПОДКЛЮЧЕНИЕ К СЕРВЕРУ БАЗЫ ДАННЫХ...." WINDOW NOWAIT 
 lnHndl = SQLCONNECT("mssql", "sa", "")
 IF lnHndl <= 0
  = MESSAGEBOX('Cannot make connection', 16, 'SQL Connect Error')
  RETURN 
 ELSE
 ENDIF
 WAIT CLEAR 
 
 WAIT "ВЫБОР БД PDPSTDSTORAGE..." WINDOW NOWAIT 
 lrslt = SQLEXEC(lnHndl, "USE PDPSTDSTORAGE")
 IF lrslt<=0
  = MESSAGEBOX('Cannot select database!', 16, 'PDPSTDSTORAGE')
 ELSE 
 ENDIF 
 WAIT CLEAR 
 
 SELECT tapdiagsuidet
 SCAN 
  SCATTER MEMVAR 
  UpdString = 'UPDATE t_tapdiagnosis_31571_uid SET next_uid=?m.next_uid WHERE row=?m.row'
  lrslt = SQLEXEC(lnHndl, UpdString)
  IF lrslt<=0
   = MESSAGEBOX('Cannot process update!', 16, '')
  ENDIF 
 ENDSCAN 
 USE
 WAIT CLEAR 

 SELECT lastuidet
 m.entity = "t_tapdiagnosis_31571"
 SCAN 
  SCATTER MEMVAR 
  IF ALLTRIM(m.entity) != "t_tapdiagnosis_31571"
   LOOP 
  ENDIF 

  UpdString = 'UPDATE last_uid SET uid=?m.uid WHERE entity=?m.entity AND spid=?m.row'
  lrslt = SQLEXEC(lnHndl, UpdString)
  IF lrslt<=0
   = MESSAGEBOX('Cannot process update!', 16, '')
  ENDIF 
 ENDSCAN 
 USE 
 WAIT CLEAR 

 WAIT "ОТКЛЮЧЕНИЕ ОТ БАЗЫ ДАННЫХ..." WINDOW NOWAIT 
 ln_disconn = SQLDISCONNECT(lnHndl)
 DO CASE 
  CASE ln_disconn ==  1 && Соединение успешно закрыто
 *  = MESSAGEBOX('Connection closed', 48, 'SQL Connect Message')
  CASE ln_disconn == -1 && Ошибка уровня соединения
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  CASE ln_disconn == -2 && Ошибка уровня окружения
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  OTHERWISE && Такого быть не должно, но все-таки...
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
 ENDCASE 

 RELEASE lnhndl, lresult, ln_disconn
 
 WAIT CLEAR 
RETURN 