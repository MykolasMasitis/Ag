PROCEDURE CorUIDInTap10874
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ ПРОВЕРИТЬ UID'+;
  CHR(13)+CHR(10)+'В ТАБЛИЦЕ T_TAP_10874?'+CHR(13)+CHR(10),4+32,'')==7
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

 WAIT "..." WINDOW NOWAIT 
 SelString = "SELECT TOP 100 uid FROM t_tap_10874 ORDER BY uid DESC"

 lrslt = SQLEXEC(lnHndl, SelString)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, '')
 ELSE 
  CREATE TABLE &tc_comdir\tapuid ("row" i, uid i, next_uid i) 
  USE 
  =OpenFile(tc_comdir+'\tapuid', 'tapuid', 'excl')
  SELECT sqlresult
  BROWSE 
  SCAN 
   m.uid    = uid 
   m.struid = ALLTRIM(STR(m.uid))
   m.spid   = INT(VAL(RIGHT(m.struid,2)))
   m.row = m.spid - 1
   m.next_uid = INT(VAL(ALLTRIM(STR(INT(VAL(SUBSTR(m.struid,1,LEN(m.struid)-2)))+1))+RIGHT(m.struid,2)))
   
   IF m.spid<=0
    LOOP 
   ENDIF 

   INSERT INTO tapuid FROM MEMVAR 

  ENDSCAN 
  USE 
  SELECT tapuid
  INDEX ON row TAG row
  SET ORDER TO row
  COPY TO &tc_comdir\ttt
  SET ORDER TO 
  DELETE TAG ALL 
  ZAP 
  APPEND FROM &tc_comdir\ttt

  m.entity = "t_tap_10874"
  SCAN 
   SCATTER MEMVAR 
   UpdString = 'UPDATE t_tap_10874_uid SET next_uid=?m.next_uid WHERE row=?m.row'
   lrslt = SQLEXEC(lnHndl, UpdString)
   IF lrslt<=0
    = MESSAGEBOX('Cannot process update!', 16, '')
   ENDIF 

   UpdString = 'UPDATE last_uid SET uid=?m.uid WHERE entity=?m.entity AND spid=?m.row'
   lrslt = SQLEXEC(lnHndl, UpdString)
   IF lrslt<=0
    = MESSAGEBOX('Cannot process update!', 16, '')
   ENDIF 

  ENDSCAN 
  USE
  fso.DeleteFile(tc_comdir+'\ttt.dbf')
  
 ENDIF 
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