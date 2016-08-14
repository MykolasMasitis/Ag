PROCEDURE KillCards
 SET NULL OFF 
 
 WAIT "ПОДКЛЮЧЕНИЕ К СЕРВЕРУ БАЗЫ ДАННЫХ...." WINDOW NOWAIT 
 lnHndl = SQLCONNECT("mssql", "sa", "")
 IF lnHndl <= 0
  = MESSAGEBOX('Cannot make connection', 16, 'SQL Connect Error')
  RETURN 
 ELSE
 ENDIF
 WAIT CLEAR 
 
 m.DBName = "PDPStdStorage"
 lrslt = SQLEXEC(lnHndl, "USE "+m.DBName)
 IF lrslt<=0
  = MESSAGEBOX('Cannot select database!', 16, m.DBName)
 ELSE 
 ENDIF 

 WAIT "ОБНУЛЕНИЕ ТАБЛИЦЫ DBO.T_BOOK_54341..." WINDOW NOWAIT 
 DelString = "truncate table t_book_54341"
 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_book_54341')
 ELSE 
*  = MESSAGEBOX('Records deleted!', 48, 't_book_54341')
 ENDIF 
 WAIT CLEAR 

 WAIT "ОТКЛЮЧЕНИЕ ТРИГГЕРА ОБНОВЛЕНИЯ T_TAP_10874..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tap_10874 disable trigger trt_tap_10874_update"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("ТРИГГЕР ОТКЛЮЧЕН!",0+64,"")
 WAIT CLEAR 

 WAIT "ОТКЛЮЧЕНИЕ ТРИГГЕРА УДАЛЕНИЯ T_BOOK_65067..." WINDOW NOWAIT 
 TriggerOff = "alter table t_book_65067 disable trigger trt_book_65067_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("ТРИГГЕР ОТКЛЮЧЕН!",0+64,"")
 WAIT CLEAR 

 WAIT "УДАЛЕНИЕ НЕИСПОЛЬЗУЕМЫХ КАРТ..." WINDOW NOWAIT 
 SelString = "select patientuid_32736 as uid, coderegion_37290 as c_t from t_address_47256 where version_end=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
*  MESSAGEBOX('Cannot process select!', 16, 't_address_47256')
 ELSE 
  SELECT sqlresult
  m.uid      = 0
  SCAN FOR INT(VAL(c_t))=77
   m.uid=uid

   UpdateString = "update t_tap_10874 set bookuid_60769=NULL where patientuid_40511=?m.uid and version_end=9223372036854775807"
   lrslt = SQLEXEC(lnHndl, UpdateString)
   IF lrslt<=0
*    MESSAGEBOX('Cannot update record!', 16, 't_tap_10874')
   ENDIF 

   DeleteString = "delete t_book_65067 where patientuid_37756=?m.uid and version_end=9223372036854775807"
   lrslt = SQLEXEC(lnHndl, DeleteString)
   IF lrslt<=0
*    MESSAGEBOX('Cannot delete record!', 16, 't_book_65067')
   ENDIF 
  ENDSCAN 
  USE 
 ENDIF 
 WAIT CLEAR 

 WAIT "ВКЛЮЧЕНИЕ ТРИГГЕРА ОБНОВЛЕНИЯ T_TAP_10874..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tap_10874 enable trigger trt_tap_10874_update"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("ТРИГГЕР ОТКЛЮЧЕН!",0+64,"")
 WAIT CLEAR 

 WAIT "ВКЛЮЧЕНИЕ ТРИГГЕРА УДАЛЕНИЯ T_BOOK_65067..." WINDOW NOWAIT 
 TriggerOff = "alter table t_book_65067 enable trigger trt_book_65067_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("ТРИГГЕР ОТКЛЮЧЕН!",0+64,"")
 WAIT CLEAR 

* WAIT "ОБНУЛЕНИЕ ТАБЛИЦЫ DBO.T_BOOK_65067..." WINDOW NOWAIT 
* DelString = "truncate table t_book_65067"

* lrslt = SQLEXEC(lnHndl, delstring)
* IF lrslt<=0
*  MESSAGEBOX('Cannot process delete!', 16, 't_book_65067')
* ELSE 
*  = MESSAGEBOX('Records deleted!', 48, 't_book_65067')
* ENDIF 
* WAIT CLEAR 

* SelString = "select * from t_book_65067_uid"
* lrslt = SQLEXEC(lnHndl, selstring)
* IF lrslt<=0
*  MESSAGEBOX('Cannot process select!', 16, 't_book_65067_uid')
* ELSE 
*  SELECT sqlresult
*  m.row      = 0
*  m.next_uid = 0
*  SCAN 
*   m.row = row
*   m.next_uid = next_uid
*   m.newuid = m.row + 1
*   IF m.next_uid != m.newuid
*    UpdateString = "update t_book_65067_uid set next_uid=?m.newuid where row=?m.row"
*    lrslt = SQLEXEC(lnHndl, UpdateString)
*    IF lrslt<=0
*     MESSAGEBOX('Cannot update record!', 16, 't_book_65067_uid')
*    ENDIF 
*   ENDIF 
*  ENDSCAN 
*  USE 
* ENDIF 

 WAIT "ОТКЛЮЧЕНИЕ ОТ СЕРВЕРА..." WINDOW NOWAIT 
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
 WAIT CLEAR 

 RELEASE lnhndl, lresult, ln_disconn
 
RETURN 