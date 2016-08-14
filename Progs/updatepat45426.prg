PROCEDURE UpdatePat45426
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

 SelString = "select CAST(version_begin as numeric(20)) as v_beg, CAST(version_end as numeric(20)) as v_end, uid as uid ;
  from t_patient_10905 where version_end=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process select!', 16, 't_patient_10905')
 ELSE 

  WAIT "ОБНУЛЕНИЕ ТАБЛИЦЫ DBO.T_PATIENT_45426..." WINDOW NOWAIT 
  DelString = "truncate table t_patient_45426"
  lrslt = SQLEXEC(lnHndl, delstring)
  IF lrslt<=0
   MESSAGEBOX('Cannot process delete!', 16, 't_patient_45426')
  ELSE 
   = MESSAGEBOX('Records deleted!', 48, 't_patient_45426')
  ENDIF 
  WAIT CLEAR 
 
  SELECT sqlresult
  BROWSE 
  SCAN 
   WAIT STR(RECNO(),6)+'/'+STR(RECCOUNT(),6) WINDOW NOWAIT 
   m.version_begin = v_beg
   m.version_end = '9223372036854775807'
   m.is_top = 1
   m.created_by = 2

   m.deleted_by = .NULL.
   m.patientuid_29790 = uid
   m.checkvector_05009 = .NULL.
   m.checkdate_58618 = .NULL.
   m.speccase_10311 = '0'

   UpdateString01 = "insert into t_patient_45426"
   UpdateString02 = "(version_begin,version_end,is_top,created_by,patientuid_29790,speccase_10311) "
   UpdateString03 =  "values(?m.version_begin,?m.version_end,?m.is_top,?m.created_by,?m.patientuid_29790,?m.speccase_10311)"

   UpdateString = UpdateString01 + UpdateString02 + UpdateString03
   lrslt = SQLEXEC(lnHndl, UpdateString)
   IF lrslt<=0
    MESSAGEBOX('Cannot insert record!'+STR(m.version_end), 16, 't_patient_45426')
    EXIT 
   ENDIF 
  ENDSCAN 
  USE 
 ENDIF 

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