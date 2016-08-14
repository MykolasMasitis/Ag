PROCEDURE imp_sql2fox

 SET DEFAULT TO d:\vkms.sql
 
 lnHndl = SQLCONNECT("sqlsrv", "sa", "")
 IF lnHndl <= 0
  = MESSAGEBOX('Cannot make connection', 16, 'SQL Connect Error')
  RETURN 
 ELSE
  = MESSAGEBOX('Connection made', 48, 'SQL Connect Message')
 ENDIF

* lnResult = SQLEXEC(lnHndl, "select db_id('kmss') as aaa", "rslt")

 IF !FILE('d:\vkms.sql\base\kms.mdf')
  = MESSAGEBOX('Попытка создать БД', 48, 'SQL Connect Message')
  lnResult = SQLEXEC(lnHndl, 'create database kms on (name="kms", filename="d:\vkms.sql\base\kms.mdf") ;
   log on (name="kms_log", filename="d:\vkms.sql\base\kms.ldf")')
  = MESSAGEBOX(ALLTRIM(STR(lnresult)), 48, 'SQL Connect Message')
  lnResult = SQLEXEC(lnHndl, "create table people (fam varchar(25), im varchar(20), ot varchar(20))")
  = MESSAGEBOX(ALLTRIM(STR(lnresult)), 48, 'SQL Connect Message')
 ENDIF 
 
 =SQLEXEC(lnHndl, 'use kms')
 
 USE base\kms IN 0 ALIAS kms SHARED 
 USE base\speed IN 0 ALIAS speed excl 
 USE base\kms_test IN 0 ALIAS kms_test EXCLUSIVE 
 SELECT kms_test 
 ZAP 
 SELECT speed
 ZAP

 SELECT kms
 WAIT "Конвертация, ждите..." WINDOW NOWAIT 
 SCAN 
  m.fam = fam
  m.im  = im
  m.ot  = ot
  
  m.t_beg  = 0
  m.t_end  = 0
  m.t_cont = 0
  
  m.t_beg = SECONDS()
  =SQLEXEC(lnHndl, 'insert into people (fam, im, ot) values (?m.fam, ?m.im, ?m.ot)')
  m.t_end = SECONDS()
  m.t_cont_sql = m.t_end - m.t_beg
  
  m.t_beg = SECONDS()
  INSERT INTO kms_test (fam, im, ot) VALUES (m.fam, m.im, m.ot)
  m.t_end = SECONDS()
  m.t_cont_dbf = m.t_end - m.t_beg
  
  INSERT INTO speed (exp_sql, exp_dbf) VALUES (m.t_cont_sql, m.t_cont_dbf)
 
 ENDSCAN 
 WAIT CLEAR 
 
 USE 
  
 ln_disconn = SQLDISCONNECT(lnHndl)
 DO CASE 
  CASE ln_disconn ==  1 && Соединение успешно закрыто
   = MESSAGEBOX('Connection closed', 48, 'SQL Connect Message')
  CASE ln_disconn == -1 && Ошибка уровня соединения
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  CASE ln_disconn == -2 && Ошибка уровня окружения
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  OTHERWISE && Такого быть не должно, но все-таки...
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
 ENDCASE 
 
 RELEASE lnhndl, lnresult, ln_disconn

RETURN 