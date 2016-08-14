PROCEDURE PackAttPats
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ХОТИТЕ УПАКОВАТЬ БД ПАЦИЕНТОВ?'+CHR(13)+CHR(10),4+16,'')==7
  RETURN 
 ENDIF 

 WAIT "ПОДКЛЮЧЕНИЕ К СЕРВЕРУ БАЗЫ ДАННЫХ...." WINDOW NOWAIT 
 lnHndl = SQLCONNECT("mssql", "sa", "")
 IF lnHndl <= 0
  =AERROR(errarr)
  =MESSAGEBOX(CHR(13)+CHR(10)+'Cannot make connection'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
  RETURN 
 ELSE
*  = MESSAGEBOX('Connection made', 48, 'SQL Connect Message')
 ENDIF
 WAIT CLEAR 

 lrslt = SQLEXEC(lnHndl, "USE PDPStdStorage")
 IF lrslt<=0
  =AERROR(errarr)
  =MESSAGEBOX(CHR(13)+CHR(10)+'Cannot select database!'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
  RETURN 
 ELSE 
*  = MESSAGEBOX('Database selected!', 48, 'PDPStdStorage')
 ENDIF 

 WAIT "ОТКЛЮЧЕНИЕ ТРИГГЕРА УДАЛЕНИЯ T_ATTACHEDPATS_27258..." WINDOW NOWAIT 
 TriggerOff = "alter table t_attachedpats_27258 disable trigger trt_attachedpats_27258_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("ТРИГГЕР ОТКЛЮЧЕН!",0+64,"")
 WAIT CLEAR 

 WAIT "УПАКОВКА БАЗЫ T_ATTACHEDPATS_27258..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_ATTACHEDPATS_27258 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'НЕ УДАЕТСЯ УПАКОВАТЬ T_ADDRESS_47256'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "ВКЛЮЧЕНИЕ ТРИГГЕРА УДАЛЕНИЯ T_ATTACHEDPATS_27258..." WINDOW NOWAIT 
 TriggerOff = "alter table t_attachedpats_27258 enable trigger trt_attachedpats_27258_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("ТРИГГЕР ОТКЛЮЧЕН!",0+64,"")
 WAIT CLEAR 


 WAIT "ОТКЛЮЧЕНИЕ ОТ БАЗЫ ДАННЫХ..." WINDOW NOWAIT 
 ln_disconn = SQLDISCONNECT(lnHndl)
 DO CASE 
  CASE ln_disconn ==  1 && Соединение успешно закрыто
*   = MESSAGEBOX('Connection closed', 48, 'SQL Connect Message')
  CASE ln_disconn == -1 && Ошибка уровня соединения
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  CASE ln_disconn == -2 && Ошибка уровня окружения
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  OTHERWISE && Такого быть не должно, но все-таки...
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
 ENDCASE 
 WAIT CLEAR 

RETURN 