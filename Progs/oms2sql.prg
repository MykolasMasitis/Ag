PROCEDURE oms2sql

 m.d_beg = m.tdat1
 m.d_end = m.tdat2
 m.choice = .f.
 DO FORM SelectPeriod
 IF m.choice!=.t.
  RETURN 
 ENDIF 
 m.cd_beg = STR(YEAR(m.d_beg),4)+PADL(MONTH(m.d_beg),2,'0')+PADL(DAY(m.d_beg),2,'0')
 m.cd_end = STR(YEAR(m.d_end),4)+PADL(MONTH(m.d_end),2,'0')+PADL(DAY(m.d_end),2,'0')
 
 SET NULL OFF 
 
 WAIT "ПОДКЛЮЧЕНИЕ К СЕРВЕРУ БАЗЫ ДАННЫХ...." WINDOW NOWAIT 
 
 lnHndl = SQLCONNECT("mssql", "sa", "")
 IF lnHndl <= 0
  = MESSAGEBOX('Cannot make connection', 16, 'SQL Connect Error')
  RETURN 
 ELSE
*  = MESSAGEBOX('Connection made', 48, 'SQL Connect Message')
 ENDIF
 
 
 m.DataBaseName = "PDPStdStorage"
 m.TableName = "t_tapservice_61560"
 m.OutTableName = "tapservice_61560"
 
 lrslt = SQLEXEC(lnHndl, "USE "+m.DataBaseName)
 IF lrslt<=0
  = MESSAGEBOX('Cannot select database!', 16, m.DataBaseName)
 ELSE 
*  = MESSAGEBOX('Database selected!', 48, m.DataBaseName)
 ENDIF 

 WAIT m.TableName WINDOW NOWAIT 

 SelString01 = "select CAST(substring(serv.servicecode_20924,3,6) as numeric(6)) as cod, CAST(SUM(serv.count_23546) as numeric(6)) as k_u from t_tapservice_61560 serv "
 SelString02 = "join t_tapdiagnosis_31571 diags on diags.uid=serv.diagnosisuid_34765 "
 SelString03 = "join t_tap_10874 tap on tap.uid=diags.tapuid_30432 "
 SelString04 = "where serv.version_end=9223372036854775807 and tap.medicalprogramm_28647='ОМСМК' and tap.fillingdate_36966 between ?m.cd_beg and ?m.cd_end "
 SelString05 = "group by serv.servicecode_20924 order by cod"

 SelString = SelString01 + SelString02 + SelString03 + SelString04 + SelString05

 lrslt = SQLEXEC(lnHndl, SelString)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, m.TableName)
 ELSE 
*  = MESSAGEBOX('Records selected!', 48, m.TableName)
*  COPY TO tc_basedir + '\' + m.OutTableName
  =OpenFile(tc_comdir+'\tarifn', 'tarif', 'shar', 'cod')
  SELECT sqlresult
  SET RELATION TO cod INTO tarif
  m.sumacc = 0
  SCAN 
   m.sumacc = m.sumacc + k_u*tarif.tarif
  ENDSCAN 
  SET RELATION OFF INTO tarif
  USE 
  USE IN tarif
  
  MESSAGEBOX(TRANSFORM(m.sumacc, '99 999 999.99'),0+64,'Сумма')
  
  RELEASE m.sumacc
 ENDIF 

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

RELEASE lnhndl, lresult, ln_disconn

WAIT CLEAR 
 
RETURN 