PROCEDURE MakeOtdel
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ ПОСТРОИТЬ ФАЙЛ OTDEL?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 

 IF fso.FileExists(tc_basedir+'\Otdel.dbf')
  IF MESSAGEBOX(CHR(13)+CHR(10)+'ФАЙЛ OTDEL УЖЕ СФОРМИРОВАН!'+CHR(13)+CHR(10)+;
   'ВЫ ХОТИТЕ СФОРМИРОВАТЬ ЕГО ЗАНОВО?',4+32,'') = 7
    RETURN 
  ELSE 
   fso.DeleteFile(tc_basedir+'\otdel.dbf')
  ENDIF 
 ENDIF 
 
 WAIT "ПОДКЛЮЧЕНИЕ К СЕРВЕРУ БАЗЫ ДАННЫХ...." WINDOW NOWAIT 
 
 lnHndl = SQLCONNECT("mssql", "sa", "")
 IF lnHndl <= 0
  =AERROR(errarr)
  = MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'Cannot make connection')
  RETURN 
 ELSE
*  = MESSAGEBOX('Connection made', 48, 'SQL Connect Message')
 ENDIF
 
 lrslt = SQLEXEC(lnHndl, "USE PDPStdStorage")
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'Cannot select database!')
  RETURN 
 ELSE 
*  = MESSAGEBOX('Database selected!', 48, 'PDPAccStorage')
 ENDIF 

 SET NULL ON 
 CREATE TABLE &tc_basedir\otdel (uid i, puid i, iotd c(4), otd c(4), "name" c(50))
 SET NULL OFF 

 SelString01 = "select uid, code_43420 as iotd, name_65044 as name "
 SelString02 = "from t_orgunit_29211 ;
  where version_end=9223372036854775807"

 SelString = SelString01 + SelString02

 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 't_orgunit_29211')
 ELSE 
  SCAN 
   SCATTER MEMVAR 
   m.name = ALLTRIM(name)
   INSERT INTO otdel FROM MEMVAR 
  ENDSCAN 
  USE 
 ENDIF 
 WAIT CLEAR 

 SelString01 = "select uid as puid, orgunituid_09716 as uid "
 SelString02 = "from t_serviceprovider_55189  where version_end=9223372036854775807"

 SelString = SelString01 + SelString02

 lrslt = SQLEXEC(lnHndl, selstring, 'sqlrslt')
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 't_serviceprovider_55189')
 ELSE 
  INDEX ON uid TAG uid 
  SET ORDER TO uid
 ENDIF 
 
 SELECT otdel
 SET RELATION TO uid INTO sqlrslt
 REPLACE FOR !EMPTY(sqlrslt.uid) puid WITH sqlrslt.puid
 SET RELATION OFF INTO sqlrslt
 USE IN sqlrslt

 SelString01 = "select provideruid_43262 as puid, iotd_32045 as otd "
 SelString02 = "from t_serviceprovider_58836  where version_end=9223372036854775807"

 SelString = SelString01 + SelString02

 lrslt = SQLEXEC(lnHndl, selstring, 'sqlrslt')
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 't_serviceprovider_58836')
 ELSE 
  INDEX ON puid TAG puid 
  SET ORDER TO puid
 ENDIF 
 
 SELECT otdel
 SET RELATION TO puid INTO sqlrslt
 REPLACE FOR !EMPTY(sqlrslt.puid) otd WITH sqlrslt.otd
 SET RELATION OFF INTO sqlrslt
 USE IN sqlrslt

 USE IN otdel

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

 RELEASE lnhndl, lresult, ln_disconn

 MESSAGEBOX('ОБРАБОТКА ЗАВЕРШЕНА!',0+64, '')
RETURN 