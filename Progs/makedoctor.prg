PROCEDURE MakeDoctor
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ ПОСТРОИТЬ ФАЙЛ DOCTOR?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 

 IF fso.FileExists(tc_basedir+'\Doctor.dbf')
  IF MESSAGEBOX(CHR(13)+CHR(10)+'ФАЙЛ DOCTOR УЖЕ СФОРМИРОВАН!'+CHR(13)+CHR(10)+;
   'ВЫ ХОТИТЕ СФОРМИРОВАТЬ ЕГО ЗАНОВО?',4+32,'') = 7
    RETURN 
  ELSE 
   fso.DeleteFile(tc_basedir+'\doctor.dbf')
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
 CREATE TABLE &tc_basedir\doctor (uid i, pcod c(10), doc c(6), sn_pol c(25), fam c(25), im c(20), ot c(20), dr d, ;
 w n(1), d_on d, d_off d, lgot_r l)
 SET NULL OFF 

* SelString01 = "select a.uid, a.code_47321 as doc, a.fam_45430 as fam, a.im_03922 as im, a.ot_43242 as ot, ;
  CAST(b.sex_25414 as numeric(1)) as w, b.birthday_16574 as dr, policyseries_29342 as s_pol, policynumber_43545 as n_pol "
* SelString02 = "from t_employee_22089 a right outer join t_employee_49492 b on b.employeeuid_11522=a.uid ;
  where a.version_end=9223372036854775807 and b.version_end=9223372036854775807 order by fam, im, ot"

 SelString01 = "select a.uid, a.code_47321 as doc, a.fam_45430 as fam, a.im_03922 as im, a.ot_43242 as ot, "
 SelString02 = "takeondate_51957 as d_on, dischargedate_63406 as d_off from t_employee_22089 a ;
  where a.version_end=9223372036854775807 order by fam, im, ot"

 SelString = SelString01 + SelString02

 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 't_employee_22089')
 ELSE 
  SCAN 
   SCATTER MEMVAR 
*   m.sn_pol = ALLTRIM(ALLTRIM(m.s_pol)+' '+ ALLTRIM(m.n_pol))
   INSERT INTO doctor FROM MEMVAR 
  ENDSCAN 
  USE 
*  USE IN doctor
 ENDIF 
 WAIT CLEAR 

 SelString01 = "select employeeuid_11522 as uid, CAST(b.sex_25414 as numeric(1)) as w, b.birthday_16574 as dr, ;
  policyseries_29342 as s_pol, policynumber_43545 as n_pol, privilegeprescriptionissuer_29350 as lgot_r "
 SelString02 = "from t_employee_49492 b where b.version_end=9223372036854775807"

 SelString = SelString01 + SelString02

 lrslt = SQLEXEC(lnHndl, selstring, 'sqlrslt')
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 't_employee_49492')
 ELSE 
  INDEX ON uid TAG uid 
  SET ORDER TO uid
 ENDIF 
 
 SELECT doctor
 SET RELATION TO uid INTO sqlrslt
 REPLACE FOR !EMPTY(sqlrslt.uid) sn_pol WITH ALLTRIM(ALLTRIM(sqlrslt.s_pol)+' '+ ALLTRIM(sqlrslt.n_pol)), ;
  dr WITH sqlrslt.dr, w WITH sqlrslt.w, lgot_r WITH sqlrslt.lgot_r
 SET RELATION OFF INTO sqlrslt
 USE IN sqlrslt
 USE IN doctor

* WAIT "T_EMPLOYEE_22089..." WINDOW NOWAIT 
 
* SqlCommand = 'declare curdct cursor global for ;
  select code_47321 as doc, fam_45430 as fam, im_03922 as im, ot_43242 as ot from t_employee_22089'
* lrslt = SQLEXEC(lnHndl, SqlCommand)
* IF lrslt<=0
*  =AERROR(errarr)
*  = MESSAGEBOX(ALLTRIM(errarr(2)), 16, 't_employee_22089')
* ELSE 
*  = MESSAGEBOX('Command proccessed sucessfully!', 0+64, 't_employee_22089')
* ENDIF 

* SqlCommand = 'open curdct'
* lrslt = SQLEXEC(lnHndl, SqlCommand)
* IF lrslt<=0
*  =AERROR(errarr)
*  = MESSAGEBOX(ALLTRIM(errarr(2)), 16, 't_employee_22089')
* ELSE 
*  = MESSAGEBOX('Command proccessed sucessfully!', 0+64, 't_employee_22089')
* ENDIF 

* m.doc = ''
* m.fam = ''
* m.im  = ''
* m.ot  = ''

* SqlCommand = 'declare @docvar varchar(15)'
* lrslt = SQLEXEC(lnHndl, SqlCommand)
* IF lrslt<=0
*  =AERROR(errarr)
*  = MESSAGEBOX(ALLTRIM(errarr(2)), 16, 't_employee_22089')
* ELSE 
*  = MESSAGEBOX('Command proccessed sucessfully!', 0+64, 't_employee_22089')
* ENDIF 

* SqlCommand = 'FETCH NEXT FROM GLOBAL curdct INTO @docvar, ?m.fam, ?m.im, ?m.ot'
* lrslt = SQLEXEC(lnHndl, SqlCommand)
* IF lrslt<=0
*  =AERROR(errarr)
*  = MESSAGEBOX(ALLTRIM(errarr(2)), 16, 't_employee_22089')
* ELSE 
*  = MESSAGEBOX('Command proccessed sucessfully!', 0+64, 't_employee_22089')
* ENDIF 
 
* SqlCommand = 'select @docvar as fam'
* lrslt = SQLEXEC(lnHndl, SqlCommand, 'fcurs')
* IF lrslt<=0
*  =AERROR(errarr)
*  = MESSAGEBOX(ALLTRIM(errarr(2)), 16, 't_employee_22089')
* ELSE 
*  = MESSAGEBOX('@docvar = '+fam, 0+64, 't_employee_22089')
*  m.fam = fam
*  USE 
* ENDIF 

* m.fvar = -1
* SqlCommand = 'select @@FETCH_STATUS as fstat'
* lrslt = SQLEXEC(lnHndl, SqlCommand, 'fcurs')
* IF lrslt<=0
*  =AERROR(errarr)
*  = MESSAGEBOX(ALLTRIM(errarr(2)), 16, 't_employee_22089')
* ELSE 
*  = MESSAGEBOX('@@FETCH_STATUS = '+STR(fstat,2), 0+64, 't_employee_22089')
*  m.fvar = fstat
*  USE 
* ENDIF 

* MESSAGEBOX(m.fam,0+64,'')

 IF 3=2
 DO WHILE m.fvar==0

  m.doc = ''
  m.fam = ''
  m.im  = ''
  m.ot  = ''

  SqlCommand = 'FETCH NEXT FROM GLOBAL curdct INTO ?m.doc, ?m.fam, ?m.im, ?m.ot'
  lrslt = SQLEXEC(lnHndl, SqlCommand)
  IF lrslt<=0
   =AERROR(errarr)
   = MESSAGEBOX(ALLTRIM(errarr(2)), 16, 't_employee_22089')
   EXIT 
  ELSE 
 *  = MESSAGEBOX('Command proccessed sucessfully!', 0+64, 't_employee_22089')
  ENDIF 

  SqlCommand = 'select @@FETCH_STATUS as fstat'
  lrslt = SQLEXEC(lnHndl, SqlCommand, 'fcurs')
  IF lrslt<=0
   =AERROR(errarr)
   = MESSAGEBOX(ALLTRIM(errarr(2)), 16, 't_employee_22089')
  ELSE 
   = MESSAGEBOX('@@FETCH_STATUS = '+STR(fstat,2), 0+64, 't_employee_22089')
   m.fvar = fstat
   USE 
  ENDIF 
 
  MESSAGEBOX(m.fam,0+64,'')
 
 ENDDO 
 ENDIF 

* SqlCommand = 'close curdct'
* lrslt = SQLEXEC(lnHndl, SqlCommand)
* IF lrslt<=0
*  =AERROR(errarr)
*  = MESSAGEBOX(ALLTRIM(errarr(2)), 16, 't_employee_22089')
* ELSE 
*  = MESSAGEBOX('Command proccessed sucessfully!', 0+64, 't_employee_22089')
* ENDIF 

* SqlCommand = 'deallocate curdct'
* lrslt = SQLEXEC(lnHndl, SqlCommand)
* IF lrslt<=0
*  =AERROR(errarr)
*  = MESSAGEBOX(ALLTRIM(errarr(2)), 16, 't_employee_22089')
* ELSE 
*  = MESSAGEBOX('Command proccessed sucessfully!', 0+64, 't_employee_22089')
* ENDIF 


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