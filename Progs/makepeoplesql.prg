PROCEDURE MakePeopleSQL

 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ ПОСТРОИТЬ ФАЙЛ РЕГИСТРА'+;
  CHR(13)+CHR(10)+'ИЗ MS SQL?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 SET NULL OFF 

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
 
 WAIT "T_PATIENT_10905..." WINDOW NOWAIT 
 SelString01 = "select uid, CAST(fam_18565 as char(25)) as fam, ;
  CAST(im_53316 as char(20))as im, CAST(ot_48206 as char(20)) as ot,;
  birthday_38523 as dr, sex_40994 as w, registrationdate_10831 as regdate, updatedate_30831 as upddate "
 SelString02 = "from t_patient_10905 where version_end=9223372036854775807 order by uid"
 SelString = SelString01 + SelString02

 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 't_patient_10905')
 ELSE 
  COPY TO &tc_comdir\t_patient_10905
  USE 
 ENDIF 
 WAIT CLEAR 
 
* WAIT "T_PATIENT_45426..." WINDOW NOWAIT 
* SelString01 = "select patientuid_29790 as patuid, checkvector_05009 as verz, checkdate_58618 as derz, speccase_10311 as d_type "
* SelString02 = "from t_patient_45426 where version_end=9223372036854775807"
* SelString = SelString01 + SelString02

* lrslt = SQLEXEC(lnHndl, selstring)
* IF lrslt<=0
*  = MESSAGEBOX('Cannot process select!', 16, 't_patient_45426')
* ELSE 
*  COPY TO &tc_comdir\t_patient_45426
*  USE 
* ENDIF 
* WAIT CLEAR 

* WAIT "T_PATIENTADDRESS_56626..." WINDOW NOWAIT 
* SelString01 = "select addressuid_61183 as arduid, codemskstreet_40382 as ul "
* SelString02 = "from t_patientaddress_56626 where version_end=9223372036854775807"
* SelString = SelString01 + SelString02

* lrslt = SQLEXEC(lnHndl, selstring)
* IF lrslt<=0
*  = MESSAGEBOX('Cannot process select!', 16, 't_patientaddress_56626')
* ELSE 
*  COPY TO &tc_comdir\t_patientaddress_56626
*  USE 
* ENDIF 
* WAIT CLEAR 

 WAIT "T_POLICY_43176 (Policy)..." WINDOW NOWAIT 
 && Patient_Policy
 SelString = "select uid, patientuid_09882 as patientuid, smouid_25976 as smouid, series_14820 as s_pol,;
  number_12574 as n_pol, state_19333 as state, type_07731 as type;
  from t_policy_43176 where version_end=9223372036854775807 order by patientuid"

 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 't_policy_43176')
 ELSE 
*  = MESSAGEBOX('Records selected!', 48, 't_policy_43176')
  COPY TO &tc_comdir\t_policy_43176
  USE 
 ENDIF 
 WAIT CLEAR 
 
* WAIT "T_POLICY_53495..." WINDOW NOWAIT 
* && Patient_Policy
* SelString = "select policyuid_26734 as poluid, checkvector_17389 as verz, checkdate_08921 as derz ;
*  from t_policy_53495 where version_end=9223372036854775807"

* lrslt = SQLEXEC(lnHndl, selstring)
* IF lrslt<=0
*  = MESSAGEBOX('Cannot process select!', 16, 't_policy_53495')
* ELSE 
**  = MESSAGEBOX('Records selected!', 48, 't_policy_53495')
*  COPY TO &tc_comdir\t_policy_53495
*  USE 
* ENDIF 
* WAIT CLEAR 

 WAIT "T_ADDRESS_47256 (Address)..." WINDOW NOWAIT 

 && Patient_Address
 SelString01 = "select uid,patientuid_32736 as patientuid,coderegion_37290 as c_t,codestreet_11408 as ul,"
 SelString02 = "housenumber_07908 as dom, buildingnumber_35985 as kor,;
  unitnumber_64212 as str,flatnumber_60133 as kv from t_address_47256 where version_end=9223372036854775807 order by patientuid"
 SelString = SelString01 + SelString02

 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 't_address_47256')
 ELSE 
 * = MESSAGEBOX('Records selected!', 48, 't_address_47256')
  COPY TO &tc_comdir\t_address_47256
  USE 
 ENDIF 

 WAIT CLEAR 
 
 WAIT "T_BOOK_65067  (EHR_Book)..." WINDOW NOWAIT 
 && EHR_Book
 SelString = "select uid,patientuid_37756 as patientuid,number_50713 as c_i ;
  from t_book_65067 where version_end=9223372036854775807 order by patientuid"

 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 't_book_65067')
 ELSE 
 * = MESSAGEBOX('Records selected!', 48, 't_book_65067')
  COPY TO &tc_comdir\t_book_65067
  USE 
 ENDIF 
 WAIT CLEAR 

* && Street
* SelString = "select * from t_street_45317 where is_top=1"

* lrslt = SQLEXEC(lnHndl, selstring)
* IF lrslt<=0
*  = MESSAGEBOX('Cannot process select!', 16, 't_street_45317')
* ELSE 
*  = MESSAGEBOX('Records selected!', 48, 't_street_45317')
*  COPY TO &tc_comdir\gstreet
*  USE 
* ENDIF 

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
 
 WAIT "СБОРКА РЕГИСТРА..." WINDOW NOWAIT 

 SET NULL OFF 

 IF fso.FileExists(tc_comdir+'\people.dbf')
  fso.DeleteFile(tc_comdir+'\people.dbf')
 ENDIF 

 CREATE TABLE &tc_comdir\people ;
  (ok l, used l, id i, c_t n(2), bookid i, c_i c(25), policyuid i, s_pol c(25), n_pol c(25),;
   sn_pol c(25), sn_polerz c(25), ;
   qerz c(2), v_erz2 c(3), d_erz2 d, fam c(25), im c(20), ot c(20), w n(1), dr d, ;
    adrid i, ul n(5), kladr c(4), dom c(7), kor c(5), "str" c(5), kv c(5), regdate t, upddate t) 
 USE 
 SET NULL ON 

 =OpenFile(tc_comdir+'\people', 'people', 'excl')
 SELECT people
 INDEX on id TAG id 
 SET ORDER TO id

 =OpenFile(tc_comdir+'\street', 'street', 'shared', 'ul')
 =OpenFile(tc_comdir+'\territxx', 'territ', 'shared', 'c_t')
 
 =OpenFile(tc_comdir+'\t_patient_10905', 'patient', 'shar')
 =OpenFile(tc_comdir+'\t_policy_43176', 'policy', 'excl')
 =OpenFile(tc_comdir+'\t_address_47256', 'address', 'excl')
 =OpenFile(tc_comdir+'\t_book_65067', 'book', 'excl')
 SELECT policy
 INDEX for state='1' on patientuid TAG uid
 SET ORDER TO uid
 SELECT address
 INDEX on patientuid TAG uid
 SET ORDER TO uid
 SELECT book
 INDEX on patientuid TAG uid
 SET ORDER TO uid

 SELECT patient
 SET RELATION TO uid INTO book
 SET RELATION TO uid INTO policy ADDITIVE 
 SET RELATION TO uid INTO address ADDITIVE 
 SCAN
  m.used = .f. 
  m.sn_polerz = ""
  m.qerz = ""
  m.v_erz2 = ""
  m.d_erz2 = {}
  m.id      = uid
  m.w       = INT(VAL(ALLTRIM(w)))
  m.dr      = IIF(BETWEEN(dr,{01.01.1900 00:00:00},DATETIME()), TTOD(dr), {})
  m.fam     = PROPER(ALLTRIM(fam))
  m.im      = PROPER(ALLTRIM(im))
  m.ot      = PROPER(ALLTRIM(ot))
  m.regdate = IIF(BETWEEN(regdate,{01.01.1994 00:00:00},DATETIME()),regdate, {})
  m.upddate = IIF(BETWEEN(upddate,{01.01.1994 00:00:00},DATETIME()),upddate, {})
  
  m.c_i = ALLTRIM(book.c_i)
  m.bookid = book.uid
*  m.c_i = ALLTRIM(c_i)
*  m.bookid = bookid
  
  m.policyuid = policy.uid
  m.s_pol = ALLTRIM(policy.s_pol)
  m.n_pol = ALLTRIM(policy.n_pol)
  m.sn_pol = IIF(!EMPTY(m.s_pol), m.s_pol + ' ' + m.n_pol, m.n_pol)
  
  m.adrid = address.uid
  m.c_t   = INT(VAL(address.c_t))
  m.kladr = address.ul
  m.ul    = IIF(!EMPTY(address.ul), IIF(SEEK(address.ul,'street', 'kladr'),street.ul,0),0)
  m.dom   = ALLTRIM(address.dom)
  m.kor   = ALLTRIM(address.kor)
  m.str   = ALLTRIM(address.str)
  m.kv    = ALLTRIM(address.kv)
  
  m.str = IIF(ISNULL(m.str), "", m.str)
  
  m.ok = .t.
  m.ok = IIF(!SEEK(m.c_t, 'territ'), .f., ok)
  m.ok = IIF(EMPTY(policyuid), .f., ok)
  m.ok = IIF(!INLIST(m.w,1,2), .f., ok)
  m.ok = IIF(!SEEK(m.ul, 'street'), .f., ok)
  m.ok = IIF(c_t=77 and EMPTY(m.kv), .f., ok)
  m.ok = IIF(MONTH(m.dr)=1 and DAY(m.dr)=1, .f., ok)
  m.ok = IIF(m.c_t=77 and (!IsKms(m.sn_pol) and !IsVs(m.sn_pol) and !IsEnp(m.sn_pol)), .f., ok)

*  IF m.c_t>0
   INSERT INTO people FROM MEMVAR 
*  ENDIF 

 ENDSCAN 
 SET RELATION OFF INTO book
 SET RELATION OFF INTO policy
 SET RELATION OFF INTO address
 USE 
 
 SELECT policy
 SET ORDER TO
 DELETE TAG ALL 
 USE 
 SELECT address
 SET ORDER TO
 DELETE TAG ALL 
 USE 
 SELECT book
 SET ORDER TO
 DELETE TAG ALL 
 USE 

 USE IN people
 USE IN street
 
 WAIT CLEAR 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'PEOPLE СФОРМИРОВАН!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 