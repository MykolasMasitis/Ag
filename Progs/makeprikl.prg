PROCEDURE MakePrikl

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
 
 WAIT "T_ATTACHEDPATS_27258..." WINDOW NOWAIT 
 SelString01 = "select patientuid_42554 as patuid, policyuid_49710 as policyuid, uid, ;
  CAST(s_pol_64779 as char(6)) as s_pol, CAST(n_pol_48945 as char(16)) as n_pol, ;
  CAST(tip_d_10874 as char(1)) as tip_d, "
 SelString02 = "CAST(spos_20148 as char(2)) as spos, CAST(q_27385 as char(2)) as q, ;
  CAST(fam_02179 as char(25)) as fam, CAST(im_29795 as char(20))as im, CAST(ot_15411 as char(20)) as ot, ;
  dr_25711 as dr, w_65202 as w, "
 SelString03 = "date_in_17018 as date_in, date_out_23025 as date_out from t_attachedpats_27258 ;
  where version_end=9223372036854775807 order by patuid"
 SelString = SelString01 + SelString02 + SelString03

 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 't_attachedpats_27258')
 ELSE 
  COPY TO &tc_comdir\attppl
  USE 
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