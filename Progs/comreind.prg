FUNCTION comreind
 IF !fso.FileExists(tc_comdir+'\logfile.dbf')
  CREATE TABLE tc_comdir+'\logfile' ;
   (id c(8), period c(6), sent t, recieved t, tip c(12), totrecs n(6))
  INDEX on id TAG id 
  USE 
 ENDIF 

 pnResult = 0
 pnResult = OpenFile(tc_comdir+'\logfile', 'logfile', 'Excl')
 IF pnResult==0
  SELECT logfile
  INDEX on id TAG id 
  USE 
 ENDIF 

 pnResult = 0
 pnResult = OpenFile('&tc_comdir\tarifn', 'tarif', 'Excl')
 IF pnResult==0
  SELECT tarif
  INDEX ON cod TAG cod
  USE 
 ENDIF 

 pnResult = 0
 pnResult = OpenFile('&tc_comdir\street', 'street', 'Excl')
 IF pnResult==0
  SELECT street
  INDEX ON ul TAG ul
  INDEX on kladr TAG kladr
  USE 
 ENDIF 

 pnResult = 0
 pnResult = OpenFile('&tc_comdir\osoerzxx', 'osoerz', 'Excl')
 IF pnResult==0
  SELECT osoerz
  INDEX ON ans_r TAG ans_r
  USE 
 ENDIF 

 pnResult = 0
 pnResult = OpenFile('&tc_comdir\smo', 'smo', 'Excl')
 IF pnResult==0
  SELECT smo
  INDEX ON q TAG q
  INDEX ON q_ogrn TAG q_ogrn
  INDEX ON code TAG code
  USE 
 ENDIF 

 pnResult = 0
 pnResult = OpenFile('&tc_comdir\territxx', 'ter', 'Excl')
 IF pnResult==0
  SELECT ter
  INDEX ON c_t TAG c_t
  USE 
 ENDIF 

RETURN 