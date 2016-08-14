PROCEDURE ProcessAM
 IF MESSAGEBOX(CHR(13)+CHR(10)+'бш унрхре ябепхрэ пецхярп я нрбернл?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 AMFile = tc_comdir+'\Am_'+PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)
 olddir = SYS(5)+SYS(2003)
 SET DEFAULT TO (tc_comdir)
 IF !fso.FileExists(AMFile+'.dbf')
  AnsFile=GETFILE('dbf')
 ELSE 
  AnsFile = AMFile
 ENDIF 
 SET DEFAULT TO (olddir)
 
 IF EMPTY(AnsFile)
  olddir = SYS(5)+SYS(2003)
  MESSAGEBOX(CHR(13)+CHR(10)+'бш мхвецн ме бшапюкх!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 ffile = fso.GetFile(AnsFile+'.dbf')
 CreateDate = ffile.DateLastModified
 RELEASE m.ffile 
 IF DATETIME()-CreateDate>12*60*60
  IF MESSAGEBOX(CHR(13)+CHR(10)+; 
   'тюик нрберю '+ALLTRIM(AnsFile)+CHR(13)+CHR(10)+; 
   'ашк янгдюм '+TTOC(CreateDate) +'.'+CHR(13)+CHR(10)+; 
   'яеивюя '+TTOC(DATE())+'.'+CHR(13)+CHR(10)+; 
   CHR(13)+CHR(10)+'напюаюршбюрэ бяе пюбмн?',4+32, 'бмхлюмхе') == 7
   RETURN 
  ENDIF 
 ENDIF 

 oApp.CodePage(AnsFile+'.dbf', 866, .t.)

 IF OpenFile(tc_comdir+'\people', 'people', 'shar')>0
  RETURN 
 ENDIF 
 IF OpenFile(AnsFile, 'answer', 'excl')>0
  USE IN people
  RETURN 
 ENDIF 
 IF OpenFile(tc_comdir+'\osoerzxx', 'osoerz', 'shar', 'ans_r')>0
  USE IN people
  USE IN answer
  RETURN 
 ENDIF 
 
 WAIT "напюанрйю..." WINDOW NOWAIT 
 SELECT answer
 INDEX on recid TAG recid 
 SET ORDER TO recid
 SELECT people
 SET RELATION TO PADL(RECNO(),6,'0') INTO answer
 SET RELATION TO v_erz2 INTO osoerz ADDITIVE 
 
 COUNT FOR !EMPTY(answer.recid) TO m.nRelated
 COUNT FOR c_t=77 TO m.RecsInPeople

 WAIT CLEAR 
 
 IF m.RecsInPeople != m.nRelated
  IF MESSAGEBOX(CHR(13)+CHR(10)+'йнк-бн гюохяеи б пецхярпе ('+STR(m.RecsInPeople,5)+')'+CHR(13)+CHR(10)+;
   'ме янбоюдюер я йнк-бнл гюохяеи б нрбере ('+STR(m.nRelated,5)+')'+CHR(13)+CHR(10)+;
   'опнднкфхрэ?'+CHR(13)+CHR(10), 4+32, '') == 7
   SET RELATION OFF INTO answer
   SET RELATION OFF INTO osoerz
   USE 
   SELECT answer 
   SET ORDER TO 
   DELETE TAG ALL 
   USE 
   RETURN 
  ENDIF 
 ENDIF 
 
 WAIT "напюанрйю..." WINDOW NOWAIT 
 SCAN 
  IF !EMPTY(answer.recid)
   m.qerz = answer.q
   m.v_erz2 = answer.ans_r
   m.sn_polerz = ''
   d_erz2 = {}
   DO CASE 
    CASE answer.tip_d='я'
     m.sn_polerz = answer.s_pol+' '+answer.n_pol
     d_erz2 = DATE()
    CASE answer.tip_d='о'
     m.sn_polerz = answer.n_pol
     d_erz2 = DATE()
    OTHERWISE 
   ENDCASE 
   REPLACE sn_polerz WITH m.sn_polerz, qerz WITH m.qerz, v_erz2 WITH m.v_erz2, d_erz2 WITH m.d_erz2
  ENDIF 
  REPLACE ok WITH IIF(!EMPTY(c_i) AND osoerz.kl=='y' AND sn_pol==sn_polerz, .t., .f.)
 ENDSCAN 

 SET RELATION OFF INTO answer
 SET RELATION OFF INTO osoerz
 SELECT answer 
 SET ORDER TO 
 DELETE TAG ALL 
 USE 
 USE IN people
 
* SELECT people
* SET RELATION TO v_erz2 INTO osoerz
 
* SCAN 
*  REPLACE ok WITH IIF(!EMPTY(c_i) AND osoerz.kl=='y', .t., .f.)
* ENDSCAN 
 
* SET RELATION OFF INTO osoerz
* USE 
 USE IN osoerz

 WAIT CLEAR 

 MESSAGEBOX('напюанрйю гюйнмвемю!', 0+64, '')

RETURN 