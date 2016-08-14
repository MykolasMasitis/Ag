PROCEDURE MakeExpFile
 IF MESSAGEBOX(CHR(13)+CHR(10)+'¬€ ’Œ“»“≈ œŒ—“–Œ»“‹ ‘¿…À › —œŒ–“¿?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(tc_expimp)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ »ÃœŒ–“¿!'+CHR(13)+CHR(10),0+16,tc_expimp)
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(tc_outdir)
  fso.CreateFolder(tc_outdir)
 ENDIF 
 
 fso.DeleteFile(tc_outdir+'\*.*')
 
 olddir = SYS(5)+SYS(2003)
 SET DEFAULT TO (tc_expimp)
 IF !fso.FileExists(tc_expimp+'\expimp.dat')
  ImpFile=GETFILE('dat')
 ELSE 
  ImpFile = tc_expimp+'\expimp.dat'
 ENDIF 
 SET DEFAULT TO (olddir)
 
 IF EMPTY(ImpFile)
  olddir = SYS(5)+SYS(2003)
  MESSAGEBOX(CHR(13)+CHR(10)+'¬€ Õ»◊≈√Œ Õ≈ ¬€¡–¿À»!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
  
 ffile = fso.GetFile(ImpFile)
 IF ffile.size >= 2
  fhandl = ffile.OpenAsTextStream
  lcHead = fhandl.Read(2)
  fhandl.Close
 ELSE 
  lcHead = ''
 ENDIF 

 llIsOneZip = .f.
 IF lcHead == 'PK' && ›ÚÓ zip-Ù‡ÈÎ!
  UnzipOpen(ImpFile)
  DataPack ='datapack.xml'
  IF UnzipGotoFileByName(DataPack)
   llIsOneZip = .t.
   UnzipClose()
  ENDIF 
  UnzipClose()
 ELSE 
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈ “Œ ¬€¡–¿À»!'+CHR(13)+CHR(10),0+16,ImpFile)
  RETURN 
 ENDIF 

 IF llIsOneZip = .F.
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈ “Œ ¬€¡–¿À»!'+CHR(13)+CHR(10),0+16,ImpFile)
  RETURN 
 ENDIF 

 fso.CopyFile(tc_expimp+'\expimp.dat', tc_outdir+'\expimp.dat')
 IF !UnZipQuick(tc_outdir+'\expimp.dat', tc_outdir+'\')
  MESSAGEBOX('Õ≈ ”ƒ¿ÀŒ—‹ –¿—œ¿ Œ¬¿“‹ ¿–’»¬!',0+16,'expimp.dat')
  RETURN 
 ENDIF 
 
 WAIT "œŒƒ√Œ“Œ¬ ¿..." WINDOW NOWAIT 

 fso.DeleteFile(tc_outdir+'\expimp.dat')

 fso.DeleteFile(tc_outdir+'\'+'EHR_Book.csv')
 fso.DeleteFile(tc_outdir+'\'+'MskOMS_Patient.csv')
 fso.DeleteFile(tc_outdir+'\'+'MskOMS_PatientAddress.csv')
 fso.DeleteFile(tc_outdir+'\'+'MskOMS_Policy.csv')
 fso.DeleteFile(tc_outdir+'\'+'MskOMS_TAP.csv')
 fso.DeleteFile(tc_outdir+'\'+'MskOMS_TAPService.csv')
 fso.DeleteFile(tc_outdir+'\'+'Patient_Address.csv')
 fso.DeleteFile(tc_outdir+'\'+'Patient_Patient.csv')
 fso.DeleteFile(tc_outdir+'\'+'Patient_Policy.csv')
 fso.DeleteFile(tc_outdir+'\'+'Stat_TAP.csv')
 fso.DeleteFile(tc_outdir+'\'+'Stat_TAPDiagnosis.csv')
 fso.DeleteFile(tc_outdir+'\'+'Stat_TAPService.csv')

 && ›ÚÓ ÚÂ, ÍÓÚÓ˚Â Ì‡‰Ó Ó·ÌÛÎËÚ¸!
 fso.DeleteFile(tc_outdir+'\'+'LPU_HealthCareAreaAssign.csv')
 fso.DeleteFile(tc_outdir+'\'+'Patient_HealthCareArea.csv')
 fso.DeleteFile(tc_outdir+'\'+'Patient_DUL.csv')
 fso.DeleteFile(tc_outdir+'\'+'Patient_Occupation.csv')
 fso.DeleteFile(tc_outdir+'\'+'Stat_DSAudit.csv')
 fso.DeleteFile(tc_outdir+'\'+'Stat_TapAudit.csv')
 fso.DeleteFile(tc_outdir+'\'+'Stat_DetachedService.csv')
 fso.DeleteFile(tc_outdir+'\'+'MskOMS_DetachedService.csv')
 && ›ÚÓ ÚÂ, ÍÓÚÓ˚Â Ì‡‰Ó Ó·ÌÛÎËÚ¸!

 IF !fso.FileExists(tc_basedir+'\talon.dbf')
  RETURN 
 ENDIF 
 IF !fso.FileExists(tc_comdir+'\people.dbf')
  RETURN 
 ENDIF 
 IF !fso.FileExists(tc_comdir+'\street.dbf')
  RETURN 
 ENDIF 
 IF !fso.FileExists(tc_comdir+'\smo.dbf')
  RETURN 
 ENDIF 
 
 IF OpenFile(tc_basedir+'\talon', 'talon', 'excl')>0
  RETURN 
 ENDIF 
 IF OpenFile(tc_comdir+'\people', 'people', 'shar', 'id')>0
  USE IN talon
  RETURN 
 ENDIF 
 IF OpenFile(tc_comdir+'\street', 'street', 'shar', 'ul')>0
  USE IN talon
  USE IN people 
  RETURN 
 ENDIF 
 IF OpenFile(tc_comdir+'\smo', 'smo', 'shar', 'q')>0
  USE IN street
  USE IN talon
  USE IN people 
  RETURN 
 ENDIF 
 
 WAIT CLEAR 

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ LPU_HealthCareAreaAssign.csv..." WINDOW NOWAIT 
 fhandl=fso.OpenTextFile(tc_outdir+'\LPU_HealthCareAreaAssign.csv',2,.t.)
 fhandl.WriteLine('ProfessionalUID,HealthCareAreaUID')
 fhandl.Close

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Patient_HealthCareArea.csv..." WINDOW NOWAIT 
 fhandl=fso.OpenTextFile(tc_outdir+'\Patient_HealthCareArea.csv',2,.t.)
 fhandl.WriteLine('PatientUID,HealthCareAreaUID')
 fhandl.Close

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Patient_DUL.csv..." WINDOW NOWAIT 
 fhandl=fso.OpenTextFile(tc_outdir+'\Patient_DUL.csv',2,.t.)
 fhandl.WriteLine('DULUID,PatientUID,DULType,DULSeries,DULNumber,DULOwnerEx,DULOwner,DULOwner_Fam,DULOwner_Im,DULOwner_Ot,DULOwner_Sex,DULOwner_Birthday,Disabled,IssueDate,VoidDate,Issuer,UpdateDate,IssueState')
 fhandl.Close

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Patient_Occupation.csv..." WINDOW NOWAIT 
 fhandl=fso.OpenTextFile(tc_outdir+'\Patient_Occupation.csv',2,.t.)
 fhandl.WriteLine('PatientUID,OrganisationUID,OccupationMode,Profession,Position,Phone,OrganisationBranchUID')
 fhandl.Close

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Stat_DSAudit.csv..." WINDOW NOWAIT 
 fhandl=fso.OpenTextFile(tc_outdir+'\Stat_DSAudit.csv',2,.t.)
 fhandl.WriteLine('ServiceUID,Defect,Comments')
 fhandl.Close

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Stat_TAPAudit.csv..." WINDOW NOWAIT 
 fhandl=fso.OpenTextFile(tc_outdir+'\Stat_TapAudit.csv',2,.t.)
 fhandl.WriteLine('TAPUID,Defect,Comments')
 fhandl.Close

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Stat_DetachedService.csv..." WINDOW NOWAIT 
 fhandl=fso.OpenTextFile(tc_outdir+'\Stat_DetachedService.csv',2,.t.)
 fhandl.WriteLine('ServiceUID,LPUUID,PatientUID,DULUID,AddressUID,BookUID,PolicyUID,ProfessionalUID,RenderingDate,Diagnosis,MedicalProgramm,ServiceCode,Count,Trauma,NurseUID,AppointmentUID')
 fhandl.Close

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ MskOMS_DetachedService.csv..." WINDOW NOWAIT 
 fhandl=fso.OpenTextFile(tc_outdir+'\MskOMS_DetachedService.csv',2,.t.)
 fhandl.WriteLine('ServiceUID,SpecCase,LPUFor')
 fhandl.Close
 
 WAIT "Œ¡–¿¡Œ“ ¿..." WINDOW NOWAIT 

 m.CreatedSum = 0
 
 CREATE CURSOR ehr_book (bookuid i, policyuid i, patientuid i, "number" c(25), v_ans c(3), d_ans d, adruid i, ul n(5))
 INDEX ON bookuid TAG bookuid
* INDEX ON adruid TAG adruid UNIQUE 
 INDEX ON policyuid TAG policyuid UNIQUE 
 SET ORDER TO bookuid
 
 CREATE CURSOR tap (tapuid i)
 INDEX on tapuid TAG tapuid 
 SET ORDER TO tapuid
 
 SELECT talon 
 INDEX FOR ok ON npatuid TAG npatuid
 SET ORDER TO npatuid
 
 SCAN 
  m.npatuid = npatuid
  m.tapuid = ntapuid
  IF !SEEK(m.tapuid, 'tap')
   INSERT INTO tap (tapuid) VALUES (m.tapuid)
  ELSE 
  ENDIF 
  IF SEEK(m.npatuid, 'people')
   m.bookuid = people.bookid
   m.policyuid = people.policyuid
   m.c_i     = people.c_i
   m.v_ans   = people.v_erz2
   m.d_ans   = people.d_erz2
   m.adruid  = people.adrid
   m.ul      = people.ul
  ELSE 
   m.bookuid = 0
   m.policyuid = 0
   m.c_i     = ''
   m.v_ans   = ''
   m.d_ans   = {}
   m.adruid  = 0
   m.ul      = 0
  ENDIF 

  IF !SEEK(m.bookuid, 'ehr_book') AND m.bookuid != 0
   INSERT INTO ehr_book (bookuid, policyuid, patientuid, "number", v_ans, d_ans, adruid, ul) VALUES ;
   (m.bookuid, m.policyuid, m.npatuid, m.c_i, m.v_ans, m.d_ans, m.adruid, m.ul)
  ENDIF 

  m.CreatedSum = m.CreatedSum + s_all 
 ENDSCAN 

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ EHR_Book.csv, MskOMS_Patient, MskOMS_PatientAddress......" WINDOW NOWAIT 
 SELECT ehr_book
 fhandl1=fso.OpenTextFile(tc_outdir+'\EHR_Book.csv',2,.t.)
 fhandl1.WriteLine('BookUID,PatientUID,BookType,Number,Disabled')

 fhandl2=fso.OpenTextFile(tc_outdir+'\MskOMS_Patient.csv',2,.t.)
 fhandl2.WriteLine('PatientUID,CheckVector,CheckDate,SpecCase')

 fhandl3=fso.OpenTextFile(tc_outdir+'\MskOMS_PatientAddress.csv',2,.t.)
 fhandl3.WriteLine('AddressUID,CodeMskStreet,SettlementType')

 fhandl4=fso.OpenTextFile(tc_outdir+'\MskOMS_Policy.csv',2,.t.)
 fhandl4.WriteLine('PolicyUID,CheckVector,CheckDate,SpecCase,PolSeries,PolNumber,SMOUID')

 SCAN 
  m.StringToWrite1 = ALLTRIM(STR(bookuid))+','+ALLTRIM(STR(patientuid))+',01,'+ALLTRIM(number)+',0'
  fhandl1.WriteLine(m.StringToWrite1)

  m.checkdate = STR(YEAR(d_ans),4)+'-'+PADL(MONTH(d_ans),2,'0')+'-'+PADL(DAY(d_ans),2,'0')+ ' 14:32:00'
  m.StringToWrite2 = ALLTRIM(STR(patientuid))+','+ALLTRIM(v_ans)+','+m.checkdate+',0'
  fhandl2.WriteLine(m.StringToWrite2)

  m.StringToWrite3 = ALLTRIM(STR(adruid))+','+PADL(ul,5,'0')+',1'
  fhandl3.WriteLine(m.StringToWrite3)

  m.checkdate = IIF(BETWEEN(d_ans,{01.01.2000},DATE()), STR(YEAR(d_ans),4)+'-'+PADL(MONTH(d_ans),2,'0')+'-'+PADL(DAY(d_ans),2,'0')+ ' 14:32:00', '')
  m.StringToWrite4 = ALLTRIM(STR(policyuid))+','+ALLTRIM(v_ans)+','+m.checkdate+',,,,'
  fhandl4.WriteLine(m.StringToWrite4)
 ENDSCAN 
 fhandl1.Close
 fhandl2.Close
 fhandl3.Close
 fhandl4.Close

* WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ MskOMS_Patient..." WINDOW NOWAIT 
* fhandl=fso.OpenTextFile(tc_outdir+'\MskOMS_Patient.csv',2,.t.)
* fhandl.WriteLine('PatientUID,CheckVector,CheckDate,SpecCase')

* SCAN 
*  m.checkdate = STR(YEAR(d_ans),4)+'-'+PADL(MONTH(d_ans),2,'0')+'-'+PADL(DAY(d_ans),2,'0')+ ' 14:32:00'
*  m.StringToWrite = ALLTRIM(STR(patientuid))+','+ALLTRIM(v_ans)+','+m.checkdate+',0'
*  fhandl.WriteLine(m.StringToWrite)
* ENDSCAN 
* fhandl.Close

* WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ MskOMS_PatientAddress..." WINDOW NOWAIT 
* fhandl=fso.OpenTextFile(tc_outdir+'\MskOMS_PatientAddress.csv',2,.t.)
* fhandl.WriteLine('AddressUID,CodeMskStreet,SettlementType')
* SET ORDER TO adruid 
* SCAN 
*  m.StringToWrite = ALLTRIM(STR(adruid))+','+PADL(ul,5,'0')+',1'
*  fhandl.WriteLine(m.StringToWrite)
* ENDSCAN 
* fhandl.Close
 
* WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ MskOMS_Policy..." WINDOW NOWAIT 
* fhandl=fso.OpenTextFile(tc_outdir+'\MskOMS_Policy.csv',2,.t.)
* fhandl.WriteLine('PolicyUID,CheckVector,CheckDate,SpecCase,PolSeries,PolNumber,SMOUID')
* SET ORDER TO policyuid
* SCAN 
*  m.checkdate = IIF(BETWEEN(d_ans,{01.01.2000},DATE()), STR(YEAR(d_ans),4)+'-'+PADL(MONTH(d_ans),2,'0')+'-'+PADL(DAY(d_ans),2,'0')+ ' 14:32:00', '')
*  m.StringToWrite = ALLTRIM(STR(policyuid))+','+ALLTRIM(v_ans)+','+m.checkdate+',,,,'
*  fhandl.WriteLine(m.StringToWrite)
* ENDSCAN 
* fhandl.Close
 
 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ MskOMS_TAP..." WINDOW NOWAIT 
 SELECT tap
 fhandl=fso.OpenTextFile(tc_outdir+'\MskOMS_TAP.csv',2,.t.)
 fhandl.WriteLine('TAPUID,SpecCase')
 SCAN 
  m.StringToWrite = ALLTRIM(STR(tapuid))+','
  fhandl.WriteLine(m.StringToWrite)
 ENDSCAN 
 fhandl.Close
 
 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ MskOMS_TAPService..." WINDOW NOWAIT 
 SELECT talon 
 INDEX FOR ok on diaguid TAG diaguid
 SET ORDER TO diaguid
 fhandl=fso.OpenTextFile(tc_outdir+'\MskOMS_TAPService.csv',2,.t.)
 fhandl.WriteLine('DiagnosisUID,ServiceCode,SpecCase')
 SCAN 
  m.StringToWrite = ALLTRIM(STR(diaguid))+','+ALLTRIM(coduid)+',0'
  fhandl.WriteLine(m.StringToWrite)
 ENDSCAN 
 fhandl.Close

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Patient_Address..." WINDOW NOWAIT 
* INDEX on npatuid TAG npatuid
 SET ORDER TO npatuid
 fhandl=fso.OpenTextFile(tc_outdir+'\Patient_Address.csv',2,.t.)
 fhandl.WriteLine('PatientUID,AddressType,Disabled,PostIndex,CodeState,CodeRegion,CodeDistrict,CodeSettlement,CodeStreet,RegionName,DistrictName,SettlementName,StreetName,HouseNumber,BuildingNumber,UnitNumber,FlatNumber,AddressUID,DoorCode,NonFormalizedAddress,UpdateDate')
 SELECT people
 SET ORDER TO id
 SET RELATION TO id INTO talon 
 SET RELATION TO ul INTO street ADDITIVE 
 
 SCAN FOR !EMPTY(talon.patientuid)
  m.PatientUID = ALLTRIM(STR(id))
  m.AddressType = '1'
  m.Disabled = '0'
  m.PostIndex = '""'
  m.CodeState = '643'
  m.CodeRegion = '77'
  m.CodeDistrict = '000'
  m.CodeSettlement = '000000'
  m.CodeStreet = kladr
  m.RegionName = 'ÃÓÒÍ‚‡ „'
  m.DistrictName = '""'
  m.SettlementName = '""'
  m.StreetName = ALLTRIM(street.street)
  m.HouseNumber = ALLTRIM(dom)
  m.BuildingNumber = ALLTRIM(kor)
  m.UnitNumber = ALLTRIM(str)
  m.FlatNumber = ALLTRIM(kv)
  m.AddressUID = ALLTRIM(STR(adrid))
  m.DoorCode = '""'
  m.NonFormalizedAddress = '""'
  m.UpdateDate = ALLTRIM(upddate)

  m.StringToWrite = m.PatientUID+','+m.AddressType+','+m.Disabled+','+m.PostIndex+','+m.CodeState+','+;
  m.CodeRegion+','+m.CodeDistrict+','+m.CodeSettlement+','+m.CodeStreet+','+m.RegionName+','+m.DistrictName+','+;
  m.SettlementName+','+m.StreetName+','+ m.HouseNumber+','+m.BuildingNumber+','+m.UnitNumber+','+m.FlatNumber+','+;
  m.AddressUID+','+m.DoorCode+','+m.NonFormalizedAddress+','+m.UpdateDate
  fhandl.WriteLine(m.StringToWrite)
 ENDSCAN 
 fhandl.Close
 
 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Patient_Patient..." WINDOW NOWAIT 
 SELECT people 
 SET RELATION OFF INTO street
 fhandl=fso.OpenTextFile(tc_outdir+'\Patient_Patient.csv',2,.t.)
 fhandl.Write('PatientUID,Code,Fam,Im,Ot,Birthday,BirthPlace,Sex,SocialStatus,InvGroup,InvFromChildhood,InvFirstTimeFix,InvReset,SNILS,LPUUID,HomePhone,WorkPhone')
 fhandl.WriteLine(',EmploymentPlace,Profession,Position,Dependant,BloodGroup,Rh,IsVillager,RegistrationDate,UpdateDate,CloseRegistrationCause,CloseRegistrationDate,Citizenship,Comment')
 SCAN FOR !EMPTY(talon.patientuid)
  m.PatientUID=ALLTRIM(STR(id))
  m.Code = '""'
  m.Fam = ALLTRIM(fam)
  m.Im = ALLTRIM(im)
  m.Ot = ALLTRIM(ot)
  m.Birthday = STR(YEAR(dr),4)+'-'+PADL(MONTH(dr),2,'0')+'-'+PADL(DAY(dr),2,'0')
  m.BirthPlace = '""'
  m.Sex = STR(w,1)
  m.SocialStatus = ''
  m.InvGroup = '0'
  m.InvFromChildhood = '0'
  m.InvFirstTimeFix = '0'
  m.InvReset = '0'
  m.SNILS = '""'
  m.LPUUID = ''
  m.HomePhone = '""'
  m.WorkPhone = '""'
  m.EmploymentPlace = '""'
  m.Profession = '""'
  m.Position = '""'
  m.Dependant = '0'
  m.BloodGroup = ''
  m.Rh = ''
  m.IsVillager = '0'
  m.RegistrationDate = ALLTRIM(regdate)
  m.UpdateDate = ALLTRIM(upddate)
  m.CloseRegistrationCause = ''
  m.CloseRegistrationDate = ''
  m.Citizenship = '643'
  m.Comment = '""'
  m.StringToWrite = m.PatientUID+','+m.Code+','+m.Fam+','+m.Im+','+m.Ot+','+m.Birthday+','+m.BirthPlace+','+m.Sex+','+;
  m.SocialStatus+','+m.InvGroup+','+m.InvFromChildhood+','+m.InvFirstTimeFix+','+m.InvReset+','+m.SNILS+','+m.LPUUID+','+;
  m.HomePhone+','+m.WorkPhone+','+m.EmploymentPlace+','+m.Profession+','+m.Position+','+m.Dependant+','+m.BloodGroup+','+m.Rh+','+;
  m.IsVillager+','+m.RegistrationDate+','+m.UpdateDate+','+m.CloseRegistrationCause+','+m.CloseRegistrationDate+','+;
  m.Citizenship+','+m.Comment
  fhandl.WriteLine(m.StringToWrite)
 ENDSCAN 
 fhandl.Close

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Patient_Policy..." WINDOW NOWAIT 
 fhandl=fso.OpenTextFile(tc_outdir+'\Patient_Policy.csv',2,.t.)
 fhandl.WriteLine('PolicyUID,PatientUID,SMOUID,Series,Number,SMOName,PolicyState,Type,IssueDate,VoidDate,UpdateDate')
 
 SCAN FOR !EMPTY(talon.npatuid)
  m.PolicyUID = ALLTRIM(STR(PolicyUID))
  m.PatientUID = ALLTRIM(STR(id))
*  m.SMOUID = ALLTRIM(STR(smoid))
  m.SMOUID = IIF(SEEK(qerz, 'smo'), ALLTRIM(smo.code), '')
  m.Series = ALLTRIM(s_pol)
  m.Number = ALLTRIM(n_pol)
  m.SMOName = '""'
  m.PolicyState = '1'
  m.Type = '1'
  m.IssueDate = ''
  m.VoidDate = ''
  m.UpdateDate = ALLTRIM(upddate)
  
  m.StringToWrite = m.PolicyUID+','+m.PatientUID+','+m.SMOUID+','+m.Series+','+m.Number+','+m.SMOName+','+m.PolicyState+','+;
  m.Type+','+m.IssueDate+','+m.VoidDate+','+m.UpdateDate
  fhandl.WriteLine(m.StringToWrite)
 ENDSCAN 
 fhandl.Close
 SET RELATION OFF INTO talon 
 
 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Stat_TAP..." WINDOW NOWAIT 
 SELECT talon 
* INDEX on ntapuid TAG ntapuid
* SET ORDER TO ntapuid
 SET RELATION TO npatuid INTO people
 fhandl=fso.OpenTextFile(tc_outdir+'\Stat_TAP.csv',2,.t.)
 fhandl.Write('TAPUID,LPUUID,PatientUID,DULUID,AddressUID,PolicyUID,BookUID,FillingDate,DoctorUID,NurseUID,MedicalProgramm,')
 fhandl.WriteLine('ServicePlace,VisitPurpose,VisitResult,ChangedDiagnosis,ChangeDiagnosisDate,DVNAction,DVNCause,DVNOpenDate,DVNCloseDate,DVNSex,DVNAge,DVNSeries,DVNNumber,AppointmentUID') 
 SCAN 

  m.TAPUID = ALLTRIM(STR(ntapuid))
  m.LPUUID = '55'
  m.PatientUID = ALLTRIM(STR(npatuid))
  m.DULUID = ''
  m.AddressUID = ALLTRIM(STR(people.adrid))
  m.PolicyUID = ALLTRIM(STR(people.policyuid))
  m.BookUID = ALLTRIM(STR(people.bookid))
  m.FillingDate = STR(YEAR(d_u),4)+'-'+PADL(MONTH(d_u),2,'0')+'-'+PADL(DAY(d_u),2,'0')
  m.DoctorUID = IIF(doctoruid>0, ALLTRIM(STR(doctoruid)), '')
  m.NurseUID = IIF(nurseuid>0, ALLTRIM(STR(nurseuid)), '')
  m.MedicalProgramm = 'ŒÃ—Ã '
  m.ServicePlace = '1'
  m.VisitPurpose = ''
  m.VisitResult = ''
  m.ChangedDiagnosis = ''
  m.ChangeDiagnosisDate = ''
  m.DVNAction = '""'
  m.DVNCause = ''
  m.DVNOpenDate = ''
  m.DVNCloseDate = ''
  m.DVNSex = ''
  m.DVNAge = '0'
  m.DVNSeries = '""'	
  m.DVNNumber = '""'
  m.AppointmentUID = ''

  m.StringToWrite = m.TAPUID+','+m.LPUUID+','+m.PatientUID+','+m.DULUID+','+m.AddressUID+','+m.PolicyUID+','+;
   m.BookUID+','+m.FillingDate+','+m.DoctorUID+','+m.NurseUID+','+m.MedicalProgramm+','+m.ServicePlace+','+;
   m.VisitPurpose+','+m.VisitResult+','+m.ChangedDiagnosis+','+m.ChangeDiagnosisDate+','+m.DVNAction+','+;
   m.DVNCause+','+m.DVNOpenDate+','+m.DVNCloseDate+','+m.DVNSex+','+m.DVNAge+','+m.DVNSeries+','+;
   m.DVNNumber+','+m.AppointmentUID

  fhandl.WriteLine(m.StringToWrite)
 ENDSCAN 
 SET RELATION OFF INTO people
 fhandl.Close
 
 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Stat_TAPDiagnosis..." WINDOW NOWAIT 
 
 USE IN tap

 CREATE CURSOR tap (tapuid i)
 INDEX ON tapuid TAG tapuid 
 SET ORDER TO tapuid
 
 SELECT talon 
 SET ORDER TO diaguid
 
 fhandl=fso.OpenTextFile(tc_outdir+'\Stat_TAPDiagnosis.csv',2,.t.)
 fhandl.WriteLine('DiagnosisUID,TAPUID,ICDCode,IsMain,DeseaseNature,MonitoringState,Trauma')
 SCAN 
  m.DiagnosisUID = ALLTRIM(STR(diaguid))
  m.TAPUID = ALLTRIM(STR(ntapuid))
  m.ICDCode = ALLTRIM(ds)
  m.IsMain = '1'
  m.DeseaseNature = ''
  m.MonitoringState = '" "'
  m.Trauma = '""'
  
  m.StringToWrite = m.DiagnosisUID+','+m.TAPUID+','+m.ICDCode+','+m.IsMain+','+m.DeseaseNature+','+m.MonitoringState+','+m.Trauma

  IF !SEEK(INT(VAL(m.tapuid)), 'tap')
   INSERT INTO tap (tapuid) VALUES (INT(VAL(m.tapuid)))
   fhandl.WriteLine(m.StringToWrite)
  ENDIF 
 ENDSCAN 
 fhandl.Close
 
 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Stat_TAPService..." WINDOW NOWAIT 
 fhandl=fso.OpenTextFile(tc_outdir+'\Stat_TAPService.csv',2,.t.)
 fhandl.WriteLine('DiagnosisUID,ServiceCode,Count')
 SCAN 
  m.DiagnosisUID = ALLTRIM(STR(diaguid))
  m.ServiceCode =  ALLTRIM(coduid)
  m.Count = ALLTRIM(TRANSFORM(k_u, '99.99'))

  m.StringToWrite = m.DiagnosisUID+','+m.ServiceCode+','+m.Count
  fhandl.WriteLine(m.StringToWrite)
 ENDSCAN 
 fhandl.Close

 USE IN ehr_book
 USE IN tap 
 USE IN talon
 USE IN people
 USE IN street
 USE IN smo
 
 WAIT CLEAR 

 IF ZipOpen('expimp.dat',tc_outdir+'\',.t.)
 ELSE 
  MESSAGEBOX('Õ≈¬Œ«ÃŒ∆ÕŒ Œ“ –€“‹ ¿–’»¬!',0+16,'expimp.dat')
  RETURN 
 ENDIF 

 IF !ZipFolder(tc_outdir+'\',.f.)
  MESSAGEBOX('Õ≈ ”ƒ¿ÀŒ—‹ ƒŒ¡¿¬»“‹ ‘¿…À ¬ ¿–’»¬!',0+16,'*.csv')
 ENDIF 
 
 ZipClose()
 
 fso.DeleteFile(tc_outdir+'\*.csv')
 fso.DeleteFile(tc_outdir+'\datapack.xml')
 
 MESSAGEBOX('—‘Œ–Ã»–Œ¬¿Õ —◊≈“ Õ¿ —”ÃÃ” '+TRANSFORM(m.CreatedSum,'9 999 999.99'),0+64,'')
 
RETURN 