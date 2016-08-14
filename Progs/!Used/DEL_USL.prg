Proc del_usl
dele for !found
dele for !replace
dele for pol!=new_pol
dele for year(wasborn)!=year(new_born)

dele for inli(cod_usl,1561,1601)

*dele for inli(cod_usl,1561,1601,16006,16007,1584,8048,6075,3018,22720,22721,;
* 22722,9151,9152,9102,9156,9153,5004)
*dele for inli(cod_usl,2001,2002,22,23,1017,1584,50124,5004,6017,6069,35236,35223,3051)
*dele for inli(cod_usl,6014,6016,6035,6081,13006,2008,3055,3053,35002,35005,35212,35310)
*dele for betw(cod_usl,57001,57017)
*dele for betw(cod_usl,25001,25990)
*dele for betw(cod_usl,26001,26266)
*dele for betw(cod_usl,1211,1227)
*dele for betw(cod_usl,40001,40093)
*dele for betw(cod_usl,43006,43035)
*dele for betw(cod_usl,1261,1267)
*dele for betw(cod_usl,7002,7081)
*dele for betw(cod_usl,9001,9362)

dele for betw(cod_diag,'Z32.0 ','Z39.2 ')
dele for betw(cod_diag,'O20.0 ','O99.8 ')
dele for inli(cod_diag,'Z10.0','Z01.8','Z01.6','Z98.8')

*DELE FOR INLI(COD_usl,1041,1042,1045,1047)
*dele for inli(cod_usl,1301,1302,1303,1305,1307)