Attribute VB_Name = "iFix_DB_Import"

'============== EXCEL SOURCE TABLE INFO ==============
' Table Name: "MasterV3_DB"
' Table Columns: ["iFIX BLOCK TYPE","DATA TYPE","ADDRESS","OLD TAGNAME (IF APPLICABLE)","TAGNAME","OLD DESCRIPTION (IF APPLICABLE)","DESCRIPTION","Sensor Lo Limit (Eng Units)","Sensor HI Limit (Eng Units)","UNITS","On Code","Off Code"]
' Notes:
'   - For any field in a row that contains a literal "-", the field should be treated as blank/empty and should be exported as "" in the final output file.
'   - The "iFIX BLOCK TYPE" field determines which block type format (DA, DR, MDI, AA, AR) the row of data should be exported as in the final output file.

'============== EXCEL DESTINATION INFO ==============
' Sheet Name: "iFix v5.8 Export"

'=============== FILE HEADER FORMAT ===============
' [NodeName : EFSVR,Database : ELMFORK,File Name : ChaseLynk_ExcelCompiled,Date : MM/DD/YYYY,Time :  HH:MM:SS AM/PM]
' #AI4544,AO1,DI29935,AR14326,DR25580,RM4,MDI584,BL1,AA1308,DA2377,DT10,CA13,FN1,TX6,EV69,TM1#

'================== DA FORMAT ==================
' Two header rows of formatting and preamble before the data rows. All fields in data rows must contain quotation marks surrounding the data.
' [BLOCK TYPE,TAG,NEXT BLOCK,DESCRIPTION,I/O DEVICE,H/W OPTIONS,I/O ADDRESS,INITIAL A/M STATUS,INITIAL SCAN,SCAN TIME,
'   INVERT OUTPUT,OPEN TAG,CLOSE TAG,ALARM ENABLE,ALARM AREA(S),ALARM PRIORITY,ALARM TYPE,ENABLE OUTPUT,SECURITY AREA 1,SECURITY AREA 2,
'   SECURITY AREA 3,ALARM AREA 1,ALARM AREA 2,ALARM AREA 3,ALARM AREA 4,ALARM AREA 5,ALARM AREA 6,ALARM AREA 7,ALARM AREA 8,ALARM AREA 9,
'   ALARM AREA 10,ALARM AREA 11,ALARM AREA 12,ALARM AREA 13,ALARM AREA 14,ALARM AREA 15,USER FIELD 1,USER FIELD 2,CONTACT TAG,CONTACT MODE,
'   REMOTE ACK TAG,ALARM SUSPENSION TAG,DELAY TIME,REALARM TIME,ESIG TYPE,ESIG ALLOW CONT USE,ESIG XMPT ALARM ACK,ESIG UNSIGNED WRITES,PDR Update Rate,PDR Access Time,
'   PDR Deadband,PDR Latch,PDR Disable Output,PDR Array Length,Hist Description,Hist Collect,Hist Interval,Hist Offset,Hist Time Res,Hist Compress,
'   Hist Deadband,Hist Comp Type,Hist Comp Time,Continuous Contact Output,Suppress COMM Alarm]
' !A_NAME,A_TAG,A_NEXT,A_DESC,A_IODV,A_IOHT,A_IOAD,A_IAM,A_ISCAN,A_SCANT,
'   A_INV,A_OPENDESC,A_CLOSEDESC,A_IENAB,A_ADI,A_PRI,A_ALMCK,A_EOUT,A_SA1,A_SA2,
'   A_SA3,A_AREA1,A_AREA2,A_AREA3,A_AREA4,A_AREA5,A_AREA6,A_AREA7,A_AREA8,A_AREA9,
'   A_AREA10,A_AREA11,A_AREA12,A_AREA13,A_AREA14,A_AREA15,A_ALMEXT1,A_ALMEXT2,A_CCALM,A_CCMOD,
'   A_AACK,A_ASUSP,A_DELAY,A_REALM,A_ESIGTYPE,A_ESIGCONT,A_ESIGACK,A_ESIGTRAP,A_PDR_UPDATERATE,A_PDR_ACCESSTIME,
'   A_PDR_DEADBAND,A_PDR_LATCHDATA,A_PDR_DISABLEOUT,A_PDR_ARRAYLENGTH,A_HIST_DESC,A_HIST_COLLECT,A_HIST_INTERVAL,A_HIST_OFFSET,A_HIST_TIMERES,A_HIST_COMPRESS,
'   A_HIST_DEADBAND,A_HIST_COMPTYPE,A_HIST_COMPTIME,A_CTK_PERSIST,A_COMM_SURPRESS!
' "DA","<TAGNAME>","","<DESCRIPTION>","MBE","","<ADDRESS>","AUTO","ON","1",
'   "NO","<OFF CODE>","<ON CODE>","ENABLE","NONE","LOW","COS","NO","NONE","NONE",
'   "NONE","ALL","","","","","","","","",
'   "","","","","","","","","","",
'   "","","00:00:00:00","00:00:00:00","NONE","YES","NO","REJECT","1,000","300,000",
'   "0","NO","NO","0","","NO","5,000.0","0.0","Milliseconds","DISABLE",
'   "0.0000","Absolute","0","DISABLE","DISABLE"

'================== DR FORMAT ==================
' Two header rows of formatting and preamble before the data rows. All fields in data rows must contain quotation marks surrounding the data.
' [BLOCK TYPE,TAG,DESCRIPTION,I/O DEVICE,H/W OPTIONS,I/O ADDRESS TYPE,I/O ADDRESS,ENABLE OUTPUT,INVERT OUTPUT,OPEN TAG,
'   CLOSE TAG,EVENT MESSAGES,ALARM AREA(S),SECURITY AREA 1,SECURITY AREA 2,SECURITY AREA 3,ALARM AREA 1,ALARM AREA 2,ALARM AREA 3,ALARM AREA 4,
'   ALARM AREA 5,ALARM AREA 6,ALARM AREA 7,ALARM AREA 8,ALARM AREA 9,ALARM AREA 10,ALARM AREA 11,ALARM AREA 12,ALARM AREA 13,ALARM AREA 14,
'   ALARM AREA 15,USER FIELD 1,USER FIELD 2,ESIG TYPE,ESIG ALLOW CONT USE,ESIG XMPT ALARM ACK,ESIG UNSIGNED WRITES,PDR Update Rate,PDR Access Time,PDR Deadband,
'   PDR Latch,PDR Disable Output,PDR Array Length,Hist Description,Hist Collect,Hist Interval,Hist Offset,Hist Time Res,Hist Compress,Hist Deadband,
'   Hist Comp Type,Hist Comp Time]
' !A_NAME,A_TAG,A_DESC,A_IODV,A_IOHT,A_NUMS,A_IOAD,A_OUT,A_INV,A_OPENDESC,
'   A_CLOSEDESC,A_EVENT,A_ADI,A_SA1,A_SA2,A_SA3,A_AREA1,A_AREA2,A_AREA3,A_AREA4,
'   A_AREA5,A_AREA6,A_AREA7,A_AREA8,A_AREA9,A_AREA10,A_AREA11,A_AREA12,A_AREA13,A_AREA14,
'   A_AREA15,A_ALMEXT1,A_ALMEXT2,A_ESIGTYPE,A_ESIGCONT,A_ESIGACK,A_ESIGTRAP,A_PDR_UPDATERATE,A_PDR_ACCESSTIME,A_PDR_DEADBAND,
'   A_PDR_LATCHDATA,A_PDR_DISABLEOUT,A_PDR_ARRAYLENGTH,A_HIST_DESC,A_HIST_COLLECT,A_HIST_INTERVAL,A_HIST_OFFSET,A_HIST_TIMERES,A_HIST_COMPRESS,A_HIST_DEADBAND,
'   A_HIST_COMPTYPE,A_HIST_COMPTIME!
' "DR","<TAGNAME>","<DESCRIPTION>","MBE","","DECIMAL","<ADDRESS>","NO","NO","<OFF CODE>",
'   "<ON CODE>","DISABLE","NONE","NONE","NONE","NONE","ALL","","","",
'   "","","","","","","","","","",
'   "","","","NONE","YES","NO","REJECT","1,000","300,000","0",
'   "NO","NO","0","","NO","5,000.0","0.0","Milliseconds","DISABLE","0.0000",
'   "Absolute","0"

'================== MDI FORMAT ==================
' Two header rows of formatting and preamble before the data rows. All fields in data rows must contain quotation marks surrounding the data.
' [BLOCK TYPE,TAG,NEXT BLK,DESCRIPTION,ENABLE DEVICE 1,ENABLE DEVICE 2,DEVICE0,OPTIONS0,ADDRESS0,DEVICE1,
'   OPTIONS1,ADDRESS1,DEVICE2,OPTIONS2,ADDRESS2,INITIAL SCAN,SCAN TIME,INITIAL A/M STATUS,REALARM,LABEL0,
'   LABEL1,LABEL2,LABEL3,LABEL4,LABEL5,LABEL6,LABEL7,ALM0,ALM1,ALM2,
'   ALM3,ALM4,ALM5,ALM6,ALM7,ALARM ENABLE,ALARM AREA(S),ALARM PRIORITY,SECURITY AREA 1,SECURITY AREA 2,
'   SECURITY AREA 3,ALARM AREA 1,ALARM AREA 2,ALARM AREA 3,ALARM AREA 4,ALARM AREA 5,ALARM AREA 6,ALARM AREA 7,ALARM AREA 8,ALARM AREA 9,
'   ALARM AREA 10,ALARM AREA 11,ALARM AREA 12,ALARM AREA 13,ALARM AREA 14,ALARM AREA 15,USER FIELD 1,USER FIELD 2,ESIG TYPE,ESIG ALLOW CONT USE,
'   ESIG XMPT ALARM ACK,ESIG TRAP LEVEL,PDR Update Rate,PDR Access Time,PDR Deadband,PDR Latch,PDR Disable Output,PDR Array Length,Hist Description,Hist Collect,
'   Hist Interval,Hist Offset,Hist Time Res,Hist Compress,Hist Deadband,Hist Comp Type,Hist Comp Time]
' !A_NAME,A_TAG,A_NEXT,A_DESC,A_IO1,A_IO2,A_DIDV0,A_DIHT0,A_DIAD0,A_DIDV1,
'   A_DIHT1,A_DIAD1,A_DIDV2,A_DIHT2,A_DIAD2,A_ISCAN,A_SCANT,A_IAM,A_REALM,A_TS0,
'   A_TS1,A_TS2,A_TS3,A_TS4,A_TS5,A_TS6,A_TS7,A_AE0,A_AE1,A_AE2,
'   A_AE3,A_AE4,A_AE5,A_AE6,A_AE7,A_IENAB,A_ADI,A_PRI,A_SA1,A_SA2,
'   A_SA3,A_AREA1,A_AREA2,A_AREA3,A_AREA4,A_AREA5,A_AREA6,A_AREA7,A_AREA8,A_AREA9,
'   A_AREA10,A_AREA11,A_AREA12,A_AREA13,A_AREA14,A_AREA15,A_ALMEXT1,A_ALMEXT2,A_ESIGTYPE,A_ESIGCONT,
'   A_ESIGACK,A_ESIGTRAP,A_PDR_UPDATERATE,A_PDR_ACCESSTIME,A_PDR_DEADBAND,A_PDR_LATCHDATA,A_PDR_DISABLEOUT,A_PDR_ARRAYLENGTH,A_HIST_DESC,A_HIST_COLLECT,
'   A_HIST_INTERVAL,A_HIST_OFFSET,A_HIST_TIMERES,A_HIST_COMPRESS,A_HIST_DEADBAND,A_HIST_COMPTYPE,A_HIST_COMPTIME!
' "MDI","<TAGNAME>","","<DESCRIPTION>","NO","NO","SIM","","0:0","",
'   "","","","","","ON","1","AUTO","NO","",
'   "","","","","","","","NO","NO","NO",
'   "NO","NO","NO","NO","NO","ENABLE","NONE","LOW","NONE","NONE",
'   "NONE","ALL","","","","","","","","",
'   "","","","","","","","","NONE","YES",
'   "NO","REJECT","1,000","300,000","0","NO","NO","0","","NO",
'   "5,000.0","0.0","Milliseconds","DISABLE","0.0000","Absolute","0"

'================== AA FORMAT ==================
' Two header rows of formatting and preamble before the data rows. All fields in data rows must contain quotation marks surrounding the data.
' [BLOCK TYPE,TAG,NEXT BLOCK,DESCRIPTION,INITIAL SCAN,SCAN TIME,SMOOTHING,I/O DEVICE,H/W OPTIONS,I/O ADDRESS,
'   SIGNAL CONDITIONING,LOW EGU LIMIT,HIGH EGU LIMIT,DEAD BAND,EGU TAG,ALARM ENABLE,INITIAL A/M STATUS,REMOTE ACK TAG,ALARM SUSPENSION TAG,TARGET VALUE,
'   ALARM AREA(S),ENABLE OUTPUT,SECURITY AREA 1,SECURITY AREA 2,SECURITY AREA 3,OTHER ALARM PRIORITY,OTHER CC-TAG,OTHER CC-MODE,OTHER DELAY TIME,OTHER RE-ALARM TIME,
'   HIHI ALARM LIMIT,HIHI ALARM PRIORITY,HIHI CC-TAG,HIHI CC-MODE,HIHI DELAY TIME,HIHI RE-ALARM TIME,LOLO ALARM LIMIT,LOLO ALARM PRI RITY.,LOLO CC-TAG,LOLO CC-MODE,
'   LOLO DELAY TIME,LOLO RE-ALARM TIME,HI ALARM LIMIT,HI ALARM PRIORITY,HI CC-TAG,HI CC-MODE,HI DELAY TIME,HI RE-ALARM TIME,LO ALARM LIMIT,LO ALARM PRIORITY,
'   LO CC-TAG,LO CC-MODE,LO DELAY TIME,LO RE-ALARM TIME,ROC ALARM LIMIT,ROC ALARM PRIORITY,ROC CC-TAG,ROC CC-MODE,ROC DELAY TIME,ROC RE-ALARM TIME,
'   DEV ALARM LIMIT,DEV ALARM PRIORITY,DEV CC-TAG,DEV CC-MODE,DEV DELAY TIME,DEV RE-ALARM TIME,ALARM AREA 1,ALARM AREA 2,ALARM AREA 3,ALARM AREA 4,
'   ALARM AREA 5,ALARM AREA 6,ALARM AREA 7,ALARM AREA 8,ALARM AREA 9,ALARM AREA 10,ALARM AREA 11,ALARM AREA 12,ALARM AREA 13,ALARM AREA 14,
'   ALARM AREA 15,USER FIELD 1,USER FIELD 2,ESIG TYPE,ESIG ALLOW CONT USE,ESIG XMPT ALARM ACK,ESIG UNSIGNED WRITES,PDR Update Rate,PDR Access Time,PDR Deadband,
'   PDR Latch,PDR Disable Output,PDR Array Length,Hist Description,Hist Collect,Hist Interval,Hist Offset,Hist Time Res,Hist Compress,Hist Deadband,
'   Hist Comp Type,Hist Comp Time,Scale Enabled,Scale Clamping,Scale Use EGU,Scale Raw Low,Scale Raw High,Scale Low,Scale High,Continuous Contact Output,
'   Suppress COMM Alarm]
' !A_NAME,A_TAG,A_NEXT,A_DESC,A_ISCAN,A_SCANT,A_SMOTH,A_IODV,A_IOHT,A_IOAD,
'   A_IOSC,A_ELO,A_EHI,A_DBAND,A_EGUDESC,A_IENAB,A_IAM,A_AACK,A_ASUSP,A_TRTAG,
'   A_ADI,A_EOUT,A_SA1,A_SA2,A_SA3,A_OTPRI,A_CCOTH,A_MDEOT,A_DTOTH,A_RATOT,
'   A_HIHI,A_HHPRI,A_CCHH,A_MDEHH,A_DTHH,A_RATHH,A_LOLO,A_LLPRI,A_CCLL,A_MDELL,
'   A_DTLL,A_RATLL,A_HI,A_HIPRI,A_CCHI,A_MDEHI,A_DTHI,A_RATHI,A_LO,A_LOPRI,
'   A_CCLO,A_MDELO,A_DTLO,A_RATLO,A_ROC,A_RCPRI,A_CCROC,A_MDERC,A_DTROC,A_RATRC,
'   A_DEV,A_DVPRI,A_CCDEV,A_MDEDV,A_DTDEV,A_RATDV,A_AREA1,A_AREA2,A_AREA3,A_AREA4,
'   A_AREA5,A_AREA6,A_AREA7,A_AREA8,A_AREA9,A_AREA10,A_AREA11,A_AREA12,A_AREA13,A_AREA14,
'   A_AREA15,A_ALMEXT1,A_ALMEXT2,A_ESIGTYPE,A_ESIGCONT,A_ESIGACK,A_ESIGTRAP,A_PDR_UPDATERATE,A_PDR_ACCESSTIME,A_PDR_DEADBAND,
'   A_PDR_LATCHDATA,A_PDR_DISABLEOUT,A_PDR_ARRAYLENGTH,A_HIST_DESC,A_HIST_COLLECT,A_HIST_INTERVAL,A_HIST_OFFSET,A_HIST_TIMERES,A_HIST_COMPRESS,A_HIST_DEADBAND,
'   A_HIST_COMPTYPE,A_HIST_COMPTIME,A_SCALE_ENABLED,A_SCALE_CLAMP,A_SCALE_USEEGU,A_SCALE_RAWLOW,A_SCALE_RAWHIGH,A_SCALE_LOW,A_SCALE_HIGH,A_CTK_PERSIST,
'   A_COMM_SURPRESS!
' "AA","<TAGNAME>","","<DESCRIPTION>","ON","2","0","MBE","<DATA TYPE>","<ADDRESS>",
'   "None","<Sensor Lo Limit (Eng Units)>","<Sensor HI Limit (Eng Units)>","0.00","<UNITS>","ENABLE","AUTO","","","",
'   "NONE","NO","NONE","NONE","NONE","LOW","","","00:00:00:00","00:00:00:00",
'   "85.00","LOW","","","00:00:00:00","00:00:00:00","15.00","LOW","","",
'   "00:00:00:00","00:00:00:00","75.00","LOW","","","00:00:00:00","00:00:00:00","25.00","LOW",
'   "","","00:00:00:00","00:00:00:00","","LOW","","","","",
'   "","LOW","","","","","ALL","","","",
'   "","","","","","","","","","",
'   "","","","NONE","YES","NO","REJECT","1,000","300,000","0",
'   "NO","NO","0","","NO","5,000.0","0.0","Milliseconds","DISABLE","0.0000",
'   "Absolute","0","NO","NO","YES","0.00","65,535.00","0.00","100.00","DISABLE",
'   "DISABLE"

'================== AR FORMAT ==================
' Two header rows of formatting and preamble before the data rows. All fields in data rows must contain quotation marks surrounding the data.
' [BLOCK TYPE,TAG,DESCRIPTION,I/O DEVICE,H/W OPTIONS,I/O ADDRESS TYPE,I/O ADDRESS,SIGNAL CONDITIONING,LOW EGU LIMIT,HIGH EGU LIMIT,
'   EGU TAG,OUTPUT ENABLE,EVENT MESSAGES,ALARM AREA(S),SECURITY AREA 1,SECURITY AREA 2,SECURITY AREA 3,ALARM AREA 1,ALARM AREA 2,ALARM AREA 3,
'   ALARM AREA 4,ALARM AREA 5,ALARM AREA 6,ALARM AREA 7,ALARM AREA 8,ALARM AREA 9,ALARM AREA 10,ALARM AREA 11,ALARM AREA 12,ALARM AREA 13,
'   ALARM AREA 14,ALARM AREA 15,USER FIELD 1,USER FIELD 2,ESIG TYPE,ESIG ALLOW CONT USE,ESIG XMPT ALARM ACK,ESIG UNSIGNED WRITES,PDR Update Rate,PDR Access Time,
'   PDR Deadband,PDR Latch,PDR Disable Output,PDR Array Length,Hist Description,Hist Collect,Hist Interval,Hist Offset,Hist Time Res,Hist Compress,
'   Hist Deadband,Hist Comp Type,Hist Comp Time,Scale Enabled,Scale Clamping,Scale Use EGU,Scale Raw Low,Scale Raw High,Scale Low,Scale High]
' !A_NAME,A_TAG,A_DESC,A_IODV,A_IOHT,A_NUMS,A_IOAD,A_IOSC,A_ELO,A_EHI,
'   A_EGUDESC,A_OUT,A_EVENT,A_ADI,A_SA1,A_SA2,A_SA3,A_AREA1,A_AREA2,A_AREA3,
'   A_AREA4,A_AREA5,A_AREA6,A_AREA7,A_AREA8,A_AREA9,A_AREA10,A_AREA11,A_AREA12,A_AREA13,
'   A_AREA14,A_AREA15,A_ALMEXT1,A_ALMEXT2,A_ESIGTYPE,A_ESIGCONT,A_ESIGACK,A_ESIGTRAP,A_PDR_UPDATERATE,A_PDR_ACCESSTIME,
'   A_PDR_DEADBAND,A_PDR_LATCHDATA,A_PDR_DISABLEOUT,A_PDR_ARRAYLENGTH,A_HIST_DESC,A_HIST_COLLECT,A_HIST_INTERVAL,A_HIST_OFFSET,A_HIST_TIMERES,A_HIST_COMPRESS,
'   A_HIST_DEADBAND,A_HIST_COMPTYPE,A_HIST_COMPTIME,A_SCALE_ENABLED,A_SCALE_CLAMP,A_SCALE_USEEGU,A_SCALE_RAWLOW,A_SCALE_RAWHIGH,A_SCALE_LOW,A_SCALE_HIGH!
' "AR","<TAGNAME>","<DESCRIPTION>","MBE","<DATA TYPE>","DECIMAL","<ADDRESS>","None","<Sensor Lo Limit (Eng Units)>","<Sensor HI Limit (Eng Units)>",
'   "<UNITS>","NO","DISABLE","NONE","NONE","NONE","NONE","ALL","","",
'   "","","","","","","","","","",
'   "","","","","NONE","YES","NO","REJECT","1,000","300,000",
'   "0","NO","NO","0","","NO","5,000.0","0.0","Milliseconds","DISABLE",
'   "0.0000","Absolute","0","NO","NO","YES","0.00","65,535.00","0.00","100.00"

'================== FILE FORMAT ==================
' The format of the final import file should be:
' FILE HEADER 1
' FILE HEADER 2
' 
' DA HEADER 1
' DA HEADER 2
' DA DATA ROWS
' 
' DR HEADER 1
' DR HEADER 2
' DR DATA ROWS
' 
' MDI HEADER 1
' MDI HEADER 2
' MDI DATA ROWS
' 
' AA HEADER 1
' AA HEADER 2
' AA DATA ROWS
' 
' AR HEADER 1
' AR HEADER 2
' AR DATA ROWS
'
' [-------------------------------------------------End of Block List-------------------------------------------------]

Option Explicit

'================== MODULE CONFIG ==================

Private Const CFG_SOURCE_TABLE_NAME As String = "MasterV3_DB"
Private Const CFG_DESTINATION_SHEET_NAME As String = "iFix v5.8 Export"
Private Const CFG_DEFAULT_FILE_STEM As String = "ChaseLynk_ExcelCompiled"
Private Const CFG_PULL_JCC1_ATTRIBUTES As Boolean = True
Private Const CFG_DEBUG_WRITE_OUTPUT_TO_COLUMNS As Boolean = True

Private Const BLOCK_DA As String = "DA"
Private Const BLOCK_DR As String = "DR"
Private Const BLOCK_MDI As String = "MDI"
Private Const BLOCK_AA As String = "AA"
Private Const BLOCK_AR As String = "AR"

Private Const FIELD_SOURCE_BLOCK_TYPE As String = "iFIX BLOCK TYPE"
Private Const FIELD_SOURCE_SENSOR_ELO As String = "Sensor Lo Limit (Eng Units)"
Private Const FIELD_SOURCE_SENSOR_EHI As String = "Sensor HI Limit (Eng Units)"
Private Const FIELD_SOURCE_OLD_TAGNAME_PRIMARY As String = "OLD TAGNAME (IF APPLICABLE)"
Private Const FIELD_SOURCE_OLD_TAGNAME_ALT1 As String = "OLD TAGNAME IF APPLICABLE"
Private Const FIELD_SOURCE_OLD_TAGNAME_ALT2 As String = "OLD TAGNAME"
Private Const FIELD_SOURCE_TAGNAME As String = "TAGNAME"

Private Const FIELD_A_TAG As String = "A_TAG"
Private Const FIELD_A_IENAB As String = "A_IENAB"
Private Const FIELD_A_ELO As String = "A_ELO"
Private Const FIELD_A_EHI As String = "A_EHI"
Private Const FIELD_A_LOLO As String = "A_LOLO"
Private Const FIELD_A_LO As String = "A_LO"
Private Const FIELD_A_HI As String = "A_HI"
Private Const FIELD_A_HIHI As String = "A_HIHI"

Private Const VALUE_ENABLE As String = "ENABLE"
Private Const VALUE_HEADER_TAG As String = "TAG"

Private Const JCC1_PULL_BLOCK As String = BLOCK_AA
Private Const JCC1_SOURCE_SHEET_NAME As String = "JCC1DB OG"
Private Const JCC1_RANGE_NAME_PREFIX As String = "JCC1DB_OG_"
Private Const JCC1_SOURCE_TAG_PREFIX As String = "LP1"
Private Const JCC1_LOOKUP_TAG_PREFIX As String = "JCC1"

Private Const AA_PERCENT_LOLO As Double = 0.15
Private Const AA_PERCENT_LO As Double = 0.3
Private Const AA_PERCENT_HI As Double = 0.7
Private Const AA_PERCENT_HIHI As Double = 0.85
Private Const AA_CALCULATED_SIGNIFICANT_FIGURES As Long = 3
Private Const AA_MAX_SIGNIFICANT_SHIFT As Long = 15
Private Const AA_NUMERIC_OUTPUT_FORMAT As String = "0.###############"
Private Const NUMERIC_COMPARE_EPSILON As Double = 0.000001

Public Sub Export_iFix_DB_Import()
	Dim wb As Workbook
	Dim sourceTable As ListObject
	Dim destinationSheet As Worksheet
	Dim columnMap As Object
	Dim blockOrder As Variant
	Dim bucketByType As Object
	Dim countByType As Object
	Dim unknownTypeCounts As Object
	Dim jcc1LookupByType As Object
	Dim rowValues As Object
	Dim jcc1SeedFields As Object
	Dim usedJcc1Seed As Boolean
	Dim usedJcc1LookupAttempt As Boolean
	Dim rowItem As ListRow
	Dim blockType As String
	Dim outputLines As Collection
	Dim selectedSavePath As String
	Dim selectedFileStem As String
	Dim saveChoice As Variant
	Dim visibleRows As Long
	Dim exportedRows As Long
	Dim skippedRows As Long
	Dim summary As String
	Dim jcc1LookupHits As Long
	Dim jcc1LookupAttempts As Long
	Dim jcc1LookupNoTag As Long
	Dim currentStep As String
	Dim typeBucket As Collection
	Dim key As Variant

	On Error GoTo ErrorHandler

	currentStep = "Initialize workbook context"
	Set wb = ThisWorkbook
	currentStep = "Locate source table"
	Set sourceTable = FindTableInWorkbook(wb, CFG_SOURCE_TABLE_NAME)
	If sourceTable Is Nothing Then
		MsgBox "Could not find source table '" & CFG_SOURCE_TABLE_NAME & "'.", vbExclamation, "iFix Export"
		Exit Sub
	End If

	currentStep = "Build source column map"
	Set columnMap = BuildColumnMap(sourceTable)
	If Not columnMap.Exists(NormalizeKey(FIELD_SOURCE_BLOCK_TYPE)) Then
		MsgBox "Source table is missing the 'iFIX BLOCK TYPE' column.", vbExclamation, "iFix Export"
		Exit Sub
	End If

	blockOrder = GetBlockOrder()
	Set bucketByType = CreateObject("Scripting.Dictionary")
	bucketByType.CompareMode = vbTextCompare
	Set countByType = CreateObject("Scripting.Dictionary")
	countByType.CompareMode = vbTextCompare
	Set unknownTypeCounts = CreateObject("Scripting.Dictionary")
	unknownTypeCounts.CompareMode = vbTextCompare

	For Each key In blockOrder
		bucketByType.Add CStr(key), New Collection
		countByType.Add CStr(key), 0&
	Next key

	currentStep = "Build JCC1 lookup maps"
	Set jcc1LookupByType = BuildJcc1LookupByType(wb, CFG_PULL_JCC1_ATTRIBUTES)

	For Each rowItem In sourceTable.ListRows
		If IsTableRowVisible(rowItem.Range) Then
			visibleRows = visibleRows + 1
			blockType = UCase$(GetFieldFromRow(rowItem, columnMap, FIELD_SOURCE_BLOCK_TYPE))
			currentStep = "Process source row " & rowItem.Index & " (block '" & blockType & "')"

			If bucketByType.Exists(blockType) Then
				Set rowValues = BuildPlaceholderValueMap(rowItem, columnMap)
				Set jcc1SeedFields = GetJcc1SeedRow(blockType, rowValues, jcc1LookupByType, CFG_PULL_JCC1_ATTRIBUTES, usedJcc1Seed, usedJcc1LookupAttempt)
				If CFG_PULL_JCC1_ATTRIBUTES And SupportsJcc1Pull(blockType) Then
					If usedJcc1LookupAttempt Then
						jcc1LookupAttempts = jcc1LookupAttempts + 1
					Else
						jcc1LookupNoTag = jcc1LookupNoTag + 1
					End If
					If usedJcc1Seed Then
						jcc1LookupHits = jcc1LookupHits + 1
					End If
				End If
				Set typeBucket = bucketByType(blockType)
				typeBucket.Add BuildDataRowFromTemplate(blockType, rowValues, jcc1SeedFields)
				countByType(blockType) = CLng(countByType(blockType)) + 1
				exportedRows = exportedRows + 1
			Else
				skippedRows = skippedRows + 1
				If Len(blockType) > 0 Then
					unknownTypeCounts(blockType) = CLng(GetDictCount(unknownTypeCounts, blockType)) + 1
				End If
			End If
		End If
	Next rowItem

	If visibleRows = 0 Then
		MsgBox "No visible rows were found in the source table. Clear or adjust the filter and try again.", vbInformation, "iFix Export"
		Exit Sub
	End If

	currentStep = "Prompt for CSV save path"
	saveChoice = Application.GetSaveAsFilename( _
		InitialFileName:="CNL_iFix58_Import_" & Format(Now, "yyyymmdd_hhnnss") & ".csv", _
		FileFilter:="CSV Files (*.csv), *.csv")

	If VarType(saveChoice) = vbBoolean And saveChoice = False Then
		selectedSavePath = vbNullString
		selectedFileStem = CFG_DEFAULT_FILE_STEM
	Else
		selectedSavePath = CStr(saveChoice)
		selectedFileStem = GetFileStemFromPath(selectedSavePath)
		If Len(selectedFileStem) = 0 Then
			selectedFileStem = CFG_DEFAULT_FILE_STEM
		End If
	End If

	currentStep = "Build output lines"
	Set outputLines = BuildOutputLines(blockOrder, bucketByType, countByType, selectedFileStem)

	currentStep = "Write output to destination sheet"
	Set destinationSheet = GetOrCreateWorksheet(wb, CFG_DESTINATION_SHEET_NAME)
	If CFG_DEBUG_WRITE_OUTPUT_TO_COLUMNS Then
		WriteLinesToSheetAsColumns destinationSheet, outputLines
	Else
		WriteLinesToSheet destinationSheet, outputLines
	End If

	If Len(selectedSavePath) > 0 Then
		currentStep = "Write CSV file"
		SaveLinesToCsv selectedSavePath, outputLines
	End If

	summary = "Export completed." & vbCrLf & _
			  "Visible rows scanned: " & visibleRows & vbCrLf & _
			  "Rows exported: " & exportedRows & vbCrLf & _
			  "Rows skipped (blank/unknown block type): " & skippedRows & vbCrLf & _
			  "Review output sheet: " & CFG_DESTINATION_SHEET_NAME

	If Len(selectedSavePath) > 0 Then
		summary = summary & vbCrLf & "CSV file: " & selectedSavePath
	Else
		summary = summary & vbCrLf & "CSV save was canceled."
	End If

	If CFG_PULL_JCC1_ATTRIBUTES Then
		summary = summary & vbCrLf & "JCC1 attribute pull mode: ON (rows matched: " & jcc1LookupHits & _
							  ", lookups attempted: " & jcc1LookupAttempts & ", rows without lookup tag: " & jcc1LookupNoTag & ")"
		summary = summary & vbCrLf & "JCC1 named-range rows loaded by block: " & BuildJcc1LookupCountsSummary(jcc1LookupByType, blockOrder)
	Else
		summary = summary & vbCrLf & "JCC1 attribute pull mode: OFF"
	End If

	If CFG_DEBUG_WRITE_OUTPUT_TO_COLUMNS Then
		summary = summary & vbCrLf & "Output sheet mode: DEBUG columns"
	Else
		summary = summary & vbCrLf & "Output sheet mode: Standard CSV lines"
	End If

	If unknownTypeCounts.Count > 0 Then
		summary = summary & vbCrLf & "Unknown iFix block types found: " & BuildUnknownTypeSummary(unknownTypeCounts)
	End If

	MsgBox summary, vbInformation, "iFix Export"
	Exit Sub

ErrorHandler:
	MsgBox "iFix export failed (" & Err.Number & "): " & Err.Description & vbCrLf & _
		   "Source: " & Err.Source & vbCrLf & _
		   "Step: " & currentStep, vbCritical, "iFix Export"
End Sub

Private Function BuildOutputLines(ByVal blockOrder As Variant, ByVal bucketByType As Object, ByVal countByType As Object, ByVal fileStem As String) As Collection
	Dim lines As Collection
	Dim blockType As Variant
	Dim rowsForBlock As Collection
	Dim header1Fields As Variant
	Dim header2Fields As Variant
	Dim dataTemplateFields As Variant
	Dim i As Long

	Set lines = New Collection
	lines.Add BuildFileHeaderLine1(fileStem)
	lines.Add BuildFileHeaderLine2(countByType)
	lines.Add vbNullString

	For Each blockType In blockOrder
		header1Fields = GetBlockHeaderFieldArray(CStr(blockType), 1)
		header2Fields = GetBlockHeaderFieldArray(CStr(blockType), 2)
		dataTemplateFields = GetBlockDataTemplateArray(CStr(blockType))

		Call ValidateTemplateCounts(CStr(blockType), header1Fields, header2Fields, dataTemplateFields)

		lines.Add BuildBracketedHeaderLine(header1Fields, "[", "]")
		lines.Add BuildBracketedHeaderLine(header2Fields, "!", "!")

		Set rowsForBlock = bucketByType(CStr(blockType))
		For i = 1 To rowsForBlock.Count
			lines.Add CStr(rowsForBlock(i))
		Next i

		lines.Add vbNullString
	Next blockType

	lines.Add "[-------------------------------------------------End of Block List-------------------------------------------------]"
	Set BuildOutputLines = lines
End Function

Private Function BuildFileHeaderLine1(ByVal fileStem As String) As String
	BuildFileHeaderLine1 = "[NodeName : EFSVR,Database : ELMFORK,File Name : " & fileStem & ",Date : " & _
						   Format(Date, "mm/dd/yyyy") & ",Time :  " & Format(Time, "hh:nn:ss AM/PM") & "]"
End Function

Private Function BuildFileHeaderLine2(ByVal countByType As Object) As String
	BuildFileHeaderLine2 = "#AI0,AO0,DI0,AR" & CLng(GetDictCount(countByType, BLOCK_AR)) & _
						   ",DR" & CLng(GetDictCount(countByType, BLOCK_DR)) & ",RM0,MDI" & CLng(GetDictCount(countByType, BLOCK_MDI)) & _
						   ",BL0,AA" & CLng(GetDictCount(countByType, BLOCK_AA)) & ",DA" & CLng(GetDictCount(countByType, BLOCK_DA)) & _
						   ",DT0,CA0,FN0,TX0,EV0,TM0#"
End Function

Private Function BuildDataRowFromTemplate(ByVal blockType As String, ByVal rowValues As Object, ByVal seedDefaults As Object) As String
	Dim templateFields As Variant
	Dim header2Fields As Variant
	Dim fields As Variant
	Dim resolvedValue As String
	Dim headerFieldName As String
	Dim aaCalculatedThresholds As Object
	Dim normalizedBlockType As String
	Dim isAaBlock As Boolean
	Dim i As Long

	templateFields = GetBlockDataTemplateArray(blockType)
	If IsEmpty(templateFields) Then
		BuildDataRowFromTemplate = vbNullString
		Exit Function
	End If

	normalizedBlockType = NormalizeKey(blockType)
	isAaBlock = (normalizedBlockType = "AA")

	header2Fields = GetBlockHeaderFieldArray(blockType, 2)
	If isAaBlock Then
		Set aaCalculatedThresholds = BuildAaCalculatedAlarmThresholds(rowValues)
	End If

	ReDim fields(LBound(templateFields) To UBound(templateFields))
	For i = LBound(templateFields) To UBound(templateFields)
		If InStr(1, CStr(templateFields(i)), "<", vbBinaryCompare) > 0 Then
			resolvedValue = ResolveTemplateTokens(CStr(templateFields(i)), rowValues)
		Else
			resolvedValue = CStr(templateFields(i))
		End If

		headerFieldName = vbNullString
		If IsArray(header2Fields) Then
			headerFieldName = CStr(header2Fields(i))
		End If

		If Len(headerFieldName) > 0 Then
			If isAaBlock Then
				resolvedValue = GetDictValueOrDefault(aaCalculatedThresholds, headerFieldName, resolvedValue)
			End If
			resolvedValue = GetSeedDefaultByField(headerFieldName, seedDefaults, resolvedValue)
		End If

		fields(i) = QuoteCsvField(resolvedValue)
	Next i

	BuildDataRowFromTemplate = Join(fields, ",")
End Function

Private Function GetSeedDefaultByField(ByVal fieldName As String, ByVal seedDefaults As Object, ByVal fallbackValue As String) As String
	Dim fieldKey As String

	GetSeedDefaultByField = fallbackValue
	If seedDefaults Is Nothing Then
		Exit Function
	End If

	fieldKey = NormalizeKey(fieldName)
	If Len(fieldKey) = 0 Then
		Exit Function
	End If

	If seedDefaults.Exists(fieldKey) Then
		GetSeedDefaultByField = CStr(seedDefaults(fieldKey))
	End If
End Function

Private Function GetDictValueOrDefault(ByVal valueDict As Object, ByVal fieldName As String, ByVal fallbackValue As String) As String
	Dim fieldKey As String

	GetDictValueOrDefault = fallbackValue
	If valueDict Is Nothing Then
		Exit Function
	End If

	fieldKey = NormalizeKey(fieldName)
	If Len(fieldKey) = 0 Then
		Exit Function
	End If

	If valueDict.Exists(fieldKey) Then
		GetDictValueOrDefault = CStr(valueDict(fieldKey))
	End If
End Function

Private Function BuildAaCalculatedAlarmThresholds(ByVal rowValues As Object) As Object
	Dim calculatedValues As Object
	Dim eloValue As Double
	Dim ehiValue As Double
	Dim euRange As Double

	Set calculatedValues = CreateObject("Scripting.Dictionary")
	calculatedValues.CompareMode = vbTextCompare

	If Not TryParseDouble(GetMapValue(rowValues, FIELD_SOURCE_SENSOR_ELO), eloValue) Then
		Set BuildAaCalculatedAlarmThresholds = calculatedValues
		Exit Function
	End If
	If Not TryParseDouble(GetMapValue(rowValues, FIELD_SOURCE_SENSOR_EHI), ehiValue) Then
		Set BuildAaCalculatedAlarmThresholds = calculatedValues
		Exit Function
	End If

	euRange = ehiValue - eloValue
	If euRange <= 0 Then
		Set BuildAaCalculatedAlarmThresholds = calculatedValues
		Exit Function
	End If

	calculatedValues(FIELD_A_LOLO) = FormatNumericForTemplate(eloValue + (euRange * AA_PERCENT_LOLO))
	calculatedValues(FIELD_A_LO) = FormatNumericForTemplate(eloValue + (euRange * AA_PERCENT_LO))
	calculatedValues(FIELD_A_HI) = FormatNumericForTemplate(eloValue + (euRange * AA_PERCENT_HI))
	calculatedValues(FIELD_A_HIHI) = FormatNumericForTemplate(eloValue + (euRange * AA_PERCENT_HIHI))

	Set BuildAaCalculatedAlarmThresholds = calculatedValues
End Function

Private Function TryParseDouble(ByVal textValue As String, ByRef parsedValue As Double) As Boolean
	Dim cleanedText As String

	cleanedText = Trim$(textValue)
	If Len(cleanedText) = 0 Then
		Exit Function
	End If

	cleanedText = Replace(cleanedText, ",", "")
	If IsNumeric(cleanedText) Then
		parsedValue = CDbl(cleanedText)
		TryParseDouble = True
	End If
End Function

Private Function FormatNumericForTemplate(ByVal numericValue As Double) As String
	Dim roundedValue As Double
	Dim formattedValue As String

	On Error GoTo FallbackValue

	roundedValue = RoundToSignificantFigures(numericValue, AA_CALCULATED_SIGNIFICANT_FIGURES)
	formattedValue = Format$(roundedValue, AA_NUMERIC_OUTPUT_FORMAT)
	FormatNumericForTemplate = NormalizeFormattedNumericText(formattedValue)
	Exit Function

FallbackValue:
	formattedValue = Format$(numericValue, AA_NUMERIC_OUTPUT_FORMAT)
	FormatNumericForTemplate = NormalizeFormattedNumericText(formattedValue)
End Function

Private Function NormalizeFormattedNumericText(ByVal valueText As String) As String
	Dim outputText As String

	outputText = Trim$(valueText)

	Do While Len(outputText) > 0 And Right$(outputText, 1) = "0" And InStr(1, outputText, ".", vbBinaryCompare) > 0
		outputText = Left$(outputText, Len(outputText) - 1)
	Loop

	If Len(outputText) > 0 And Right$(outputText, 1) = "." Then
		outputText = Left$(outputText, Len(outputText) - 1)
	End If

	If Len(outputText) = 0 Then
		outputText = "0"
	End If

	NormalizeFormattedNumericText = outputText
End Function

Private Function RoundToSignificantFigures(ByVal numericValue As Double, ByVal significantFigures As Long) As Double
	Dim magnitude As Long
	Dim shiftPlaces As Long
	Dim scaleFactor As Double

	If significantFigures < 1 Then
		RoundToSignificantFigures = numericValue
		Exit Function
	End If

	If numericValue = 0 Then
		RoundToSignificantFigures = 0
		Exit Function
	End If

	magnitude = Int(Log(Abs(numericValue)) / Log(10#))
	shiftPlaces = significantFigures - 1 - magnitude

	If shiftPlaces > AA_MAX_SIGNIFICANT_SHIFT Then
		shiftPlaces = AA_MAX_SIGNIFICANT_SHIFT
	ElseIf shiftPlaces < -AA_MAX_SIGNIFICANT_SHIFT Then
		shiftPlaces = -AA_MAX_SIGNIFICANT_SHIFT
	End If

	If shiftPlaces >= 0 Then
		scaleFactor = 10# ^ shiftPlaces
		RoundToSignificantFigures = Round(numericValue * scaleFactor, 0) / scaleFactor
	Else
		scaleFactor = 10# ^ (-shiftPlaces)
		RoundToSignificantFigures = Round(numericValue / scaleFactor, 0) * scaleFactor
	End If
End Function

Private Function AreValuesEquivalent(ByVal leftValue As String, ByVal rightValue As String) As Boolean
	Dim leftNumber As Double
	Dim rightNumber As Double

	If TryParseDouble(leftValue, leftNumber) And TryParseDouble(rightValue, rightNumber) Then
		AreValuesEquivalent = (Abs(leftNumber - rightNumber) < NUMERIC_COMPARE_EPSILON)
	Else
		AreValuesEquivalent = (NormalizeKey(NormalizeImportedTemplateValue(leftValue)) = NormalizeKey(NormalizeImportedTemplateValue(rightValue)))
	End If
End Function

'================== JCC1 LOOKUP HELPERS ==================

Private Function BuildJcc1LookupByType(ByVal wb As Workbook, ByVal pullEnabled As Boolean) As Object
	Dim lookupByType As Object
	Dim ogSheet As Worksheet
	Dim blockTagMap As Object
	Dim allowedFieldMap As Object

	Set lookupByType = CreateObject("Scripting.Dictionary")
	lookupByType.CompareMode = vbTextCompare

	If Not pullEnabled Then
		Set BuildJcc1LookupByType = lookupByType
		Exit Function
	End If

	On Error Resume Next
	Set ogSheet = wb.Worksheets(JCC1_SOURCE_SHEET_NAME)
	On Error GoTo 0
	If ogSheet Is Nothing Then
		Set BuildJcc1LookupByType = lookupByType
		Exit Function
	End If

	Set allowedFieldMap = GetJcc1AllowedFieldMap()
	Set blockTagMap = BuildTagToFieldMapFromNamedRange(wb, ogSheet, JCC1_RANGE_NAME_PREFIX & JCC1_PULL_BLOCK, allowedFieldMap)
	If blockTagMap.Count > 0 Then
		PutObjectInDictionary lookupByType, JCC1_PULL_BLOCK, blockTagMap
	End If

	Set BuildJcc1LookupByType = lookupByType
End Function

Private Function BuildTagToFieldMapFromNamedRange(ByVal wb As Workbook, ByVal ogSheet As Worksheet, ByVal rangeName As String, ByVal allowedFieldMap As Object) As Object
	Dim tagMap As Object
	Dim sourceRange As Range
	Dim rangeValues As Variant
	Dim colFieldKeys() As String
	Dim tagCol As Long
	Dim rowIx As Long
	Dim colIx As Long
	Dim thisFieldKey As String
	Dim tagValue As String
	Dim sourceValue As String
	Dim rowFieldMap As Object

	Set tagMap = CreateObject("Scripting.Dictionary")
	tagMap.CompareMode = vbTextCompare

	Set sourceRange = ResolveNamedRangeReference(wb, ogSheet, rangeName)
	If sourceRange Is Nothing Then
		Set BuildTagToFieldMapFromNamedRange = tagMap
		Exit Function
	End If

	rangeValues = sourceRange.Value2
	If Not IsArray(rangeValues) Then
		Set BuildTagToFieldMapFromNamedRange = tagMap
		Exit Function
	End If
	If UBound(rangeValues, 1) < 3 Then
		Set BuildTagToFieldMapFromNamedRange = tagMap
		Exit Function
	End If

	ReDim colFieldKeys(1 To UBound(rangeValues, 2))
	For colIx = 1 To UBound(rangeValues, 2)
		thisFieldKey = ParseJcc1HeaderKey(rangeValues(2, colIx))
		If Len(thisFieldKey) > 0 Then
			If allowedFieldMap Is Nothing Or allowedFieldMap.Exists(thisFieldKey) Then
				colFieldKeys(colIx) = thisFieldKey
			End If
		End If
	Next colIx

	tagCol = FindTagColumnInJcc1Range(rangeValues, colFieldKeys)
	If tagCol = 0 Then
		tagCol = 2
	End If

	For rowIx = 3 To UBound(rangeValues, 1)
		tagValue = ReadRangeCellRawText(sourceRange, rowIx, tagCol)
		If Len(tagValue) > 0 Then
			Set rowFieldMap = CreateObject("Scripting.Dictionary")
			rowFieldMap.CompareMode = vbTextCompare
			For colIx = 1 To UBound(rangeValues, 2)
				If Len(colFieldKeys(colIx)) > 0 Then
					sourceValue = ReadRangeCellRawText(sourceRange, rowIx, colIx)
					rowFieldMap.Item(colFieldKeys(colIx)) = sourceValue
				End If
			Next colIx
			PutObjectInDictionary tagMap, NormalizeTagKey(tagValue), rowFieldMap
		End If
	Next rowIx

	Set BuildTagToFieldMapFromNamedRange = tagMap
End Function

Private Function SupportsJcc1Pull(ByVal blockType As String) As Boolean
	SupportsJcc1Pull = (NormalizeKey(blockType) = JCC1_PULL_BLOCK)
End Function

Private Function GetJcc1AllowedFieldMap() As Object
	Dim fieldMap As Object
	Dim fieldName As Variant

	Set fieldMap = CreateObject("Scripting.Dictionary")
	fieldMap.CompareMode = vbTextCompare

	For Each fieldName In GetAaHistoricalLookupFieldNames()
		fieldMap(CStr(fieldName)) = True
	Next fieldName

	Set GetJcc1AllowedFieldMap = fieldMap
End Function

Private Function ReadRangeCellRawText(ByVal sourceRange As Range, ByVal rowIx As Long, ByVal colIx As Long) As String
	On Error Resume Next
	ReadRangeCellRawText = CStr(sourceRange.Cells(rowIx, colIx).Text)
	On Error GoTo 0
End Function

Private Function ResolveNamedRangeReference(ByVal wb As Workbook, ByVal sourceSheet As Worksheet, ByVal rangeName As String) As Range
	Dim nm As Name
	Dim nmKey As String

	For Each nm In wb.Names
		nmKey = GetBareNameFromFullName(CStr(nm.Name))
		If StrComp(nmKey, rangeName, vbTextCompare) = 0 Then
			On Error Resume Next
			Set ResolveNamedRangeReference = nm.RefersToRange
			On Error GoTo 0
			If Not ResolveNamedRangeReference Is Nothing Then
				Exit Function
			End If
		End If
	Next nm

	For Each nm In sourceSheet.Names
		nmKey = GetBareNameFromFullName(CStr(nm.Name))
		If StrComp(nmKey, rangeName, vbTextCompare) = 0 Then
			On Error Resume Next
			Set ResolveNamedRangeReference = nm.RefersToRange
			On Error GoTo 0
			If Not ResolveNamedRangeReference Is Nothing Then
				Exit Function
			End If
		End If
	Next nm

	On Error Resume Next
	Set ResolveNamedRangeReference = sourceSheet.Range(rangeName)
	On Error GoTo 0
End Function

Private Function GetBareNameFromFullName(ByVal fullName As String) As String
	Dim bangPos As Long
	Dim nameText As String

	nameText = fullName
	bangPos = InStrRev(nameText, "!", -1, vbBinaryCompare)
	If bangPos > 0 Then
		nameText = Mid$(nameText, bangPos + 1)
	End If

	nameText = Replace(nameText, "'", "")
	nameText = Replace(nameText, "=", "")
	GetBareNameFromFullName = Trim$(nameText)
End Function

Private Function FindTagColumnInJcc1Range(ByVal rangeValues As Variant, ByRef colFieldKeys() As String) As Long
	Dim colIx As Long
	Dim header1 As String

	For colIx = 1 To UBound(rangeValues, 2)
		header1 = ParseJcc1HeaderKey(rangeValues(1, colIx))

		If header1 = VALUE_HEADER_TAG Or colFieldKeys(colIx) = FIELD_A_TAG Then
			FindTagColumnInJcc1Range = colIx
			Exit Function
		End If
	Next colIx
End Function

Private Function ParseJcc1HeaderKey(ByVal value As Variant) As String
	Dim cleaned As String

	cleaned = NormalizeImportedTemplateValue(value)
	cleaned = Replace(cleaned, "[", "")
	cleaned = Replace(cleaned, "]", "")
	cleaned = Replace(cleaned, "!", "")
	ParseJcc1HeaderKey = NormalizeKey(cleaned)
End Function

Private Function ShouldAttemptJcc1Lookup(ByVal blockType As String, ByVal rowValues As Object, ByVal jcc1LookupByType As Object, ByVal pullEnabled As Boolean, ByRef blockTagMap As Object, ByRef normalizedTag As String) As Boolean
	Dim lookupTag As String

	Set blockTagMap = Nothing
	normalizedTag = vbNullString

	If Not pullEnabled Then Exit Function
	If Not SupportsJcc1Pull(blockType) Then Exit Function
	If jcc1LookupByType Is Nothing Then Exit Function
	If Not jcc1LookupByType.Exists(blockType) Then Exit Function

	lookupTag = BuildJcc1LookupTag(GetSourceTagForJcc1Lookup(rowValues))
	If Len(lookupTag) = 0 Then Exit Function

	Set blockTagMap = jcc1LookupByType.Item(blockType)
	normalizedTag = NormalizeTagKey(lookupTag)
	ShouldAttemptJcc1Lookup = True
End Function

Private Function GetJcc1SeedRow(ByVal blockType As String, ByVal rowValues As Object, ByVal jcc1LookupByType As Object, ByVal pullEnabled As Boolean, ByRef foundMatch As Boolean, Optional ByRef lookupAttempted As Boolean = False) As Object
	Dim normalizedTag As String
	Dim blockTagMap As Object
	Dim historicalSeed As Object

	Set GetJcc1SeedRow = Nothing
	foundMatch = False
	lookupAttempted = False
	If Not ShouldAttemptJcc1Lookup(blockType, rowValues, jcc1LookupByType, pullEnabled, blockTagMap, normalizedTag) Then
		Exit Function
	End If

	lookupAttempted = True
	If blockTagMap.Exists(normalizedTag) Then
		Set historicalSeed = blockTagMap.Item(normalizedTag)
		If IsAaHistoricalSeedEligible(rowValues, historicalSeed) Then
			Set GetJcc1SeedRow = ExtractAaHistoricalAlarmSeed(historicalSeed, rowValues)
			If Not GetJcc1SeedRow Is Nothing Then
				If GetJcc1SeedRow.Count > 0 Then
					foundMatch = True
				Else
					Set GetJcc1SeedRow = Nothing
				End If
			End If
		End If
	End If
End Function

Private Function IsAaHistoricalSeedEligible(ByVal rowValues As Object, ByVal historicalSeed As Object) As Boolean
	Dim masterElo As String
	Dim masterEhi As String
	Dim seedElo As String
	Dim seedEhi As String
	Dim alarmEnabled As String

	If historicalSeed Is Nothing Then
		Exit Function
	End If

	alarmEnabled = GetDictValueOrDefault(historicalSeed, FIELD_A_IENAB, vbNullString)
	If NormalizeKey(alarmEnabled) <> VALUE_ENABLE Then
		Exit Function
	End If

	masterElo = GetMapValue(rowValues, FIELD_SOURCE_SENSOR_ELO)
	masterEhi = GetMapValue(rowValues, FIELD_SOURCE_SENSOR_EHI)
	seedElo = GetDictValueOrDefault(historicalSeed, FIELD_A_ELO, vbNullString)
	seedEhi = GetDictValueOrDefault(historicalSeed, FIELD_A_EHI, vbNullString)

	If Not AreValuesEquivalent(masterElo, seedElo) Then
		Exit Function
	End If
	If Not AreValuesEquivalent(masterEhi, seedEhi) Then
		Exit Function
	End If

	IsAaHistoricalSeedEligible = True
End Function

Private Function ExtractAaHistoricalAlarmSeed(ByVal historicalSeed As Object, ByVal rowValues As Object) As Object
	Dim selectedSeed As Object
	Dim fieldName As Variant
	Dim seedValue As String
	Dim seedNumericValue As Double
	Dim sourceElo As Double
	Dim sourceEhi As Double

	Set selectedSeed = CreateObject("Scripting.Dictionary")
	selectedSeed.CompareMode = vbTextCompare

	If Not TryParseDouble(GetMapValue(rowValues, FIELD_SOURCE_SENSOR_ELO), sourceElo) Then
		Set ExtractAaHistoricalAlarmSeed = selectedSeed
		Exit Function
	End If

	If Not TryParseDouble(GetMapValue(rowValues, FIELD_SOURCE_SENSOR_EHI), sourceEhi) Then
		Set ExtractAaHistoricalAlarmSeed = selectedSeed
		Exit Function
	End If

	If sourceEhi < sourceElo Then
		Set ExtractAaHistoricalAlarmSeed = selectedSeed
		Exit Function
	End If

	For Each fieldName In GetAaAlarmThresholdFieldNames()
		seedValue = GetDictValueOrDefault(historicalSeed, CStr(fieldName), vbNullString)
		If Len(seedValue) > 0 Then
			If TryParseDouble(seedValue, seedNumericValue) Then
				If seedNumericValue >= sourceElo And seedNumericValue <= sourceEhi Then
					selectedSeed(CStr(fieldName)) = seedValue
				End If
			End If
		End If
	Next fieldName

	Set ExtractAaHistoricalAlarmSeed = selectedSeed
End Function

Private Function GetSourceTagForJcc1Lookup(ByVal rowValues As Object) As String
	Dim key As Variant
	Dim normalizedKey As String

	GetSourceTagForJcc1Lookup = GetMapValue(rowValues, FIELD_SOURCE_OLD_TAGNAME_PRIMARY)
	If Len(GetSourceTagForJcc1Lookup) > 0 Then Exit Function

	GetSourceTagForJcc1Lookup = GetMapValue(rowValues, FIELD_SOURCE_OLD_TAGNAME_ALT1)
	If Len(GetSourceTagForJcc1Lookup) > 0 Then Exit Function

	GetSourceTagForJcc1Lookup = GetMapValue(rowValues, FIELD_SOURCE_OLD_TAGNAME_ALT2)
	If Len(GetSourceTagForJcc1Lookup) > 0 Then Exit Function

	For Each key In rowValues.Keys
		normalizedKey = NormalizeKey(CStr(key))
		If InStr(1, normalizedKey, "OLD", vbTextCompare) > 0 And _
		   InStr(1, normalizedKey, "TAG", vbTextCompare) > 0 And _
		   InStr(1, normalizedKey, "NAME", vbTextCompare) > 0 Then
			GetSourceTagForJcc1Lookup = CStr(rowValues.Item(CStr(key)))
			If Len(GetSourceTagForJcc1Lookup) > 0 Then Exit Function
		End If
	Next key

	GetSourceTagForJcc1Lookup = GetMapValue(rowValues, FIELD_SOURCE_TAGNAME)
End Function

Private Function BuildJcc1LookupTag(ByVal oldTagName As String) As String
	Dim convertedTag As String
	Dim normalizedOldTag As String

	normalizedOldTag = NormalizeImportedTemplateValue(oldTagName)
	If IsLikelyNewTagMarker(normalizedOldTag) Then
		BuildJcc1LookupTag = vbNullString
		Exit Function
	End If

	convertedTag = Trim$(Replace(normalizedOldTag, JCC1_SOURCE_TAG_PREFIX, JCC1_LOOKUP_TAG_PREFIX, 1, -1, vbTextCompare))
	If IsLikelyNewTagMarker(convertedTag) Then
		BuildJcc1LookupTag = vbNullString
	Else
		BuildJcc1LookupTag = convertedTag
	End If
End Function

Private Function GetBlockOrder() As Variant
	GetBlockOrder = Array(BLOCK_DA, BLOCK_DR, BLOCK_MDI, BLOCK_AA, BLOCK_AR)
End Function

Private Function GetAaAlarmThresholdFieldNames() As Variant
	GetAaAlarmThresholdFieldNames = Array(FIELD_A_LOLO, FIELD_A_LO, FIELD_A_HI, FIELD_A_HIHI)
End Function

Private Function GetAaHistoricalLookupFieldNames() As Variant
	GetAaHistoricalLookupFieldNames = Array(FIELD_A_TAG, FIELD_A_IENAB, FIELD_A_ELO, FIELD_A_EHI, FIELD_A_LOLO, FIELD_A_LO, FIELD_A_HI, FIELD_A_HIHI)
End Function

Private Function NormalizeTagKey(ByVal value As String) As String
	Dim outputText As String

	outputText = NormalizeKey(StripOuterQuotes(value))
	outputText = Replace(outputText, Chr$(160), "")
	outputText = Replace(outputText, vbTab, "")
	outputText = Replace(outputText, " ", "")
	NormalizeTagKey = outputText
End Function

Private Function BuildJcc1LookupCountsSummary(ByVal jcc1LookupByType As Object, ByVal blockOrder As Variant) As String
	Dim blockType As Variant
	Dim pieces As String
	Dim thisCount As Long
	Dim blockTagMap As Object

	For Each blockType In blockOrder
		thisCount = 0
		If Not jcc1LookupByType Is Nothing Then
			If jcc1LookupByType.Exists(CStr(blockType)) Then
				Set blockTagMap = Nothing
				On Error Resume Next
				Set blockTagMap = jcc1LookupByType.Item(CStr(blockType))
				On Error GoTo 0
				If Not blockTagMap Is Nothing Then
					thisCount = CLng(blockTagMap.Count)
				End If
			End If
		End If

		If Len(pieces) > 0 Then
			pieces = pieces & "; "
		End If
		pieces = pieces & CStr(blockType) & "=" & thisCount
	Next blockType

	BuildJcc1LookupCountsSummary = pieces
End Function

Private Sub PutObjectInDictionary(ByVal dict As Object, ByVal key As String, ByVal valueObject As Object)
	If dict.Exists(key) Then
		dict.Remove key
	End If
	dict.Add key, valueObject
End Sub

Private Function IsLikelyNewTagMarker(ByVal value As String) As Boolean
	Dim normalized As String

	normalized = NormalizeKey(value)
	If Len(normalized) = 0 Then
		IsLikelyNewTagMarker = True
		Exit Function
	End If

	If normalized = "NA" Or normalized = "N/A" Or normalized = "NEW" Or normalized = "NEW TAG" Or normalized = "NEW TAGS" Then
		IsLikelyNewTagMarker = True
		Exit Function
	End If

	If InStr(1, normalized, "N/A", vbTextCompare) > 0 Then
		IsLikelyNewTagMarker = True
	End If
End Function

Private Function NormalizeImportedTemplateValue(ByVal value As Variant) As String
	NormalizeImportedTemplateValue = StripOuterQuotes(NormalizeSourceValue(value))
End Function

Private Function StripOuterQuotes(ByVal value As String) As String
	Dim outputText As String

	outputText = value
	If Len(outputText) >= 2 Then
		If Left$(outputText, 1) = Chr$(34) And Right$(outputText, 1) = Chr$(34) Then
			outputText = Mid$(outputText, 2, Len(outputText) - 2)
		End If
	End If

	StripOuterQuotes = Replace(outputText, Chr$(34) & Chr$(34), Chr$(34))
End Function

'================== GENERAL TEXT/TOKEN HELPERS ==================

Private Function ResolveTemplateTokens(ByVal templateValue As String, ByVal rowValues As Object) As String
	Dim outputText As String
	Dim startPos As Long
	Dim endPos As Long
	Dim tokenName As String
	Dim replacement As String

	outputText = templateValue
	startPos = InStr(1, outputText, "<", vbBinaryCompare)

	Do While startPos > 0
		endPos = InStr(startPos + 1, outputText, ">", vbBinaryCompare)
		If endPos = 0 Then
			Exit Do
		End If

		tokenName = Mid$(outputText, startPos + 1, endPos - startPos - 1)
		replacement = GetMapValue(rowValues, tokenName)

		outputText = Left$(outputText, startPos - 1) & replacement & Mid$(outputText, endPos + 1)
		startPos = InStr(startPos + Len(replacement), outputText, "<", vbBinaryCompare)
	Loop

	ResolveTemplateTokens = outputText
End Function

Private Function QuoteCsvField(ByVal value As String) As String
	QuoteCsvField = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function BuildPlaceholderValueMap(ByVal rowItem As ListRow, ByVal columnMap As Object) As Object
	Dim valueMap As Object
	Dim key As Variant
	Dim colIndex As Long

	Set valueMap = CreateObject("Scripting.Dictionary")
	valueMap.CompareMode = vbTextCompare

	For Each key In columnMap.Keys
		colIndex = CLng(columnMap(CStr(key)))
		valueMap(CStr(key)) = NormalizeSourceValue(rowItem.Range.Cells(1, colIndex).Value2)
	Next key

	Set BuildPlaceholderValueMap = valueMap
End Function

Private Function GetFieldFromRow(ByVal rowItem As ListRow, ByVal columnMap As Object, ByVal columnName As String) As String
	Dim key As String
	Dim colIndex As Long

	key = NormalizeKey(columnName)
	If Not columnMap.Exists(key) Then
		GetFieldFromRow = vbNullString
		Exit Function
	End If

	colIndex = CLng(columnMap(key))
	GetFieldFromRow = NormalizeSourceValue(rowItem.Range.Cells(1, colIndex).Value2)
End Function

Private Function BuildColumnMap(ByVal sourceTable As ListObject) As Object
	Dim result As Object
	Dim listCol As ListColumn
	Dim key As String

	Set result = CreateObject("Scripting.Dictionary")
	result.CompareMode = vbTextCompare

	For Each listCol In sourceTable.ListColumns
		key = NormalizeKey(CStr(listCol.Name))
		If Len(key) > 0 Then
			result(key) = CLng(listCol.Index)
		End If
	Next listCol

	Set BuildColumnMap = result
End Function

Private Function NormalizeKey(ByVal value As String) As String
	NormalizeKey = UCase$(Trim$(value))
End Function

Private Function NormalizeSourceValue(ByVal value As Variant) As String
	Dim textValue As String

	If IsError(value) Then
		NormalizeSourceValue = vbNullString
		Exit Function
	End If

	textValue = Trim$(CStr(value))
	If textValue = "-" Then
		NormalizeSourceValue = vbNullString
	Else
		NormalizeSourceValue = textValue
	End If
End Function

Private Function GetMapValue(ByVal valueMap As Object, ByVal key As String) As String
	Dim normalized As String

	normalized = NormalizeKey(key)
	If valueMap.Exists(normalized) Then
		GetMapValue = CStr(valueMap(normalized))
	Else
		GetMapValue = vbNullString
	End If
End Function

Private Function GetDictCount(ByVal dict As Object, ByVal key As String) As Long
	If dict.Exists(key) Then
		GetDictCount = CLng(dict(key))
	Else
		GetDictCount = 0
	End If
End Function

Private Function FindTableInWorkbook(ByVal wb As Workbook, ByVal tableName As String) As ListObject
	Dim ws As Worksheet
	Dim lo As ListObject

	For Each ws In wb.Worksheets
		For Each lo In ws.ListObjects
			If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
				Set FindTableInWorkbook = lo
				Exit Function
			End If
		Next lo
	Next ws
End Function

Private Function GetOrCreateWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
	Dim ws As Worksheet

	On Error Resume Next
	Set ws = wb.Worksheets(sheetName)
	On Error GoTo 0

	If ws Is Nothing Then
		Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
		ws.Name = sheetName
	End If

	Set GetOrCreateWorksheet = ws
End Function

Private Sub WriteLinesToSheet(ByVal ws As Worksheet, ByVal lines As Collection)
	Dim outputArray() As Variant
	Dim i As Long

	ws.Cells.ClearContents
	ws.Columns("A").NumberFormat = "@"

	ReDim outputArray(1 To lines.Count, 1 To 1)
	For i = 1 To lines.Count
		outputArray(i, 1) = CStr(lines(i))
	Next i

	ws.Range("A1").Resize(lines.Count, 1).Value = outputArray
	ws.Columns("A").AutoFit
End Sub

Private Sub WriteLinesToSheetAsColumns(ByVal ws As Worksheet, ByVal lines As Collection)
	Dim outputArray() As Variant
	Dim parsedRows() As Variant
	Dim rowFields As Variant
	Dim maxCols As Long
	Dim rowCols As Long
	Dim i As Long
	Dim j As Long
	Dim baseIndex As Long

	ws.Cells.ClearContents

	ReDim parsedRows(1 To lines.Count)
	maxCols = 1

	For i = 1 To lines.Count
		parsedRows(i) = ParseCsvLineToFields(CStr(lines(i)))
		rowFields = parsedRows(i)
		rowCols = UBound(rowFields) - LBound(rowFields) + 1
		If rowCols > maxCols Then
			maxCols = rowCols
		End If
	Next i

	ReDim outputArray(1 To lines.Count, 1 To maxCols)
	For i = 1 To lines.Count
		rowFields = parsedRows(i)
		baseIndex = LBound(rowFields)
		For j = LBound(rowFields) To UBound(rowFields)
			outputArray(i, (j - baseIndex) + 1) = CStr(rowFields(j))
		Next j
	Next i

	ws.Range("A1").Resize(lines.Count, maxCols).NumberFormat = "@"
	ws.Range("A1").Resize(lines.Count, maxCols).Value = outputArray
	ws.Range(ws.Cells(1, 1), ws.Cells(1, maxCols)).EntireColumn.AutoFit
End Sub

Private Function ParseCsvLineToFields(ByVal lineText As String) As Variant
	Dim fields As Collection
	Dim currentField As String
	Dim i As Long
	Dim outputArray() As String
	Dim ch As String
	Dim inQuotes As Boolean

	Set fields = New Collection
	currentField = vbNullString

	For i = 1 To Len(lineText)
		ch = Mid$(lineText, i, 1)

		If ch = Chr$(34) Then
			If inQuotes And i < Len(lineText) And Mid$(lineText, i + 1, 1) = Chr$(34) Then
				currentField = currentField & Chr$(34)
				i = i + 1
			Else
				inQuotes = Not inQuotes
			End If
		ElseIf ch = "," And Not inQuotes Then
			fields.Add currentField
			currentField = vbNullString
		Else
			currentField = currentField & ch
		End If
	Next i

	fields.Add currentField

	ReDim outputArray(0 To fields.Count - 1)
	For i = 1 To fields.Count
		outputArray(i - 1) = CStr(fields(i))
	Next i

	ParseCsvLineToFields = outputArray
End Function

Private Sub SaveLinesToCsv(ByVal fullPath As String, ByVal lines As Collection)
	Dim fileNumber As Integer
	Dim i As Long

	fileNumber = FreeFile
	Open fullPath For Output As #fileNumber
	For i = 1 To lines.Count
		Print #fileNumber, CStr(lines(i))
	Next i
	Close #fileNumber
End Sub

Private Function IsTableRowVisible(ByVal rowRange As Range) As Boolean
	On Error Resume Next
	IsTableRowVisible = Not rowRange.EntireRow.Hidden
	On Error GoTo 0
End Function

Private Function GetFileStemFromPath(ByVal fullPath As String) As String
	Dim fileNameOnly As String
	Dim dotPos As Long

	fileNameOnly = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
	dotPos = InStrRev(fileNameOnly, ".")

	If dotPos > 1 Then
		GetFileStemFromPath = Left$(fileNameOnly, dotPos - 1)
	Else
		GetFileStemFromPath = fileNameOnly
	End If
End Function

Private Function BuildUnknownTypeSummary(ByVal unknownTypeCounts As Object) As String
	Dim key As Variant
	Dim textOut As String

	textOut = vbNullString
	For Each key In unknownTypeCounts.Keys
		If Len(textOut) > 0 Then
			textOut = textOut & "; "
		End If
		textOut = textOut & CStr(key) & " (" & CLng(unknownTypeCounts(key)) & ")"
	Next key

	BuildUnknownTypeSummary = textOut
End Function

Private Function BuildBracketedHeaderLine(ByVal headerFields As Variant, ByVal startToken As String, ByVal endToken As String) As String
	BuildBracketedHeaderLine = startToken & Join(headerFields, ",") & endToken
End Function

Private Function ValidateTemplateCounts(ByVal blockType As String, ByVal header1Fields As Variant, ByVal header2Fields As Variant, ByVal dataTemplateFields As Variant) As Boolean
	If UBound(header1Fields) <> UBound(header2Fields) Then
		Err.Raise vbObjectError + 2101, "iFix Export", "Template mismatch for block type '" & blockType & "': header field counts are different."
	End If
	If UBound(header1Fields) <> UBound(dataTemplateFields) Then
		Err.Raise vbObjectError + 2102, "iFix Export", "Template mismatch for block type '" & blockType & "': data field count does not match header."
	End If
	ValidateTemplateCounts = True
End Function

Private Function GetBlockHeaderFieldArray(ByVal blockType As String, ByVal headerNumber As Long) As Variant
	Select Case UCase$(Trim$(blockType))
		Case "DA"
			If headerNumber = 1 Then
				GetBlockHeaderFieldArray = Array("BLOCK TYPE", "TAG", "NEXT BLOCK", "DESCRIPTION", "I/O DEVICE", "H/W OPTIONS", "I/O ADDRESS", "INITIAL A/M STATUS", "INITIAL SCAN", "SCAN TIME", _
					"INVERT OUTPUT", "OPEN TAG", "CLOSE TAG", "ALARM ENABLE", "ALARM AREA(S)", "ALARM PRIORITY", "ALARM TYPE", "ENABLE OUTPUT", "SECURITY AREA 1", "SECURITY AREA 2", _
					"SECURITY AREA 3", "ALARM AREA 1", "ALARM AREA 2", "ALARM AREA 3", "ALARM AREA 4", "ALARM AREA 5", "ALARM AREA 6", "ALARM AREA 7", "ALARM AREA 8", "ALARM AREA 9", _
					"ALARM AREA 10", "ALARM AREA 11", "ALARM AREA 12", "ALARM AREA 13", "ALARM AREA 14", "ALARM AREA 15", "USER FIELD 1", "USER FIELD 2", "CONTACT TAG", "CONTACT MODE", _
					"REMOTE ACK TAG", "ALARM SUSPENSION TAG", "DELAY TIME", "REALARM TIME", "ESIG TYPE", "ESIG ALLOW CONT USE", "ESIG XMPT ALARM ACK", "ESIG UNSIGNED WRITES", "PDR Update Rate", "PDR Access Time", _
					"PDR Deadband", "PDR Latch", "PDR Disable Output", "PDR Array Length", "Hist Description", "Hist Collect", "Hist Interval", "Hist Offset", "Hist Time Res", "Hist Compress", _
					"Hist Deadband", "Hist Comp Type", "Hist Comp Time", "Continuous Contact Output", "Suppress COMM Alarm")
			ElseIf headerNumber = 2 Then
				GetBlockHeaderFieldArray = Array("A_NAME", "A_TAG", "A_NEXT", "A_DESC", "A_IODV", "A_IOHT", "A_IOAD", "A_IAM", "A_ISCAN", "A_SCANT", _
					"A_INV", "A_OPENDESC", "A_CLOSEDESC", "A_IENAB", "A_ADI", "A_PRI", "A_ALMCK", "A_EOUT", "A_SA1", "A_SA2", _
					"A_SA3", "A_AREA1", "A_AREA2", "A_AREA3", "A_AREA4", "A_AREA5", "A_AREA6", "A_AREA7", "A_AREA8", "A_AREA9", _
					"A_AREA10", "A_AREA11", "A_AREA12", "A_AREA13", "A_AREA14", "A_AREA15", "A_ALMEXT1", "A_ALMEXT2", "A_CCALM", "A_CCMOD", _
					"A_AACK", "A_ASUSP", "A_DELAY", "A_REALM", "A_ESIGTYPE", "A_ESIGCONT", "A_ESIGACK", "A_ESIGTRAP", "A_PDR_UPDATERATE", "A_PDR_ACCESSTIME", _
					"A_PDR_DEADBAND", "A_PDR_LATCHDATA", "A_PDR_DISABLEOUT", "A_PDR_ARRAYLENGTH", "A_HIST_DESC", "A_HIST_COLLECT", "A_HIST_INTERVAL", "A_HIST_OFFSET", "A_HIST_TIMERES", "A_HIST_COMPRESS", _
					"A_HIST_DEADBAND", "A_HIST_COMPTYPE", "A_HIST_COMPTIME", "A_CTK_PERSIST", "A_COMM_SURPRESS")
			Else
				GetBlockHeaderFieldArray = Empty
			End If
		Case "DR"
			If headerNumber = 1 Then
				GetBlockHeaderFieldArray = Array("BLOCK TYPE", "TAG", "DESCRIPTION", "I/O DEVICE", "H/W OPTIONS", "I/O ADDRESS TYPE", "I/O ADDRESS", "ENABLE OUTPUT", "INVERT OUTPUT", "OPEN TAG", _
					"CLOSE TAG", "EVENT MESSAGES", "ALARM AREA(S)", "SECURITY AREA 1", "SECURITY AREA 2", "SECURITY AREA 3", "ALARM AREA 1", "ALARM AREA 2", "ALARM AREA 3", "ALARM AREA 4", _
					"ALARM AREA 5", "ALARM AREA 6", "ALARM AREA 7", "ALARM AREA 8", "ALARM AREA 9", "ALARM AREA 10", "ALARM AREA 11", "ALARM AREA 12", "ALARM AREA 13", "ALARM AREA 14", _
					"ALARM AREA 15", "USER FIELD 1", "USER FIELD 2", "ESIG TYPE", "ESIG ALLOW CONT USE", "ESIG XMPT ALARM ACK", "ESIG UNSIGNED WRITES", "PDR Update Rate", "PDR Access Time", "PDR Deadband", _
					"PDR Latch", "PDR Disable Output", "PDR Array Length", "Hist Description", "Hist Collect", "Hist Interval", "Hist Offset", "Hist Time Res", "Hist Compress", "Hist Deadband", _
					"Hist Comp Type", "Hist Comp Time")
			ElseIf headerNumber = 2 Then
				GetBlockHeaderFieldArray = Array("A_NAME", "A_TAG", "A_DESC", "A_IODV", "A_IOHT", "A_NUMS", "A_IOAD", "A_OUT", "A_INV", "A_OPENDESC", _
					"A_CLOSEDESC", "A_EVENT", "A_ADI", "A_SA1", "A_SA2", "A_SA3", "A_AREA1", "A_AREA2", "A_AREA3", "A_AREA4", _
					"A_AREA5", "A_AREA6", "A_AREA7", "A_AREA8", "A_AREA9", "A_AREA10", "A_AREA11", "A_AREA12", "A_AREA13", "A_AREA14", _
					"A_AREA15", "A_ALMEXT1", "A_ALMEXT2", "A_ESIGTYPE", "A_ESIGCONT", "A_ESIGACK", "A_ESIGTRAP", "A_PDR_UPDATERATE", "A_PDR_ACCESSTIME", "A_PDR_DEADBAND", _
					"A_PDR_LATCHDATA", "A_PDR_DISABLEOUT", "A_PDR_ARRAYLENGTH", "A_HIST_DESC", "A_HIST_COLLECT", "A_HIST_INTERVAL", "A_HIST_OFFSET", "A_HIST_TIMERES", "A_HIST_COMPRESS", "A_HIST_DEADBAND", _
					"A_HIST_COMPTYPE", "A_HIST_COMPTIME")
			Else
				GetBlockHeaderFieldArray = Empty
			End If
		Case "MDI"
			If headerNumber = 1 Then
				GetBlockHeaderFieldArray = Array("BLOCK TYPE", "TAG", "NEXT BLK", "DESCRIPTION", "ENABLE DEVICE 1", "ENABLE DEVICE 2", "DEVICE0", "OPTIONS0", "ADDRESS0", "DEVICE1", _
					"OPTIONS1", "ADDRESS1", "DEVICE2", "OPTIONS2", "ADDRESS2", "INITIAL SCAN", "SCAN TIME", "INITIAL A/M STATUS", "REALARM", "LABEL0", _
					"LABEL1", "LABEL2", "LABEL3", "LABEL4", "LABEL5", "LABEL6", "LABEL7", "ALM0", "ALM1", "ALM2", _
					"ALM3", "ALM4", "ALM5", "ALM6", "ALM7", "ALARM ENABLE", "ALARM AREA(S)", "ALARM PRIORITY", "SECURITY AREA 1", "SECURITY AREA 2", _
					"SECURITY AREA 3", "ALARM AREA 1", "ALARM AREA 2", "ALARM AREA 3", "ALARM AREA 4", "ALARM AREA 5", "ALARM AREA 6", "ALARM AREA 7", "ALARM AREA 8", "ALARM AREA 9", _
					"ALARM AREA 10", "ALARM AREA 11", "ALARM AREA 12", "ALARM AREA 13", "ALARM AREA 14", "ALARM AREA 15", "USER FIELD 1", "USER FIELD 2", "ESIG TYPE", "ESIG ALLOW CONT USE", _
					"ESIG XMPT ALARM ACK", "ESIG TRAP LEVEL", "PDR Update Rate", "PDR Access Time", "PDR Deadband", "PDR Latch", "PDR Disable Output", "PDR Array Length", "Hist Description", "Hist Collect", _
					"Hist Interval", "Hist Offset", "Hist Time Res", "Hist Compress", "Hist Deadband", "Hist Comp Type", "Hist Comp Time")
			ElseIf headerNumber = 2 Then
				GetBlockHeaderFieldArray = Array("A_NAME", "A_TAG", "A_NEXT", "A_DESC", "A_IO1", "A_IO2", "A_DIDV0", "A_DIHT0", "A_DIAD0", "A_DIDV1", _
					"A_DIHT1", "A_DIAD1", "A_DIDV2", "A_DIHT2", "A_DIAD2", "A_ISCAN", "A_SCANT", "A_IAM", "A_REALM", "A_TS0", _
					"A_TS1", "A_TS2", "A_TS3", "A_TS4", "A_TS5", "A_TS6", "A_TS7", "A_AE0", "A_AE1", "A_AE2", _
					"A_AE3", "A_AE4", "A_AE5", "A_AE6", "A_AE7", "A_IENAB", "A_ADI", "A_PRI", "A_SA1", "A_SA2", _
					"A_SA3", "A_AREA1", "A_AREA2", "A_AREA3", "A_AREA4", "A_AREA5", "A_AREA6", "A_AREA7", "A_AREA8", "A_AREA9", _
					"A_AREA10", "A_AREA11", "A_AREA12", "A_AREA13", "A_AREA14", "A_AREA15", "A_ALMEXT1", "A_ALMEXT2", "A_ESIGTYPE", "A_ESIGCONT", _
					"A_ESIGACK", "A_ESIGTRAP", "A_PDR_UPDATERATE", "A_PDR_ACCESSTIME", "A_PDR_DEADBAND", "A_PDR_LATCHDATA", "A_PDR_DISABLEOUT", "A_PDR_ARRAYLENGTH", "A_HIST_DESC", "A_HIST_COLLECT", _
					"A_HIST_INTERVAL", "A_HIST_OFFSET", "A_HIST_TIMERES", "A_HIST_COMPRESS", "A_HIST_DEADBAND", "A_HIST_COMPTYPE", "A_HIST_COMPTIME")
			Else
				GetBlockHeaderFieldArray = Empty
			End If
		Case "AA"
			If headerNumber = 1 Then
				GetBlockHeaderFieldArray = Array("BLOCK TYPE", "TAG", "NEXT BLOCK", "DESCRIPTION", "INITIAL SCAN", "SCAN TIME", "SMOOTHING", "I/O DEVICE", "H/W OPTIONS", "I/O ADDRESS", _
					"SIGNAL CONDITIONING", "LOW EGU LIMIT", "HIGH EGU LIMIT", "DEAD BAND", "EGU TAG", "ALARM ENABLE", "INITIAL A/M STATUS", "REMOTE ACK TAG", "ALARM SUSPENSION TAG", "TARGET VALUE", _
					"ALARM AREA(S)", "ENABLE OUTPUT", "SECURITY AREA 1", "SECURITY AREA 2", "SECURITY AREA 3", "OTHER ALARM PRIORITY", "OTHER CC-TAG", "OTHER CC-MODE", "OTHER DELAY TIME", "OTHER RE-ALARM TIME", _
					"HIHI ALARM LIMIT", "HIHI ALARM PRIORITY", "HIHI CC-TAG", "HIHI CC-MODE", "HIHI DELAY TIME", "HIHI RE-ALARM TIME", "LOLO ALARM LIMIT", "LOLO ALARM PRI RITY.", "LOLO CC-TAG", "LOLO CC-MODE", _
					"LOLO DELAY TIME", "LOLO RE-ALARM TIME", "HI ALARM LIMIT", "HI ALARM PRIORITY", "HI CC-TAG", "HI CC-MODE", "HI DELAY TIME", "HI RE-ALARM TIME", "LO ALARM LIMIT", "LO ALARM PRIORITY", _
					"LO CC-TAG", "LO CC-MODE", "LO DELAY TIME", "LO RE-ALARM TIME", "ROC ALARM LIMIT", "ROC ALARM PRIORITY", "ROC CC-TAG", "ROC CC-MODE", "ROC DELAY TIME", "ROC RE-ALARM TIME", _
					"DEV ALARM LIMIT", "DEV ALARM PRIORITY", "DEV CC-TAG", "DEV CC-MODE", "DEV DELAY TIME", "DEV RE-ALARM TIME", "ALARM AREA 1", "ALARM AREA 2", "ALARM AREA 3", "ALARM AREA 4", _
					"ALARM AREA 5", "ALARM AREA 6", "ALARM AREA 7", "ALARM AREA 8", "ALARM AREA 9", "ALARM AREA 10", "ALARM AREA 11", "ALARM AREA 12", "ALARM AREA 13", "ALARM AREA 14", _
					"ALARM AREA 15", "USER FIELD 1", "USER FIELD 2", "ESIG TYPE", "ESIG ALLOW CONT USE", "ESIG XMPT ALARM ACK", "ESIG UNSIGNED WRITES", "PDR Update Rate", "PDR Access Time", "PDR Deadband", _
					"PDR Latch", "PDR Disable Output", "PDR Array Length", "Hist Description", "Hist Collect", "Hist Interval", "Hist Offset", "Hist Time Res", "Hist Compress", "Hist Deadband", _
					"Hist Comp Type", "Hist Comp Time", "Scale Enabled", "Scale Clamping", "Scale Use EGU", "Scale Raw Low", "Scale Raw High", "Scale Low", "Scale High", "Continuous Contact Output", _
					"Suppress COMM Alarm")
			ElseIf headerNumber = 2 Then
				GetBlockHeaderFieldArray = Array("A_NAME", "A_TAG", "A_NEXT", "A_DESC", "A_ISCAN", "A_SCANT", "A_SMOTH", "A_IODV", "A_IOHT", "A_IOAD", _
					"A_IOSC", "A_ELO", "A_EHI", "A_DBAND", "A_EGUDESC", "A_IENAB", "A_IAM", "A_AACK", "A_ASUSP", "A_TRTAG", _
					"A_ADI", "A_EOUT", "A_SA1", "A_SA2", "A_SA3", "A_OTPRI", "A_CCOTH", "A_MDEOT", "A_DTOTH", "A_RATOT", _
					"A_HIHI", "A_HHPRI", "A_CCHH", "A_MDEHH", "A_DTHH", "A_RATHH", "A_LOLO", "A_LLPRI", "A_CCLL", "A_MDELL", _
					"A_DTLL", "A_RATLL", "A_HI", "A_HIPRI", "A_CCHI", "A_MDEHI", "A_DTHI", "A_RATHI", "A_LO", "A_LOPRI", _
					"A_CCLO", "A_MDELO", "A_DTLO", "A_RATLO", "A_ROC", "A_RCPRI", "A_CCROC", "A_MDERC", "A_DTROC", "A_RATRC", _
					"A_DEV", "A_DVPRI", "A_CCDEV", "A_MDEDV", "A_DTDEV", "A_RATDV", "A_AREA1", "A_AREA2", "A_AREA3", "A_AREA4", _
					"A_AREA5", "A_AREA6", "A_AREA7", "A_AREA8", "A_AREA9", "A_AREA10", "A_AREA11", "A_AREA12", "A_AREA13", "A_AREA14", _
					"A_AREA15", "A_ALMEXT1", "A_ALMEXT2", "A_ESIGTYPE", "A_ESIGCONT", "A_ESIGACK", "A_ESIGTRAP", "A_PDR_UPDATERATE", "A_PDR_ACCESSTIME", "A_PDR_DEADBAND", _
					"A_PDR_LATCHDATA", "A_PDR_DISABLEOUT", "A_PDR_ARRAYLENGTH", "A_HIST_DESC", "A_HIST_COLLECT", "A_HIST_INTERVAL", "A_HIST_OFFSET", "A_HIST_TIMERES", "A_HIST_COMPRESS", "A_HIST_DEADBAND", _
					"A_HIST_COMPTYPE", "A_HIST_COMPTIME", "A_SCALE_ENABLED", "A_SCALE_CLAMP", "A_SCALE_USEEGU", "A_SCALE_RAWLOW", "A_SCALE_RAWHIGH", "A_SCALE_LOW", "A_SCALE_HIGH", "A_CTK_PERSIST", _
					"A_COMM_SURPRESS")
			Else
				GetBlockHeaderFieldArray = Empty
			End If
		Case "AR"
			If headerNumber = 1 Then
				GetBlockHeaderFieldArray = Array("BLOCK TYPE", "TAG", "DESCRIPTION", "I/O DEVICE", "H/W OPTIONS", "I/O ADDRESS TYPE", "I/O ADDRESS", "SIGNAL CONDITIONING", "LOW EGU LIMIT", "HIGH EGU LIMIT", _
					"EGU TAG", "OUTPUT ENABLE", "EVENT MESSAGES", "ALARM AREA(S)", "SECURITY AREA 1", "SECURITY AREA 2", "SECURITY AREA 3", "ALARM AREA 1", "ALARM AREA 2", "ALARM AREA 3", _
					"ALARM AREA 4", "ALARM AREA 5", "ALARM AREA 6", "ALARM AREA 7", "ALARM AREA 8", "ALARM AREA 9", "ALARM AREA 10", "ALARM AREA 11", "ALARM AREA 12", "ALARM AREA 13", _
					"ALARM AREA 14", "ALARM AREA 15", "USER FIELD 1", "USER FIELD 2", "ESIG TYPE", "ESIG ALLOW CONT USE", "ESIG XMPT ALARM ACK", "ESIG UNSIGNED WRITES", "PDR Update Rate", "PDR Access Time", _
					"PDR Deadband", "PDR Latch", "PDR Disable Output", "PDR Array Length", "Hist Description", "Hist Collect", "Hist Interval", "Hist Offset", "Hist Time Res", "Hist Compress", _
					"Hist Deadband", "Hist Comp Type", "Hist Comp Time", "Scale Enabled", "Scale Clamping", "Scale Use EGU", "Scale Raw Low", "Scale Raw High", "Scale Low", "Scale High")
			ElseIf headerNumber = 2 Then
				GetBlockHeaderFieldArray = Array("A_NAME", "A_TAG", "A_DESC", "A_IODV", "A_IOHT", "A_NUMS", "A_IOAD", "A_IOSC", "A_ELO", "A_EHI", _
					"A_EGUDESC", "A_OUT", "A_EVENT", "A_ADI", "A_SA1", "A_SA2", "A_SA3", "A_AREA1", "A_AREA2", "A_AREA3", _
					"A_AREA4", "A_AREA5", "A_AREA6", "A_AREA7", "A_AREA8", "A_AREA9", "A_AREA10", "A_AREA11", "A_AREA12", "A_AREA13", _
					"A_AREA14", "A_AREA15", "A_ALMEXT1", "A_ALMEXT2", "A_ESIGTYPE", "A_ESIGCONT", "A_ESIGACK", "A_ESIGTRAP", "A_PDR_UPDATERATE", "A_PDR_ACCESSTIME", _
					"A_PDR_DEADBAND", "A_PDR_LATCHDATA", "A_PDR_DISABLEOUT", "A_PDR_ARRAYLENGTH", "A_HIST_DESC", "A_HIST_COLLECT", "A_HIST_INTERVAL", "A_HIST_OFFSET", "A_HIST_TIMERES", "A_HIST_COMPRESS", _
					"A_HIST_DEADBAND", "A_HIST_COMPTYPE", "A_HIST_COMPTIME", "A_SCALE_ENABLED", "A_SCALE_CLAMP", "A_SCALE_USEEGU", "A_SCALE_RAWLOW", "A_SCALE_RAWHIGH", "A_SCALE_LOW", "A_SCALE_HIGH")
			Else
				GetBlockHeaderFieldArray = Empty
			End If
		Case Else
			GetBlockHeaderFieldArray = Empty
	End Select
End Function

Private Function GetBlockDataTemplateArray(ByVal blockType As String) As Variant
	Select Case UCase$(Trim$(blockType))
		Case "DA"
			GetBlockDataTemplateArray = Array("DA", "<TAGNAME>", "", "<DESCRIPTION>", "MBE", "", "<ADDRESS>", "AUTO", "ON", "1", _
				"NO", "<OFF CODE>", "<ON CODE>", "ENABLE", "NONE", "LOW", "COS", "NO", "NONE", "NONE", _
				"NONE", "ALL", "", "", "", "", "", "", "", "", _
				"", "", "", "", "", "", "", "", "", "", _
				"", "", "00:00:00:00", "00:00:00:00", "NONE", "YES", "NO", "REJECT", "1,000", "300,000", _
				"0", "NO", "NO", "0", "", "NO", "5,000.0", "0.0", "Milliseconds", "DISABLE", _
				"0.0000", "Absolute", "0", "DISABLE", "DISABLE")
		Case "DR"
			GetBlockDataTemplateArray = Array("DR", "<TAGNAME>", "<DESCRIPTION>", "MBE", "", "DECIMAL", "<ADDRESS>", "NO", "NO", "<OFF CODE>", _
				"<ON CODE>", "DISABLE", "NONE", "NONE", "NONE", "NONE", "ALL", "", "", "", _
				"", "", "", "", "", "", "", "", "", "", _
				"", "", "", "NONE", "YES", "NO", "REJECT", "1,000", "300,000", "0", _
				"NO", "NO", "0", "", "NO", "5,000.0", "0.0", "Milliseconds", "DISABLE", "0.0000", _
				"Absolute", "0")
		Case "MDI"
			GetBlockDataTemplateArray = Array("MDI", "<TAGNAME>", "", "<DESCRIPTION>", "NO", "NO", "MBE", "", "0:0", "", _
				"", "", "", "", "", "ON", "1", "AUTO", "NO", "", _
				"", "", "", "", "", "", "", "NO", "NO", "NO", _
				"NO", "NO", "NO", "NO", "NO", "ENABLE", "NONE", "LOW", "NONE", "NONE", _
				"NONE", "ALL", "", "", "", "", "", "", "", "", _
				"", "", "", "", "", "", "", "", "NONE", "YES", _
				"NO", "REJECT", "1,000", "300,000", "0", "NO", "NO", "0", "", "NO", _
				"5,000.0", "0.0", "Milliseconds", "DISABLE", "0.0000", "Absolute", "0")
		Case "AA"
			GetBlockDataTemplateArray = Array("AA", "<TAGNAME>", "", "<DESCRIPTION>", "ON", "2", "0", "MBE", "<DATA TYPE>", "<ADDRESS>", _
				"None", "<Sensor Lo Limit (Eng Units)>", "<Sensor HI Limit (Eng Units)>", "0.00", "<UNITS>", "ENABLE", "AUTO", "", "", "", _
				"NONE", "NO", "NONE", "NONE", "NONE", "LOW", "", "", "00:00:00:00", "00:00:00:00", _
				"85.00", "LOW", "", "", "00:00:00:00", "00:00:00:00", "15.00", "LOW", "", "", _
				"00:00:00:00", "00:00:00:00", "75.00", "LOW", "", "", "00:00:00:00", "00:00:00:00", "25.00", "LOW", _
				"", "", "00:00:00:00", "00:00:00:00", "", "LOW", "", "", "", "", _
				"", "LOW", "", "", "", "", "ALL", "", "", "", _
				"", "", "", "", "", "", "", "", "", "", _
				"", "", "", "NONE", "YES", "NO", "REJECT", "1,000", "300,000", "0", _
				"NO", "NO", "0", "", "NO", "5,000.0", "0.0", "Milliseconds", "DISABLE", "0.0000", _
				"Absolute", "0", "NO", "NO", "YES", "0.00", "65,535.00", "0.00", "100.00", "DISABLE", _
				"DISABLE")
		Case "AR"
			GetBlockDataTemplateArray = Array("AR", "<TAGNAME>", "<DESCRIPTION>", "SIM", "<DATA TYPE>", "DECIMAL", "<ADDRESS>", "None", "<Sensor Lo Limit (Eng Units)>", "<Sensor HI Limit (Eng Units)>", _
				"<UNITS>", "NO", "DISABLE", "NONE", "NONE", "NONE", "NONE", "ALL", "", "", _
				"", "", "", "", "", "", "", "", "", "", _
				"", "", "", "", "NONE", "YES", "NO", "REJECT", "1,000", "300,000", _
				"0", "NO", "NO", "0", "", "NO", "5,000.0", "0.0", "Milliseconds", "DISABLE", _
				"0.0000", "Absolute", "0", "NO", "NO", "YES", "0.00", "65,535.00", "0.00", "100.00")
		Case Else
			GetBlockDataTemplateArray = Empty
	End Select
End Function
