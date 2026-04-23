
Change "Manual Position/Speed Setpoint" to change to "Auto" when in Auto (VFD and Mod)

For OOS message, make an HMI attribute in the template to store a user defined message to be concatenated with User Name

pw:\\cdmsmith-az02-pw.bentley.com:pw_pl1\Documents\0950\113019\05 Construction Services\17 Contract Close-Out\10 Record Documents\Phase IIIB\As-Built PDFs - Phase IIIB\09 Instrumentation\Final Version\



'========================================================================================================
'========================================================================================================
' This script prepends the object description to all attribute alarm and history comments.
'========================================================================================================
'========================================================================================================

LogMessage("START");

'Variable Declaration
dim attr as string;
dim attr_name as string;
dim ptr_descAttrName as indirect;
dim ptr_description as indirect;


' Iterate all object attributes searching for those with either alarm description (*.DescAttrName) 
' or history description (*.Hist.DescAttrName) properties 
for each attr in Me._Attributes[]
	if StringInString(attr, ".DescAttrName", 1, 0) > 0 then
		attr_name = StringLeft(attr, StringInString(attr, ".", 1, 0) - 1);

		' Prepend the object description to the attribute's alarm/history comment if not already included
		if attr_name <> "" and Me.ShortDesc <> "" then 
			ptr_descAttrName.BindTo("Me." + attr);
			if StringInString(ptr_descAttrName.Value, Me.ShortDesc, 1, 0) == 0 then
				ptr_description.BindTo("Me." + attr_name + ".Description");
				ptr_descAttrName.Value = System.String.Format("{0} {1}", Me.ShortDesc, ptr_description.Value);
			endif;
		endif; 
	endif; 
next;




'========================================================================================================
'========================================================================================================
' Parses bits from the Alm_Interlock and Alm_Trip attributes to construct a string that will be handed to the Interlock/Trip graphic.
'========================================================================================================
'========================================================================================================

'Assign variable to track the bit position in the for loop
dim bitPosition as integer;

'Create booleans to evaluate Trip and Interlock Conditions
dim inInterlock as boolean;
dim inTrip as boolean;

'Label for the relevant bit, determined by progress through the for loop
dim bitPosition_Interlock_Label as string;
dim bitPosition_Trip_Label as string;

'Assign variable(s) to track the target element
dim relevantElementNum as integer;
relevantElementNum = 0;

'Create variable for the summary string
dim summaryString as string;

'Clear the summary strings for new data
Me.Cfg_Interlock_Trip_Summary = "";

for bitPosition = 1 to 32 1;
	inTrip = ((Me.Alm_Trip shr (bitPosition-1)) & 1);

	if inTrip then
		'Set the label for the relevant trip according to the list saved in the attribute
		bitPosition_Trip_Label = Me.Cfg_Trip_Labels[bitPosition];
		
		'Increment the relevant element, since there is a valid assignment
		relevantElementNum = relevantElementNum + 1;

		'Trip_Summary will take the format of "Label#,InInterlock#,InTrip#" in a string
		'Set the label of the nth element
		summaryString = summaryString + bitPosition_Trip_Label + ",0,1,";
	endif;
next;

for bitPosition = 1 to 32 1;
	inInterlock = ((Me.Alm_Interlock shr (bitPosition-1)) & 1);

	if inInterlock then
		'Set the label for the relevant interlock according to the list saved in the attribute
		bitPosition_Interlock_Label = Me.Cfg_Interlock_Labels[bitPosition];
		
		'Increment the relevant element, since there is a valid assignment
		relevantElementNum = relevantElementNum + 1;

		'Trip_Summary will take the format of "Label#,InInterlock#,InTrip#" in a string
		'Set the label of the nth element
		summaryString = summaryString + bitPosition_Interlock_Label + ",1,0,";
	endif;
next;


'Append the number of Elements that the graphic should show
summaryString = summaryString + text(relevantElementNum, "#");

'Export summary string
Me.Cfg_Interlock_Trip_Summary = summaryString;




'========================================================================================================
'========================================================================================================
' Sets the text values for the applied or pending clarifier flows
'========================================================================================================
'========================================================================================================
if (ClarifierNum == 1) then
'	if (PC2_SP_Applied == PC2_SP_Oper) then
'		PC#_Display_Up.Label.Text = "PC#2 Applied SP";
'		'PC#_Display_Up.AnalogVal.Text = text(PC2_SP_Applied, "#### GPM");
'		Upper_PC_Val = PC2_SP_Applied;
'	else
'		PC#_Display_Up.Label.Text = "PC#2 Pending SP";
'		'PC#_Display_Up.AnalogVal.Text = text(PC2_SP_Oper, "#### GPM");
'		Upper_PC_Val = PC2_SP_Oper;
'	endif;
'
'	if (PC3_SP_Applied == PC3_SP_Oper) then
'		PC#_Display_Low.Label.Text = "PC#3 Applied SP";
'		'PC#_Display_Low.AnalogVal.Text = text(PC3_SP_Applied, "#### GPM");
'		Lower_PC_Val = PC3_SP_Applied;
'	else
'		PC#_Display_Low.Label.Text = "PC#3 Pending SP";
'		'PC#_Display_Low.AnalogVal.Text = text(PC3_SP_Oper, "#### GPM");
'		Lower_PC_Val = PC3_SP_Oper;
'	endif;
	PC#_Display_Up.Label.Text = "PC#2 Flow Setpoint";
	Upper_PC_Val = PC2_SP_Applied;
	PC#_Display_Low.Label.Text = "PC#3 Flow Setpoint";
	Lower_PC_Val = PC3_SP_Applied;

elseif (ClarifierNum == 2) then
'	if (PC1_SP_Applied == PC1_SP_Oper) then
'		PC#_Display_Up.Label.Text = "PC#1 Applied SP";
'		'PC#_Display_Up.AnalogVal.Text = text(PC1_SP_Applied, "#### GPM");
'		Upper_PC_Val = PC1_SP_Applied;
'	else
'		PC#_Display_Up.Label.Text = "PC#1 Pending SP";
'		'PC#_Display_Up.AnalogVal.Text = text(PC1_SP_Oper, "#### GPM");
'		Upper_PC_Val = PC1_SP_Oper;
'	endif;
'
'	if (PC3_SP_Applied == PC3_SP_Oper) then
'		PC#_Display_Low.Label.Text = "PC#3 Applied SP";
'		'PC#_Display_Low.AnalogVal.Text = text(PC3_SP_Applied, "#### GPM");
'		Lower_PC_Val = PC3_SP_Applied;
'	else
'		PC#_Display_Low.Label.Text = "PC#3 Pending SP";
'		'PC#_Display_Low.AnalogVal.Text = text(PC3_SP_Oper, "#### GPM");
'		Lower_PC_Val = PC3_SP_Oper;
'	endif;
	PC#_Display_Up.Label.Text = "PC#1 Flow Setpoint";
	Upper_PC_Val = PC1_SP_Applied;
	PC#_Display_Low.Label.Text = "PC#3 Flow Setpoint";
	Lower_PC_Val = PC3_SP_Applied;

elseif (ClarifierNum == 3) then
'	if (PC1_SP_Applied == PC1_SP_Oper) then
'		PC#_Display_Up.Label.Text = "PC#1 Applied SP";
'		'PC#_Display_Up.AnalogVal.Text = text(PC1_SP_Applied, "#### GPM");
'		Upper_PC_Val = PC1_SP_Applied;
'	else
'		PC#_Display_Up.Label.Text = "PC#1 Pending SP";
'		'PC#_Display_Up.AnalogVal.Text = text(PC1_SP_Oper, "#### GPM");
'		Upper_PC_Val = PC1_SP_Oper;
'	endif;
'
'	if (PC2_SP_Applied == PC2_SP_Oper) then
'		PC#_Display_Low.Label.Text = "PC#2 Applied SP";
'		'PC#_Display_Low.AnalogVal.Text = text(PC2_SP_Applied, "#### GPM");
'		Lower_PC_Val = PC2_SP_Applied;
'	else
'		PC#_Display_Low.Label.Text = "PC#2 Pending SP";
'		'PC#_Display_Low.AnalogVal.Text = text(PC2_SP_Oper, "#### GPM");
'		Lower_PC_Val = PC2_SP_Oper;
'	endif;
	PC#_Display_Up.Label.Text = "PC#1 Flow Setpoint";
	Upper_PC_Val = PC1_SP_Applied;
	PC#_Display_Low.Label.Text = "PC#2 Flow Setpoint";
	Lower_PC_Val = PC2_SP_Applied;
endif;




'========================================================================================================
'========================================================================================================
'Map the live Summary String from the template to a temporary version (which also avoids triggering the script)
'========================================================================================================
'========================================================================================================
dim destructibleSummaryString as string;
destructibleSummaryString = CSV_Interlock_Trip_Summary;




'========================================================================================================
'========================================================================================================
'Set the number of elements that should be visible based on the imported string (rightmost string)
'========================================================================================================
'========================================================================================================dim numElements as integer;
numElements = stringtointg(stringright(destructibleSummaryString,1));

if numElements >= 1 then
	'Map the destructible string to an array for repeated data access
	'Variables cannot be used to size the array, so an array of the maximum expected size will be created
	'Maximum expected size is 21, based on 7 elements, each containing a label, Interlock symbol, and Trip symbol
	dim summaryArray[21] as string;
	dim summaryArraySize as integer;
		summaryArraySize = summaryArray[].length;

	dim i as integer;
	dim nextCommaPos as integer;
	dim mapString as string;

	for i = 1 to (summaryArraySize) 1;
		'Find the position of the next delimiter and map the leftmost string in the destructible string to the array
		nextCommaPos = stringinstring(destructibleSummaryString,",",1,0);
		
		if nextCommaPos <> 0 then
			mapString = stringleft(destructibleSummaryString, (nextCommaPos - 1));
		
			'Remove the leftmost string from the destructible string
			destructibleSummaryString = stringright(destructibleSummaryString, stringlen(destructibleSummaryString) - (nextCommaPos));
		
		elseif (i mod(3) == 1) then
			'Every 3i-2 position in the array is populated by a label
			mapString = "Inactive";
		
		else
			'All other positions in the array act as booleans for InInterlock and InTrip conditions
			'To be written to each update even if they don't control transparency due to numElements (optimization possible)
			mapString = "0";
		endif;
		
		'Assign mapString to each position regardless of outcome of IF statements
		summaryArray[i] = mapString;

	next;

	'1 4 7 10 13 16 19 22 X25

	'For i=1,4,7,10,13,16,19,22...
	'Could be optimized if we could dynamically assign the relevant element
	for i = 1 to summaryArraySize 3;
		if i == 1 then
			Status1.Label.Text = summaryArray[i];
			Status1.Interlock.Transparency = (100 * not(stringtointg(summaryArray[i+1])));
			Status1.Trip.Transparency = (100 * not(stringtointg(summaryArray[i+2])));
			if numElements >= ((i+2)/3) then
				Status1.Transparency = 0;
			else
				Status1.Transparency = 100;
			endif;
		endif;
		if i == 4 then
			Status2.Label.Text = summaryArray[i];
			Status2.Interlock.Transparency = (100 * not(stringtointg(summaryArray[i+1])));
			Status2.Trip.Transparency = (100 * not(stringtointg(summaryArray[i+2])));
			if numElements >= ((i+2)/3) then
				Status2.Transparency = 0;
			else
				Status2.Transparency = 100;
			endif;
		endif;
		if i == 7 then
			Status3.Label.Text = summaryArray[i];
			Status3.Interlock.Transparency = (100 * not(stringtointg(summaryArray[i+1])));
			Status3.Trip.Transparency = (100 * not(stringtointg(summaryArray[i+2])));
			if numElements >= ((i+2)/3) then
				Status3.Transparency = 0;
			else
				Status3.Transparency = 100;
			endif;
		endif;
		if i == 10 then
			Status4.Label.Text = summaryArray[i];
			Status4.Interlock.Transparency = (100 * not(stringtointg(summaryArray[i+1])));
			Status4.Trip.Transparency = (100 * not(stringtointg(summaryArray[i+2])));
			if numElements >= ((i+2)/3) then
				Status4.Transparency = 0;
			else
				Status4.Transparency = 100;
			endif;
		endif;
		if i == 13 then
			Status5.Label.Text = summaryArray[i];
			Status5.Interlock.Transparency = (100 * not(stringtointg(summaryArray[i+1])));
			Status5.Trip.Transparency = (100 * not(stringtointg(summaryArray[i+2])));
			if numElements >= ((i+2)/3) then
				Status5.Transparency = 0;
			else
				Status5.Transparency = 100;
			endif;
		endif;
		if i == 16 then
			Status6.Label.Text = summaryArray[i];
			Status6.Interlock.Transparency = (100 * not(stringtointg(summaryArray[i+1])));
			Status6.Trip.Transparency = (100 * not(stringtointg(summaryArray[i+2])));
			if numElements >= ((i+2)/3) then
				Status6.Transparency = 0;
			else
				Status6.Transparency = 100;
			endif;
		endif;
		if i == 19 then
			Status7.Label.Text = summaryArray[i];
			Status7.Interlock.Transparency = (100 * not(stringtointg(summaryArray[i+1])));
			Status7.Trip.Transparency = (100 * not(stringtointg(summaryArray[i+2])));
			if numElements >= ((i+2)/3) then
				Status7.Transparency = 0;
			else
				Status7.Transparency = 100;
			endif;
		endif;
	next;
else
	Status1.Label.Text = "No Active Trips or Interlocks";
	Status1.Interlock.Transparency = 100;
	Status1.Trip.Transparency = 100;
	Status1.Transparency = 0;
	Status2.Transparency = 100;
	Status3.Transparency = 100;
	Status4.Transparency = 100;
	Status5.Transparency = 100;
	Status6.Transparency = 100;
	Status7.Transparency = 100;
endif;




'========================================================================================================
'========================================================================================================
' Alarm setpoint auto-layout
' Minimal change on EU change
' Reset to defaults only if EU change creates overlap/tie
'========================================================================================================
'========================================================================================================
 
'Variable Declaration
dim Low as float;
dim High as float;
dim EURange as float;
dim eps as float;
dim Temp as float;
 
dim pLoLo as float;
dim pLo as float;
dim pHi as float;
dim pHiHi as float;
 
dim needReset as boolean;
 
' Tuneable threshold:
' If (Sp_HiHi - Sp_LoLo) is less than this fraction of new EU range, reset to defaults.
dim MinSpanFrac as float;
MinSpanFrac = 0.40;   '40% (adjust to 0.30 if you want less aggressive)
 
'---------------------------------------
' 1) Normalize EU limits (handle flipped)
'---------------------------------------
Low  = Me.Cfg_EULo;
High = Me.Cfg_EUHi;

if High < Low then
    Temp = High;
    High = Low;
    Low  = Temp;
endif;
 
EURange = High - Low;
 
'If range is not valid, do nothing
if EURange <= 0 then
    'Do nothing
else
 
    '---------------------------------------
    ' 2) Epsilon used only for overlap/tie detection
    '---------------------------------------
    'Epsilon = 0.1% of span, with a small absolute floor
    eps = EURange * 0.001;
    if eps < 0.0005 then
        eps = 0.0005;
    endif;
 
    '---------------------------------------
    ' 3) Start by preserving literal values (proposed = current)
    '---------------------------------------
    pLoLo = Me.Sp_LoLo;
    pLo   = Me.Sp_Lo;
    pHi   = Me.Sp_Hi;
    pHiHi = Me.Sp_HiHi;
 
    needReset = false;
 
    '---------------------------------------
    ' 4) If the CURRENT state is already "stuck" (overlap/tie/out-of-order),
    '    reset to defaults (EU change is a good time to fix it)
    '---------------------------------------
    if (pLoLo >= (pLo - eps)) or (pLo >= (pHi - eps)) or (pHi >= (pHiHi - eps)) then
        needReset = true;
    endif;
 
    '---------------------------------------
    ' 5) Clamp ONLY setpoints that are out of [Low, High]
    '---------------------------------------
    if needReset == false then
 
        if pLoLo < Low then pLoLo = Low; endif;
        if pLo   < Low then pLo   = (Low + (10 * eps)); endif;
 
        if pHi   > High then pHi   = (High - (10 * eps)); endif;
        if pHiHi > High then pHiHi = High; endif;
 
        '---------------------------------------
        ' 5a) Drastic-change trigger:
        '     If the existing SP span is too small vs new EU range, reset to defaults
        '---------------------------------------
        if (pHiHi - pLoLo) < (MinSpanFrac * EURange) then
            needReset = true;
        endif;
 
        '---------------------------------------
        ' 6) If clamping creates an overlap/tie (i.e., would "stick" in GUI),
        '    reset everything to defaults
        '---------------------------------------
        if (pLoLo >= (pLo - eps)) or (pLo >= (pHi - eps)) or (pHi >= (pHiHi - eps)) then
            needReset = true;
        endif;
 
    endif;
 
    '---------------------------------------
    ' 7) Reset to default distribution
    '---------------------------------------
    if needReset then
 
        'Defaults: 23/40/60/78% of range
        pLoLo = Low + (0.23 * EURange);
        pLo   = Low + (0.40 * EURange);
        pHi   = Low + (0.60 * EURange);
        pHiHi = Low + (0.78 * EURange);
 
    endif;
 
    '---------------------------------------
    ' 8) Write results (order irrelevant for script writes)
    '---------------------------------------
    Me.Sp_LoLo = pLoLo;
    Me.Sp_Lo   = pLo;
    Me.Sp_Hi   = pHi;
    Me.Sp_HiHi = pHiHi;
 
endif;
