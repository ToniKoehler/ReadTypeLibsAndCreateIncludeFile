# Create Headerfile (.h) with constants and enums from typelibs in VFP (Visual Foxpro)
## Description
You can create your own headerfile (.h) with all constants and enums.
This a a better and faster solution instead the old way to drag/drop all constants and enums from Tools | Object Explorer into a .prg file.

## CreateHeaderFileForExcel.prg

```foxpro
LOCAL loTLI AS TLI.TLIApplication
LOCAL loTypeLib AS TLI.TypeLibInfo
LOCAL lcFile, lcOut, loInfo, loMember

*!*  -------------------------------- TONI KOEHLER 07.04.2026 09:08:45 --------------------------------------
*!*  Creates INCLUDE file with all Excel-Enums and Constants
*!*  Requirements: tlbinf32.dll (it's a part of Visual Foxpro 9 Installation)
*!*  -------------------------------------------------------------------------------------------------------

lcFile = "C:\Temp\excel.h"
lcOut  = ""

** tlbinf32.dll must be registered e.g. by regsvr32
loTLI = CREATEOBJECT("TLI.TLIApplication")

** load Excel-TypeLib 
loTypeLib = loTLI.TypeLibInfoFromFile(FULLPATH("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"))

FOR EACH loInfo IN loTypeLib.Constants

	TEXT TO lcOut ADDITIVE TEXTMERGE NOSHOW FLAGS 1 PRETEXT 1 + 2
		[<<loInfo.NAME>>]
		
	ENDTEXT

	FOR EACH loMember IN loInfo.Members
		TEXT TO lcOut ADDITIVE TEXTMERGE NOSHOW FLAGS 1 PRETEXT 1 + 2
			#DEFINE <<loMember.NAME>> <<TRANSFORM(loMember.VALUE)>>
			
		ENDTEXT
	NEXT
NEXT

STRTOFILE(lcOut, lcFile)

? "Finish: " + lcFile

```