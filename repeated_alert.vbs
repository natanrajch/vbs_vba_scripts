'This is actually used on vbs

Dim MyMes 
Dim MyPath
Dim MyFile
Dim MyCounter

MyCounter = 0

if (Month(date) = 1) then
	MyMes = "Enero"
else if (Month(date) = 2) Then
	MyMes = "Febrero"
else if (Month(date) = 3) Then
	MyMes = "Marzo"
else if (Month(date) = 4) then
	MyMes = "Abril"
else if (Month(date) = 5) then
	MyMes = "Mayo"
else if (Month(date) = 6) then
	MyMes = "Junio"
else if (Month(date) = 7) then
	MyMes = "Julio"
else if (Month(date) = 8) then
	MyMes = "Agosto"
else if (Month(date) = 9) then
	MyMes = "Septiembre"
else if (Month(date) = 10) then
	MyMes = "Octubre"
else if (Month(date) = 11) then
	MyMes = "Noviembre"
else
	MyMes = "Diciembre"
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if

MyPath = "\\YOUR_PATH\"



Dim fs, drv, fldr, fle
set fs = CreateObject("Scripting.FileSystemObject")
set Fldr = fs.GetFolder(MyPath)

for each fle in Fldr.Files
if instr(1,fle.name,MyMes,vbTextCompare)>0 then
	MyCounter = MyCounter + 1
end if
next

If (MyCounter > 1) then
	msgbox ("Hay más de un archivo con el mes de " & MyMes)
end if

set fs = Nothing
set Fldr = Nothing