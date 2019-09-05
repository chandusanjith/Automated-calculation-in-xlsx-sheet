cd C:\Users\chandu.s\Desktop\CARDRATE\in
ren *.xlsx card.xlsx
cd C:\Users\chandu.s\Desktop\CARDRATE\scripts
tets.vbs
XlsToCsv.vbs C:\Users\chandu.s\Desktop\CARDRATE\in\card.xlsx C:\Users\chandu.s\Desktop\CARDRATE\outfile\cardrate.csv
ftp -s:ftpscript.txt.
move C:\Users\chandu.s\Desktop\CARDRATE\in\card.xlsx C:\Users\chandu.s\Desktop\CARDRATE\convrted_xls\CARDRATE%time:~0,2%%time:~3,2%%time:~6,2%_%date:~-10,2%%date:~-7,2%%date:~-4,4%.xlsx
move C:\Users\chandu.s\Desktop\CARDRATE\outfile\cardrate.csv C:\Users\chandu.s\Desktop\CARDRATE\archieve\cardrate%time:~0,2%%time:~3,2%%time:~6,2%_%date:~-10,2%%date:~-7,2%%date:~-4,4%.csv


