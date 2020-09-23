<div align="center">

## Get Hours, Minutes, AND Seconds between 2 Dates \(or of total seconds\)


</div>

### Description

This code will enable you to take two dates, find the total seconds in between them, and calculate the Hours, Minutes, and Seconds.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sparq](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sparq.md)
**Level**          |Beginner
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sparq-get-hours-minutes-and-seconds-between-2-dates-or-of-total-seconds__1-49249/archive/master.zip)





### Source Code

<BR><BR>
Dim H, M, S, Secs as Long<BR>
&nbsp;Secs = DateDiff("s", Date1, Date2)<BR>
&nbsp;H = Int(Secs / 3600)<BR>
&nbsp;Secs = Secs - (H * 3600)<BR>
&nbsp;M = Int(Secs / 60)<BR>
&nbsp;Secs = Secs - (M * 60)<BR>
&nbsp;S = Secs<BR>
&nbsp;MsgBox Format(H, "00") & ":" & _<BR>
&nbsp;&nbsp; Format(M, "00") & ":" & _<BR>
&nbsp;&nbsp; Format(S, "00")

