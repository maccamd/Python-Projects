Driver Analysis project documentation:

This python project reads Excel test reports from the RDEC faciltity and plots the speed time plot over every phase of a WLTC.
-------Future Development-------
- Save file in location with name(currently saves to root of file but with test as id no overwrites can only happen if its the same test)
- Look for errors/violations(maybe? got error bands so can see by looking) just showing
- All tests
- Make finished box pop up top when done

------Current Issues------

-----Test Info-----
----Test Paths----
Q:/CORR Project/Emissions Data/Live Raw Data/rdec_emissions/_veh_bt70bhv/CORR BT70 BHV  083  04082021 201446_PM Void.xlsx"
time = 'B'
target_speed = 'CT'
actual_speed = 'CS'
rows start at 3

C:/Users/M0082668/Documents/Python Projects/Driver Analysis/Robot driver trace.xlsx
time = 'A'
target_speed = 'B'
actual_speed = 'C'
rows start at 2

---Test Extra Info---
List element is excel row - 2 so variable[500] is row 502 in xl
Terminal print test section is at the bottom of the code

--Run Time Info--
Path determined, Columns known - 1min17secs
Path determined, Target find actual known - 1min10secs
Path determined, target and actual find - 1min18secs
Path Chosen through TKinter, target and actual find - 1min28secs (picking file takes included)

-Potential Issues-
Finsh box doesnt appear on top will be last window (it does sometimes and wont other times)

https://stackoverflow.com/questions/177287/alert-boxes-in-python
https://pypi.org/project/pywin32/