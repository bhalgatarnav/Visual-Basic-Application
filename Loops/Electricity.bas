Dim a, gamma, r, t As Double
t = ActiveSheet.cells(2,3).value
t = (t-32)*(5/9) + (273.15)

a = sqr(t*1.4*287.03)

ActiveSheet.cells(3,3).value = a