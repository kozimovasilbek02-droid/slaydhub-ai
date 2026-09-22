import sys
sys.stdout.reconfigure(encoding='utf-8')

s = "Ilmiy tadqiqot o'ziga xosligi. Tadqiqot jarayoni mantiqi"
bugged = s.replace('', "'")
recovered = bugged[1:-1:2]
print('Original :', repr(s))
print('Bugged   :', repr(bugged))
print('Recovered:', repr(recovered))
print('Matches  :', recovered == s)
