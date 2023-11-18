s = input()
count = 0
for x in s:
    if x.isupper():
        count += 1

if count > len(s) // 2:
    s = s.upper()
else:
    s = s.lower()

print(s)