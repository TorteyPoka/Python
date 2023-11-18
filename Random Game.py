import random

Name1 = input("Enter name 1: ")
Name2 = input("Enter name 2: ")
Name3 = input("Enter name 3: ")

name1RandomNo = random.randint(0, 20)
name2RandomNo = random.randint(0, 20)
name3RandomNo = random.randint(0, 20)

print(Name1, "Lucky no. ", name1RandomNo)
print(Name2, "Lucky no. ", name2RandomNo)
print(Name3, "Lucky no. ", name3RandomNo)
