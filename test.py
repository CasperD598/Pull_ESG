array = []
for x in range(1, 11):
    print(x)
    length = len(array)
    for y in range(1, 6):
        if y == 5:
            array.append(y)
            break
    if len(array) > length:
        break

print(y)