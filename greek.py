import unicodedata

# Создаем список для греческих символов
greek_letters = []

# Цикл для добавления греческих символов в список
for i in range(0x0370, 0x03ff + 1):
    greek_letters.append(chr(i))

# Цикл для вывода греческих символов
for letter in greek_letters:
    print(letter)