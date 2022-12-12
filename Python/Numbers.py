num1 = int(input("Enter your First Number: "))
num2 = int(input("Enter your Second Number: "))
if num1 == num2:
	print("These numbers are the same.")
elif num1 > num2:
	print("Your first number was bigger than your second number.")
elif num1 < num2:
	print("Your second number was bigger than your first number.")
else:
	print("How did you get here?")
input("Press Enter to exit.")