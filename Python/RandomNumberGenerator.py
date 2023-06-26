import random

print("Welcome to Python Random Number Generator")

while (1==1):
    decision = int(input("Press 1 to define min and max, or Press 2 for Numbers between 0 and 100." + "\n"))
    if decision == 1:
        num1 = int(input("\n" + "Enter your Minimum Number: "))
        num2 = int(input("Enter your Maximum Number: "))
        while(1==1):
            input(random.randint(num1,num2))
    elif decision == 2:
        while(1==1):
            input(random.randint(0,100))
    else:
        print( "\n" + "Error - please try again." + "\n")