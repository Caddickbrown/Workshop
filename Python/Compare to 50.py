# Add two numbers together and compare it to 50
import time,sys

chip = 5
question = "Enter your {} Number: "

def tPrint(text):
  for character in text:
    sys.stdout.write(character)
    sys.stdout.flush()
    time.sleep(0.02)
  
def tInput(text):
  for character in text:
    sys.stdout.write(character)
    sys.stdout.flush()
    time.sleep(0.02)
  value = input()  
  return value

def errorOut():
    tPrint("How did you get here?")
    time.sleep(3)
    tPrint("\nNo, seriously - what the hell did you do!?")
    time.sleep(5)
    tPrint("\nHow can a number not be equal to, less than, or greater than 50!?")
    time.sleep(5)
    tPrint("\nThat doesn't even make sense....")
    time.sleep(5)
    tPrint("\nWHHHYYYYY!!?!?!?!?!?!?!?!?!??!?!?!")
    time.sleep(5)
    tPrint("\nYeah whatever - move on...")

def sumNumbers(num1,num2):
    sum = num1 + num2

    if sum == 50:
        time.sleep(0.5)
        print("Your answer is",sum,"which is exactly 50.")
        time.sleep(2)
    elif sum > 50:
        time.sleep(0.5)
        print("Your answer is",sum,"which is greater than 50.")
        time.sleep(2)
    elif sum < 50:
        time.sleep(0.5)
        print("Your answer is",sum,"which is less than 50.")
        time.sleep(2)
    else:
	    errorOut()

while (chip == 5):
    num3 = int(tInput(question.format("First")))
    num4 = int(tInput(question.format("Second")))

    sumNumbers(num3,num4)

    tInput("Press enter to start again.")



