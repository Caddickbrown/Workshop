import time
import os
import winsound

def clear():
    os.system('cls' if os.name == 'nt' else 'clear')

def pomodoro():
    clear()
    print("Pomodoro Timer")
    total_cycles = int(input("Enter the number of Pomodoros you want to complete: "))

    for cycle in range(total_cycles):
        print(f"\nCycle {cycle+1} - Work")
        work_time = 25 * 60  # Converts minutes to seconds
        countdown(work_time)

        print(f"\nCycle {cycle+1} - Break")
        break_time = 5 * 60   # Converts minutes to seconds
        countdown(break_time)

def countdown(t):
    while t >= 0:
        mins, secs = divmod(t, 60)
        timer = f"{mins:02d}:{secs:02d}"

        if mins == 0 and secs == 0:
            print("Time's up!", end="")
            winsound.Beep(1000, 1000)   # Play a one-second beep sound
            break
        
        if mins == 0 and secs < 4:
            print("Bleeping", end="")

        print(timer, end="\r")
        
        time.sleep(1)
        t -= 1

if __name__ == "__main__":
    pomodoro()