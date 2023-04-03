import threading
import time

def get_input():
    return input("You have 5 seconds to enter a value: ")

def set_default():
    global var333
    time.sleep(5)
    if not var333:
        var333 = 0
        print("Time's up! Assigning var333 as 0.")

var333 = None
threading.Thread(target=set_default).start()
var333 = get_input()

if var333:
    print(f"You entered: {var333}")
    var333 = int(var333)
else:
    print(f"var333: {var333}")

time.sleep(10)

import threading
import time

def get_input():
    return input("You have 5 seconds to enter a value: ")

def set_default():
    global var333, time_up
    time.sleep(5)
    if not var333:
        time_up = True

var333 = None
time_up = False
threading.Thread(target=set_default).start()
var333 = get_input()

if not time_up:
    print(f"You entered: {var333}")
    var333 = int(var333)
else:
    var333 = 0
    print("Time's up! Assigning var333 as 0.")
    print(f"var333: {var333}")