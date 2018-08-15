from threading import Thread
import time


def timer(name, delay, repeat):
    print("Timer: " + name + " started")
    while repeat > 0:
        time.sleep(delay)
        repeat -= 1
        print(name + ": " + time.ctime(time.time()))
    print("Timer: " + name + " complete")


if __name__ == "__main__":
    t1 = Thread(target=timer, args=('THREAD-1', 1, 5))
    t2 = Thread(target=timer, args=('THREAD-2', 2, 3))
    t1.start()
    t2.start()