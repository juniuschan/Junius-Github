import threading
import queue


datas = [[1,2,3], [4,5,6], [7,8,9], [10,11]]


def do_work(alist, q):
    for i in range(len(alist)):
        alist[i] = alist[i]**2
    q.put(alist)


def multi_thread(data):
    q = queue.Queue()
    threads = []
    for i in range(len(data)):
        t = threading.Thread(target=do_work, args=(data[i], q))
        t.start()
        threads.append(t)
    for thread in threads:
        thread.join()
    results = []
    for _ in range(len(data)):
        results.append(q.get())
    print(results)


if __name__ == "__main__":
    multi_thread(datas)