from multiprocessing import Pipe, Process


def user_input():
    print(">enter anything")

        x = raw_input()
        print(x)
        break


if __name__ == "__main__":
    p1 = Process(target=user_input)
    p1.start()
    p1.join()
