import argparse


def fibn(num):
    a, b = 0, 1
    for _ in range(num):
        a, b = b, a + b
    return a


def Main():
    parser = argparse.ArgumentParser()
    parser.add_argument("num", help="the fibonacci number you wish to calculate", type=int)
    parser.add_argument("-o", "--output", help="output result to a file",\
                        action="store_true")
    args = parser.parse_args()

    result = fibn(args.num)
    print("The " + str(args.num) + "th fib number is " + str(result))

    if args.output:
        with open("fibn_result.txt", 'a') as s_file:
            s_file.write(str(result)+"\n")


if __name__ == "__main__":
    Main()