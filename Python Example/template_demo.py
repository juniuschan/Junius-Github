from string import Template


def Main():
    char = []

    char.append(dict(item="apple", cost=4.5, qyt=5))
    char.append(dict(item="coke", cost=3.5, qyt=7))
    char.append(dict(item="fish", cost=3.0, qyt=1))

    t = Template("$qyt * $item = $cost")
    total = 0
    for item in char:
        print(t.substitute(item))
        total += item["cost"]
    print("char total:")
    print("All cost " + str(total))


if __name__ == "__main__":
    Main()